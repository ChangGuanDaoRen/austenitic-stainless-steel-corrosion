# -*- coding: utf-8 -*-
"""
v1 基线建模：只验证「00.xlsx 现有数据能不能跑通一条可信的 Epit 模型」。
口径：
  - 目标 Epit(V)：只取数值行，177 行待读图/62 行文本 Epit 全部剔除（不填数）
  - 特征：成分 C/Si/N/P/S/Cu/Mn/Cr/Mo/Ni/Ti/Nb/Co/Sn + PREN(Cr+3.3Mo+16N)
  - 交叉验证：GroupKFold(5) 按文献编号分组，杜绝同篇泄漏
  - 模型：Ridge / RandomForest / GradientBoosting
输出：控制台 + model_v1_20261110.json + model_v1_特征覆盖_20261110.csv
只读 00.xlsx，不改任何源表。
"""
import json
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
from sklearn.ensemble import GradientBoostingRegressor, RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import RidgeCV
from sklearn.model_selection import GroupKFold, cross_val_predict
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import r2_score, mean_absolute_error

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "00.xlsx"
COMP = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni", "Ti", "Nb", "Co", "Sn"]


def num(v):
    """把单元格转成 float；'12.5' / 12.5 都收，其余（含 ± / 区间 / 空 / 待读图）返回 nan。"""
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return np.nan
    if re.fullmatch(r"[-+]?\d*\.?\d+", s):
        return float(s)
    return np.nan


def mid(v):
    """文本 Epit 取中值：'0.32±0.03' -> 0.32；'0.25~0.35' / '0.25-0.35' -> 0.30；'约0.2' -> 0.2。"""
    if v is None:
        return np.nan
    s = str(v).strip().replace("～", "~").replace("—", "-").replace("–", "-")
    s = s.replace("约", "").replace(" ", "")
    m = re.match(r"^([-+]?\d*\.?\d+)[±\+\-]\s*\d*\.?\d+$", s)
    if m:
        return float(m.group(1))
    m = re.match(r"^([-+]?\d*\.?\d+)~\s*([-+]?\d*\.?\d+)$", s) or \
        re.match(r"^([-+]?\d*\.?\d+)-([-+]?\d*\.?\d+)$", s)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        return (a + b) / 2
    m = re.fullmatch(r"[-+]?\d*\.?\d+", s)
    if m:
        return float(s)
    return np.nan


def load():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    hdr_comp = [rows[0][i] for i in range(2, 17)]
    assert hdr_comp[:1] == ["C"], hdr_comp
    recs = []
    last_lit = None
    for r in rows[1:]:
        # 文献编号只在数据块首行写一次，需向下填充
        if r[0] is not None:
            last_lit = r[0]
        if r[1] is None:          # 没有样本名的行视为空行
            continue
        d = {
            "文献编号": last_lit,
            "样本": r[1],
            "工艺": r[17],
            "参比": r[19],
            "温度": r[21],
            "Epit_raw": r[22],
        }
        for j, name in enumerate(hdr_comp):
            d[name] = num(r[2 + j])
        recs.append(d)
    df = pd.DataFrame(recs)
    df["Epit"] = df["Epit_raw"].map(num)
    return df


def pren(row):
    cr, mo, n = row["Cr"], row["Mo"], row["N"]
    if pd.isna(cr):
        return np.nan
    return cr + 3.3 * (0 if pd.isna(mo) else mo) + 16 * (0 if pd.isna(n) else n)


def run(df, tag):
    feats = COMP + ["PREN"]
    X = df[feats].copy()
    y = df["Epit"].values
    g = df["文献编号"].astype(str).values
    n_splits = min(5, len(np.unique(g)))
    cv = GroupKFold(n_splits=n_splits)
    out = {"口径": tag, "行数": int(len(df)), "篇数": int(df["文献编号"].nunique())}
    models = {
        "Ridge": Pipeline([
            ("imp", SimpleImputer(strategy="median", add_indicator=True)),
            ("sc", StandardScaler()),
            ("m", RidgeCV(alphas=np.logspace(-2, 3, 30))),
        ]),
        "RF": Pipeline([
            ("imp", SimpleImputer(strategy="median", add_indicator=True)),
            ("m", RandomForestRegressor(n_estimators=500, min_samples_leaf=2,
                                        random_state=0, n_jobs=-1)),
        ]),
        "GBR": Pipeline([
            ("imp", SimpleImputer(strategy="median", add_indicator=True)),
            ("m", GradientBoostingRegressor(n_estimators=400, learning_rate=0.05,
                                            max_depth=3, random_state=0)),
        ]),
    }
    for name, pipe in models.items():
        pred = cross_val_predict(pipe, X, y, groups=g, cv=cv, n_jobs=1)
        out[name] = {
            "R2": round(float(r2_score(y, pred)), 4),
            "MAE": round(float(mean_absolute_error(y, pred)), 4),
            "RMSE": round(float(np.sqrt(np.mean((y - pred) ** 2))), 4),
        }
    # 单变量基线：只用 PREN 的一元线性
    sub = df.dropna(subset=["PREN"])
    if len(sub) > 20:
        p = np.polyfit(sub["PREN"], sub["Epit"], 1)
        pred = np.polyval(p, sub["PREN"])
        out["PREN_only_polyfit"] = {
            "n": int(len(sub)),
            "R2_in_sample": round(float(r2_score(sub["Epit"], pred)), 4),
            "slope": round(float(p[0]), 5),
        }
    return out


def main():
    df_all = load()
    numeric = df_all.dropna(subset=["Epit"]).copy()
    numeric["PREN"] = numeric.apply(pren, axis=1)

    # 覆盖度统计
    cov = []
    for c in COMP + ["PREN", "工艺", "温度", "参比"]:
        if c in ("工艺", "温度", "参比"):
            nn = numeric[c].notna().sum()
        else:
            nn = numeric[c].notna().sum()
        cov.append({"字段": c, "非空行数": int(nn),
                    "覆盖率%": round(100 * nn / len(numeric), 1)})
    cov = pd.DataFrame(cov)

    results = []
    results.append(run(numeric, "A 全部数值 Epit（去 177 待读图 + 62 文本）"))

    trim = numeric[(numeric["Epit"] >= -0.3) & (numeric["Epit"] <= 1.2)].copy()
    results.append(run(trim, "B 再剔极值（-0.3 ~ 1.2 V）"))

    has_pren = numeric.dropna(subset=["PREN"]).copy()
    results.append(run(has_pren, "C 再要求 Cr 非空（PREN 可算）"))

    print("=" * 78)
    print("数据分层")
    print("=" * 78)
    print(f"00.xlsx 总行数           : {len(df_all)}")
    print(f"Epit 数值行（建模可用）  : {len(numeric)}   篇数 {numeric['文献编号'].nunique()}")
    print(f"Epit 非数值行（剔除）    : {len(df_all) - len(numeric)}")
    print(f"其中 待读图/文本/空      : {df_all['Epit'].isna().sum()}")
    print()
    print("特征覆盖度（基于数值行）")
    print(cov.to_string(index=False))
    print()
    print("=" * 78)
    print("GroupKFold(5) 按文献分组的交叉验证结果")
    print("=" * 78)
    for r in results:
        print(f"\n[{r['口径']}]  n={r['行数']}  篇数={r['篇数']}")
        print(f"  {'模型':<24}{'R2':>8}{'MAE(V)':>10}{'RMSE(V)':>10}")
        for k in ("Ridge", "RF", "GBR"):
            if k in r:
                print(f"  {k:<24}{r[k]['R2']:>8.3f}{r[k]['MAE']:>10.3f}{r[k]['RMSE']:>10.3f}")
        if "PREN_only_polyfit" in r:
            b = r["PREN_only_polyfit"]
            print(f"  {'仅 PREN 一元线性':<22}{b['R2_in_sample']:>8.3f}"
                  f"   n={b['n']} 斜率={b['slope']}")

    (ROOT / "model_v1_20261110.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    cov.to_csv(ROOT / "model_v1_特征覆盖_20261110.csv", index=False, encoding="utf-8-sig")
    numeric.to_csv(ROOT / "model_v1_建模集_20261110.csv", index=False, encoding="utf-8-sig")
    print("\n已写出: model_v1_20261110.json / model_v1_特征覆盖_20261110.csv / model_v1_建模集_20261110.csv")

    # ---------------- 诊断：低 R2 到底卡在哪 ----------------
    print("\n" + "=" * 78)
    print("诊断 1  文献间系统性偏移（方差分解，基于 500 数值行）")
    print("=" * 78)
    gm = numeric["Epit"].mean()
    grp = numeric.groupby("文献编号")["Epit"]
    sizes = grp.size()
    means = grp.mean()
    within = ((numeric["Epit"] - numeric["文献编号"].map(means)) ** 2).sum()
    between = (sizes * (means - gm) ** 2).sum()
    total = ((numeric["Epit"] - gm) ** 2).sum()
    print(f"总方差分解：文献间 {100*between/total:.1f}%  文献内 {100*within/total:.1f}%")
    print(f"各篇均值极差：{means.max():.3f} V（文献{means.idxmax()}）"
          f" ←→ {means.min():.3f} V（文献{means.idxmin()}），跨度 {means.max()-means.min():.3f} V")
    multi = sizes[sizes >= 3].index
    sd_in = numeric[numeric["文献编号"].isin(multi)].groupby("文献编号")["Epit"].std()
    print(f"篇内标准差（≥3 行的 {len(multi)} 篇）：中位数 {sd_in.median():.3f} V，最大 {sd_in.max():.3f} V")

    print("\n" + "=" * 78)
    print("诊断 2  参比电极差异（同一 PREN 区间内的组均值，抵消合金成分影响）")
    print("=" * 78)
    d = numeric.dropna(subset=["PREN"]).copy()
    d["PREN箱"] = pd.cut(d["PREN"], bins=[0, 20, 24, 28, 32, 40, 100])
    d["RE组"] = d["参比"].astype(str).str.strip().str.replace(r"\s+", "", regex=True)
    d["RE组"] = d["RE组"].where(d["RE组"].isin(["SCE", "Ag/AgCl(3MKCl)", "Ag/AgCl"]), "其它/中文写法")
    tab = d.pivot_table(index="PREN箱", columns="RE组", values="Epit",
                        aggfunc=["mean", "size"], observed=False)
    print(tab.round(3).to_string())
    print("\n参比取值分布（原文写法）：")
    print(f"  SCE {int((d['RE组']=='SCE').sum())} 行 / Ag/AgCl(3M KCl) "
          f"{int((d['RE组']=='Ag/AgCl(3MKCl)').sum())} 行 / 其它或中文写法 "
          f"{int((d['RE组']=='其它/中文写法').sum())} 行")

    print("\n" + "=" * 78)
    print("诊断 3  加「参比 + 温度」特征后 R2 变化（说明缺的是批次校正，不是成分）")
    print("=" * 78)
    d2 = numeric.dropna(subset=["PREN"]).copy()
    d2["RE"] = d2["参比"].astype(str).str.contains("Ag", case=False).astype(int)
    d2["T"] = pd.to_numeric(d2["温度"].astype(str).str.extract(r"(-?\d+\.?\d*)")[0],
                            errors="coerce")
    base = run(d2[["Epit"] + COMP + ["PREN"]].assign(文献编号=d2["文献编号"]), "仅成分+PREN")
    plus = run(d2.assign(**{"RE_hot": d2["RE"]})[["Epit"] + COMP + ["PREN", "RE_hot"]
                                                   ].assign(文献编号=d2["文献编号"]), "成分+PREN+参比哑变量")
    print(f"  RF  仅成分 : R2={base['RF']['R2']:.3f}  MAE={base['RF']['MAE']:.3f}")
    print(f"  RF  +参比 : R2={plus['RF']['R2']:.3f}  MAE={plus['RF']['MAE']:.3f}")

    print("\n" + "=" * 78)
    print("诊断 5  篇内去均值（消除文献偏移）后，成分还能否解释 Epit")
    print("=" * 78)
    cen = numeric.dropna(subset=["PREN"]).copy()
    for c in COMP + ["PREN"]:
        cen[c] = cen[c] - cen.groupby("文献编号")[c].transform("mean")
    cen["Epit"] = cen["Epit"] - cen.groupby("文献编号")["Epit"].transform("mean")
    keep = cen.groupby("文献编号")["样本"].transform("size") >= 3
    cen = cen[keep]
    print(f"  篇内中心化子集：n={len(cen)}  篇数={cen['文献编号'].nunique()}")
    r = run(cen[["Epit"] + COMP + ["PREN"]].assign(文献编号=cen["文献编号"]), "篇内去均值")
    print(f"  RF  R2={r['RF']['R2']:.3f}  MAE={r['RF']['MAE']:.3f}")
    print(f"  Ridge R2={r['Ridge']['R2']:.3f}   GBR R2={r['GBR']['R2']:.3f}")
    if "PREN_only_polyfit" in r:
        print(f"  仅 PREN 篇内线性：R2={r['PREN_only_polyfit']['R2_in_sample']:.3f}"
              f"  斜率={r['PREN_only_polyfit']['slope']} V/PREN单位")

    print("\n" + "=" * 78)
    print("诊断 6  严格 PREN 子集（Cr/Mo/N 三项均报出，不拿 0 顶替）")
    print("=" * 78)
    strict = numeric.dropna(subset=["Cr", "Mo", "N"]).copy()
    strict["PREN"] = strict["Cr"] + 3.3 * strict["Mo"] + 16 * strict["N"]
    print(f"  n={len(strict)}  篇数={strict['文献编号'].nunique()}")
    if len(strict) >= 40:
        r = run(strict[["Epit"] + COMP + ["PREN"]].assign(文献编号=strict["文献编号"]),
                "严格 PREN 子集")
        print(f"  RF  R2={r['RF']['R2']:.3f}  MAE={r['RF']['MAE']:.3f}")
        print(f"  Ridge R2={r['Ridge']['R2']:.3f}   GBR R2={r['GBR']['R2']:.3f}")
        if "PREN_only_polyfit" in r:
            print(f"  仅 PREN 一元：R2={r['PREN_only_polyfit']['R2_in_sample']:.3f}"
                  f"  斜率={r['PREN_only_polyfit']['slope']}")

    print("\n" + "=" * 78)
    print("诊断 4  文本 Epit（± / 区间）取中值纳入后，行数与 R2 变化")
    print("=" * 78)
    txt = df_all[df_all["Epit"].isna() & df_all["Epit_raw"].notna()].copy()
    txt["Epit"] = txt["Epit_raw"].map(mid)
    txt = txt.dropna(subset=["Epit"]).copy()
    txt["PREN"] = txt.apply(pren, axis=1)
    print(f"  可解析出中值的文本行：{len(txt)} / {len(df_all[df_all['Epit_raw'].notna()]) - len(numeric)}")
    if len(txt):
        merged = pd.concat([numeric, txt], ignore_index=True)
        r = run(merged[["Epit"] + COMP + ["PREN"]].assign(文献编号=merged["文献编号"]),
                f"数值 {len(numeric)} + 文本中值 {len(txt)}")
        print(f"  合并后 n={r['行数']}  篇数={r['篇数']}")
        print(f"  RF  R2={r['RF']['R2']:.3f}  MAE={r['RF']['MAE']:.3f}")
        print(f"  Ridge R2={r['Ridge']['R2']:.3f}   GBR R2={r['GBR']['R2']:.3f}")
        txt.to_csv(ROOT / "model_v1_文本Epit解析_20261110.csv", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    main()
