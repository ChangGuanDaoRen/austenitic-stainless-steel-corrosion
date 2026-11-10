# -*- coding: utf-8 -*-
"""
v2：按师兄对话口径建模对比（对话\总对话_转写.md 10-24 / 10-26 / 10-29）
清洗按师兄原话执行：
  - 10-24「元素成分不含那种元素直接写0」→ 成分缺失补 0
  - 10-24「点蚀电位统一用 SC」→ SCE 统一，Ag/AgCl 按 0.034/0.044 V 折算
  - 10-26「只考虑最后一步热处理的温度和时间」→ 子集取单步热处理且末步温度可知
  - 10-29「热轧+固溶 / 冷轧+退火两部分，不要两步热处理」→ 再按成形方式切
输出：控制台 + model_v2_20261110.json
"""
import json
import re
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
from sklearn.ensemble import GradientBoostingRegressor, RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.model_selection import GroupKFold, cross_val_predict
from sklearn.pipeline import Pipeline
from sklearn.ensemble import RandomForestRegressor as RF
from sklearn.metrics import r2_score, mean_absolute_error

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from parse_process_00_20261110 import norm, parse_form, parse_treat, num  # noqa

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "00.xlsx"
COMP = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni", "Ti", "Nb", "Co", "Sn"]

# 参比 → (统一到 SCE 的偏移, 类别)，E_SCE = E_原文 + offset
RE_MAP = {
    "SCE": (0.0, "SCE"), "饱和甘汞电极": (0.0, "SCE"), "饱和甘汞电极(SCE)": (0.0, "SCE"),
    "饱和KCl甘汞电极(SCE)": (0.0, "SCE"), "SSCE": (0.0, "SCE"),
    "SCE（0.1M KCl盐桥）": (0.0, "SCE"),
    "Ag/AgCl(3M KCl)": (-0.034, "AgAgCl"),
    "Ag/AgCl(饱和KCl)": (-0.044, "AgAgCl"),
    "Ag/AgCl": (-0.034, "AgAgCl"),
    "未提及": (np.nan, "未提及"),
}


def load():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    wb.close()
    last, recs = None, []
    for r in rows:
        if r[0] is not None:
            last = r[0]
        if r[1] is None:
            continue
        d = {"文献": last, "样本": r[1], "工艺": r[17], "参比原文": r[19],
             "温度": r[21], "Epit_raw": r[22]}
        for j, c in enumerate(COMP):
            d[c] = num(r[2 + j])
        recs.append(d)
    df = pd.DataFrame(recs)
    df["Epit"] = df["Epit_raw"].map(num)
    df = df[df["Epit"].notna()].copy()
    df["_n"] = df["工艺"].map(norm)
    df["成形"] = df["_n"].map(parse_form)
    t = df["_n"].map(parse_treat)
    df["末步"] = [x[0] for x in t]
    df["末步温度"] = [x[1] for x in t]
    df["热处理步数"] = [x[2] for x in t]
    df["两步"] = df["热处理步数"] >= 2
    return df


def unify_re(df):
    key = df["参比原文"].astype(str).str.strip()
    off = key.map(lambda k: RE_MAP.get(k, (np.nan, "其它"))[0])
    df["Epit_SCE"] = df["Epit"] + off
    df["参比类"] = key.map(lambda k: RE_MAP.get(k, (np.nan, "其它"))[1])
    return df


def evaluate(df, target, feats, tag, log):
    d = df.dropna(subset=[target]).copy()
    X = d[feats]
    y = d[target].values
    g = d["文献"].astype(str).values
    cv = GroupKFold(n_splits=min(5, len(np.unique(g))))
    pipe = Pipeline([("imp", SimpleImputer(strategy="median")),
                     ("m", RF(n_estimators=500, min_samples_leaf=2,
                              random_state=0, n_jobs=-1))])
    pred = cross_val_predict(pipe, X, y, groups=g, cv=cv)
    r = {"口径": tag, "n": int(len(d)), "篇": int(d["文献"].nunique()),
         "R2": round(float(r2_score(y, pred)), 4),
         "MAE": round(float(mean_absolute_error(y, pred)), 4)}
    log.append(r)
    return r


def main():
    df = load()
    df = unify_re(df)
    # 师兄 10-24：不含的元素写 0
    for c in COMP:
        df[c] = df[c].fillna(0.0)
    df["PREN"] = df["Cr"] + 3.3 * df["Mo"] + 16 * df["N"]
    df["PREN_strict"] = np.where(
        (df["Cr"] > 0),
        df["Cr"] + 3.3 * df["Mo"] + 16 * df["N"], np.nan)

    log = []
    print("=" * 78)
    print("v2 口径对比（RF，GroupKFold 按文献分组）")
    print("=" * 78)
    r0 = evaluate(df, "Epit", COMP + ["PREN"], "① 现状：500 条，未统一 SCE、成分补0", log)
    r1 = evaluate(df, "Epit_SCE", COMP + ["PREN"], "② 统一 SCE 后", log)

    single = df[(~df["两步"]) & df["末步温度"].notna()].copy()
    r2 = evaluate(single, "Epit_SCE", COMP + ["PREN"], "③ 单步热处理+末步温度可知（10-26 口径）", log)
    r3 = evaluate(single, "Epit_SCE", COMP + ["PREN", "末步温度"], "④ ③+末步温度作特征", log)

    s2 = single[single["成形"] != "未提及"].copy()
    r4 = evaluate(s2, "Epit_SCE", COMP + ["PREN", "末步温度"], "⑤ ④且成形方式明确", log)

    hot = single[(single["成形"] == "热轧") & (single["末步"].str.startswith("固溶"))]
    r5 = evaluate(hot, "Epit_SCE", COMP + ["PREN", "末步温度"], "⑥ 热轧+固溶（10-29）", log)
    cast = single[(single["成形"] == "铸造") & (single["末步"].str.startswith("固溶"))]
    r6 = evaluate(cast, "Epit_SCE", COMP + ["PREN", "末步温度"], "⑦ 铸造(铸态)+固溶", log)
    cold = single[(single["成形"] == "冷轧") & (single["末步"].str.startswith("退火"))]
    r7 = evaluate(cold, "Epit_SCE", COMP + ["PREN", "末步温度"], "⑧ 冷轧+退火（10-29）", log)

    for r in log:
        print(f"  {r['口径']:<48s} n={r['n']:4d} 篇={r['篇']:3d} "
              f"R2={r['R2']:+.3f} MAE={r['MAE']:.3f}")

    print("\n" + "=" * 78)
    print("参比统一情况")
    print("=" * 78)
    print(df["参比类"].value_counts().to_string())
    print(f"\n可统一到 SCE：{int(df['Epit_SCE'].notna().sum())} / {len(df)} 条")
    print(f"折算量级：Ag/AgCl(3M KCl) −0.034 V，Ag/AgCl(饱和KCl) −0.044 V")

    df.to_csv(ROOT / "model_v2_建模集_20261110.csv", index=False, encoding="utf-8-sig")
    (ROOT / "model_v2_20261110.json").write_text(
        json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\n已写出 model_v2_20261110.json / model_v2_建模集_20261110.csv")


if __name__ == "__main__":
    main()
