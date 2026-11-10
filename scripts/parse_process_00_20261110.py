# -*- coding: utf-8 -*-
"""
按师兄口径（对话\总对话_转写.md，09-28 / 10-24 / 10-26 / 10-29）给 00.xlsx 数值行分档。
规则（照师兄原话，不自行加严）：
  1. 09-28「一定要搜铸造的材料」→ 成形方式标出，铸造/铸态不排除
  2. 10-24「最好选择相同工艺，例如都热轧+固溶，或都冷轧+退火」
  3. 10-26「只考虑最后一步热处理的温度和时间」
  4. 10-29「分热轧+固溶、冷轧+退火两部分，不要两步热处理」
温度带定性质：末步 ≥1000℃ 视为固溶（原文写"退火"也归固溶，如 1100℃退火15h）；700–900℃ 为退火。
输出：process_档位_20261110.csv / 控制台三张表
仍属机抽候选，最终逐条判定需看原文（AGENTS.md 判定红线）。
"""
import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "00.xlsx"

TREAT = ["固溶", "退火", "时效", "敏化", "稳定化", "热老化", "回火"]
FORM = [("冷轧", 4), ("热轧", 3), ("锻造", 2), ("锻", 2), ("铸态", 1), ("砂型", 1),
        ("离心", 1), ("铸造", 1), ("铸锭", 1)]


def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("°C", "℃").replace("～", "-").replace("—", "-").replace("–", "-")
    s = s.replace("~", "-").replace(" ", "").replace("\u3000", "")
    return s


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return np.nan


def parse_form(s):
    hits = []
    for kw, lvl in FORM:
        for m in re.finditer(kw, s):
            hits.append((m.start(), lvl, kw))
    if not hits:
        return "未提及"
    hits.sort()
    last = hits[-1]
    kw = last[2]
    if kw in ("锻", "锻造"):
        return "锻造"
    if kw in ("铸态", "铸造", "铸锭", "砂型", "离心"):
        return "铸造"
    return kw


def parse_treat(s):
    """返回 (末步性质, 温度, 步数, 步骤串)"""
    temps = [(m.start(), float(m.group(1))) for m in
             re.finditer(r"(?<!\d)(\d{3,4}(?:\.\d+)?)(?!\d)", s)]
    temps = [(p, t) for p, t in temps if 200 <= t <= 1400]
    steps = []
    for m in re.finditer("|".join(TREAT), s):
        kw, pos = m.group(0), m.start()
        before = [t for p, t in temps if pos - 30 <= p < pos]
        after = [t for p, t in temps if pos <= p <= pos + 25]
        T = before[-1] if before else (after[0] if after else np.nan)
        steps.append((kw, T, pos))
    if not steps:
        return ("无热处理记录", np.nan, 0, "")
    kw, T, _ = steps[-1]
    # 温度带定性质
    if not pd.isna(T) and T >= 1000:
        kind = "固溶" if kw in ("固溶", "退火") else kw
        if kw == "退火":
            kind = "固溶(原文写退火)"
    elif kw in ("固溶", "退火"):
        kind = "退火" if kw == "退火" else "固溶(低温)"
    else:
        kind = kw
    seq = "→".join(f"{k}{'' if pd.isna(t) else int(t)}" for k, t, _ in steps)
    return (kind, T, len(steps), seq)


def main():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    wb.close()
    last_lit, recs = None, []
    for r in rows:
        if r[0] is not None:
            last_lit = r[0]
        if r[1] is None:
            continue
        recs.append({"文献": last_lit, "样本": r[1], "工艺": r[17],
                     "Epit": num(r[22]), "参比": r[19]})
    df = pd.DataFrame(recs)
    df = df[df["Epit"].notna()].copy()
    df["_n"] = df["工艺"].map(norm)
    df["成形"] = df["_n"].map(parse_form)
    t = df["_n"].map(parse_treat)
    df["末步"] = [x[0] for x in t]
    df["末步温度"] = [x[1] for x in t]
    df["热处理步数"] = [x[2] for x in t]
    df["步骤串"] = [x[3] for x in t]
    df["两步"] = df["热处理步数"] >= 2

    print("=" * 78)
    print(f"00.xlsx 有数值 Epit 的行：{len(df)} 条 / {df['文献'].nunique()} 篇")
    print("=" * 78)

    print("\n【表1】成形方式 × 末步热处理 交叉表（条数 / 篇数）")
    piv = df.pivot_table(index="成形", columns="末步", values="样本",
                         aggfunc="size", fill_value=0)
    print(piv.to_string())

    print("\n【表2】师兄 10-29 的两条主体路线（单步热处理，无时效/敏化/稳定化）")
    single = df[~df["两步"]]
    for form, kind, label in [("热轧", "固溶", "热轧+固溶"),
                              ("冷轧", "退火", "冷轧+退火"),
                              ("冷轧", "固溶", "冷轧+固溶"),
                              ("铸造", "固溶", "铸造(铸态)+固溶"),
                              ("锻造", "固溶", "锻造+固溶")]:
        m = single[(single["成形"] == form) & (single["末步"].str.startswith(kind))]
        print(f"  {label:16s} {len(m):4d} 条 / {m['文献'].nunique():3d} 篇")
    m = single[(single["成形"] == "热轧") & (single["末步"].str.startswith("固溶"))]
    if len(m):
        tt = m["末步温度"]
        print(f"\n  其中「热轧+固溶」温度：中位 {tt.median():.0f}℃，"
              f"1050–1150℃ 区间 {int(((tt>=1050)&(tt<=1150)).sum())} 条")

    print("\n【表3】按温度带的末步分布（单步行）")
    band = pd.cut(single["末步温度"], bins=[0, 700, 900, 1000, 1050, 1150, 1400],
                  labels=["<700", "700-900", "900-1000", "1000-1050",
                          "1050-1150", ">1150"])
    print(pd.crosstab(band, single["末步"]).to_string())

    print("\n【表4】被排除/待定的部分")
    print(f"  两步热处理（固溶后再时效/敏化/稳定化）: {int(df['两步'].sum())} 条 / "
          f"{df[df['两步']]['文献'].nunique()} 篇 —— 师兄 10-29 明确不要")
    noh = df[df["末步"] == "无热处理记录"]
    print(f"  无热处理记录（冷轧态未写退火、铸态直接用等）: {len(noh)} 条 / "
          f"{noh['文献'].nunique()} 篇")
    nb = df[df["成形"] == "未提及"]
    print(f"  成形方式未提及: {len(nb)} 条 / {nb['文献'].nunique()} 篇")

    print("\n【表5】按师兄 10-26「只考虑最后一步热处理温度和时间」后的可用量")
    ok = single[(single["末步温度"].notna())]
    print(f"  末步温度可知且单步: {len(ok)} 条 / {ok['文献'].nunique()} 篇")
    print(f"  其中成形方式也明确的: "
          f"{len(ok[ok['成形']!='未提及'])} 条 / {ok[ok['成形']!='未提及']['文献'].nunique()} 篇")

    df.drop(columns=["_n"]).to_csv(ROOT / "process_档位_20261110.csv",
                                   index=False, encoding="utf-8-sig")
    print("\n已写出 process_档位_20261110.csv")


if __name__ == "__main__":
    main()
