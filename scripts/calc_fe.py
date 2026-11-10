# -*- coding: utf-8 -*-
"""把 Fe 由“余量”改为实算值：Fe = 100 - Σ(其余 14 个元素)。
带比较符的取上限、区间取中值、微量/加N 按文献20 正文给的 0.05/0.16 处理；
Cr/Ni 写“调整配比”的行没有数值，Fe 保持“余量”并在备注注明。"""
import sys
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
SRC = Path(r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版.xlsx")
DST = Path(sys.argv[1]) if len(sys.argv) > 1 else SRC

ELEM = [(3, "C"), (4, "Si"), (5, "N"), (6, "P"), (7, "S"), (8, "Cu"), (9, "Mn"), (10, "Cr"),
        (11, "Mo"), (12, "Ni"), (13, "Ti"), (14, "Nb"), (15, "Co"), (16, "Sn")]
COL_FE, COL_REMARK = 17, 27

# 文本 -> 数值 的替换口径
RULE = {
    "<0.01": 0.01, "≤0.03": 0.03, "≤0.04": 0.04, "≤0.03%": 0.03,
    "17.0-19.0": 18.0, "9.0-12.0": 10.5,
    "微量": 0.05,      # 文献20 第4章：基础新成分 N 仅为 0.05%
    "加N": 0.16,       # 文献20 摘要：在基础成分上多加入 0.11% N -> 0.05+0.11
    "调整配比": None,  # 无数值
}


def to_num(v):
    """返回 (数值, 是否已知)。"""
    if isinstance(v, (int, float)):
        return float(v), True
    s = str(v).strip() if v is not None else ""
    if s in RULE:
        return (RULE[s], RULE[s] is not None)
    try:
        return float(s), True
    except ValueError:
        return None, False


def process(ws, tag):
    header = ws.cell(row=1, column=1).value
    if header is None:
        return
    changed, skipped, notes = 0, [], []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=2).value
        if name in (None, ""):
            continue
        total, unknown = 0.0, []
        for c, el in ELEM:
            raw = ws.cell(row=r, column=c).value
            val, known = to_num(raw)
            if not known:
                unknown.append(f"{el}={raw}")
                continue
            if isinstance(raw, str) and raw.strip() in RULE and RULE[raw.strip()] is not None:
                ws.cell(row=r, column=c).value = RULE[raw.strip()]  # 文本换成数值
            total += val
        fe = round(100.0 - total, 2)
        if unknown:
            skipped.append((r, str(name), unknown, round(total, 2)))
            continue
        ws.cell(row=r, column=COL_FE).value = fe
        changed += 1
        if fe < 0:
            notes.append((r, name, fe))
    print(f"[{tag}] 算出 Fe: {changed} 条；跳过: {len(skipped)} 条")
    for s in skipped:
        print(f"    跳过 行{s[0]} {s[1]}: 未知 {s[2]}（已知元素合计 {s[3]}）")
    if notes:
        print(f"    !! Fe 为负的行: {notes}")
    return changed, skipped


wb = openpyxl.load_workbook(SRC)
total_changed = 0
for name in wb.sheetnames:
    if name in ("00_原始数据", "说明"):
        continue
    res = process(wb[name], name)
    if res:
        total_changed += res[0]

# 汇总 Fe 分布
ws = wb["01_清洗总表"]
fes = [ws.cell(row=r, column=COL_FE).value for r in range(3, ws.max_row + 1)
       if isinstance(ws.cell(row=r, column=COL_FE).value, (int, float))]
print(f"\n01_清洗总表 Fe 数值条数: {len(fes)}  范围: {min(fes)} ~ {max(fes)}")
print("Fe<50 的条数:", sum(1 for f in fes if f < 50))
print("Fe>90 的条数:", sum(1 for f in fes if f > 90))

# 更新说明 sheet
doc = wb["说明"]
add = [
    "",
    "五、Fe 实算（2026-10-25 追加）",
    "Fe 由“余量”改为 Fe = 100 − Σ(其余 14 个元素)，保留 2 位小数。为使求和成立，以下文本按口径折成数值：",
    "   <0.01 → 0.01；≤0.03 → 0.03；≤0.04 → 0.04（取上限）",
    "   17.0-19.0 → 18.0；9.0-12.0 → 10.5（取区间中值）",
    "   微量 → 0.05；加N → 0.16（文献20 正文：基础新成分 N 为 0.05%，含N 新成分在此基础上多加入 0.11%）",
    "文献20 的 4 条新成分试样，原文只写“调整 Cr、Ni 配比”而没有给数值，Fe 无法计算，保持“余量”待补。",
]
start = doc.max_row + 1
for i, t in enumerate(add):
    doc.cell(row=start + i, column=1).value = t

wb.save(DST)
print("\n已保存:", DST)

