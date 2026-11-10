# -*- coding: utf-8 -*-
"""校验清洗版：原件未变、总表条数正确、参比全 SCE、换算数值正确、各工艺 sheet 条数正确。"""
import sys
from collections import Counter

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
SRC = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx"
DST = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版.xlsx"

a = openpyxl.load_workbook(SRC, data_only=True)["Sheet1"]
b = openpyxl.load_workbook(DST, data_only=True)
o = b["00_原始数据"]

diff = 0
for r in range(1, a.max_row + 1):
    for c in range(1, a.max_column + 1):
        if a.cell(row=r, column=c).value != o.cell(row=r, column=c).value:
            diff += 1
print("00_原始数据 与原件的差异单元格数:", diff, "（应为 0）")
print("00_原始数据 合并单元格数:", len(o.merged_cells.ranges), "（原件 68）")

t = b["01_清洗总表"]
data = [r for r in range(3, t.max_row + 1) if t.cell(row=r, column=2).value not in (None, "")]
print("\n01_清洗总表 数据行:", len(data))

print("参比电极取值:", dict(Counter(str(t.cell(row=r, column=20).value) for r in data)))

def isnum(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False

epits = [t.cell(row=r, column=24).value for r in data]
print("Epit 数值条数:", sum(1 for v in epits if isnum(v)), " N/A:", sum(1 for v in epits if not isnum(v)))
print("Epit 范围:", min(float(v) for v in epits if isnum(v)), "~", max(float(v) for v in epits if isnum(v)))

print("\n元素列 0 值计数:")
for c, el in [(3, "C"), (4, "Si"), (5, "N"), (6, "P"), (7, "S"), (8, "Cu"), (9, "Mn"), (10, "Cr"),
              (11, "Mo"), (12, "Ni"), (13, "Ti"), (14, "Nb"), (15, "Co"), (16, "Sn"), (17, "Fe")]:
    cnt = sum(1 for r in data if t.cell(row=r, column=c).value == 0)
    txt = Counter(str(t.cell(row=r, column=c).value) for r in data
                  if not isnum(t.cell(row=r, column=c).value))
    print(f"   {el:<3} 0值 {cnt:>3} 条" + (f" | 非数值: {dict(txt)}" if txt else ""))

# 换算抽查：对比原件与清洗版的 Epit
orig_rows = {}
paper = None
for r in range(3, a.max_row + 1):
    p = a.cell(row=r, column=1).value
    if p not in (None, ""):
        paper = p
    name = a.cell(row=r, column=2).value
    if name in (None, ""):
        continue
    orig_rows[(str(paper), str(name))] = (a.cell(row=r, column=24).value, a.cell(row=r, column=20).value)

SHIFT = {"Ag/AgCl(3M KCl)": -0.031, "Ag/AgCl(sat.KCl)": -0.044, "Ag/AgCl": -0.031, "SCE": 0.0}
bad = []
paper = None
for r in range(3, t.max_row + 1):
    p = t.cell(row=r, column=1).value
    if p not in (None, ""):
        paper = p
    name = t.cell(row=r, column=2).value
    if name in (None, ""):
        continue
    key = (str(paper), str(name))
    o_epit, o_re = orig_rows[key]
    n_epit = t.cell(row=r, column=24).value
    if not isnum(o_epit):
        continue
    exp = round(float(o_epit) + SHIFT[str(o_re).strip()], 3)
    if abs(float(n_epit) - exp) > 1e-9:
        bad.append((key, o_epit, o_re, n_epit, exp))
print("\n换算校验不通过的条数:", len(bad))
for x in bad[:10]:
    print("   ", x)

print("\n各工艺 sheet 条数:")
for name in b.sheetnames:
    if name in ("00_原始数据", "01_清洗总表", "说明"):
        continue
    sh = b[name]
    n = [r for r in range(3, sh.max_row + 1) if sh.cell(row=r, column=2).value not in (None, "")]
    print(f"   {name}: {len(n)}")
    # 确认该 sheet 内的工艺关键词一致
