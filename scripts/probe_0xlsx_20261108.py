# -*- coding: utf-8 -*-
"""只读：导出 0.xlsx 全部 60 块 Y/Z/AA 原文，统计各列标注写法。"""
import openpyxl
import json
from collections import Counter

wb = openpyxl.load_workbook("0.xlsx", data_only=True)
ws = wb["Sheet1"]
MAXC = ws.max_column

blocks = []
for r in range(3, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if v is not None and str(v).strip() != "":
        blocks.append((r, v))

out = []
for (r, bid) in blocks:
    row = {"block": bid, "row": r,
           "Y": ws.cell(r, 25).value, "Z": ws.cell(r, 26).value,
           "AA": ws.cell(r, 27).value if MAXC >= 27 else None}
    out.append(row)

with open("temp/0xlsx_notes_dump_20261108.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

# 打印每块 Z/AA，供人工判读
for d in out:
    z = (str(d["Z"]).replace("\n", " ") if d["Z"] is not None else "(空)")
    aa = (str(d["AA"]).replace("\n", " ") if d["AA"] is not None else "(空)")
    print(f"块{d['block']} r{d['row']} | Z: {z[:400]} | AA: {aa[:80]}")

print("\n=== 各列值形态统计（C–X, 26 列） ===")
headers = [ws.cell(2, c).value or ws.cell(1, c).value for c in range(1, MAXC + 1)]
for c in range(3, 25):  # C..X
    vals = [ws.cell(r, c).value for r in range(3, ws.max_row + 1)]
    nonempty = [v for v in vals if v is not None and str(v).strip() != ""]
    types = Counter("num" if isinstance(v, (int, float)) else "str" for v in nonempty)
    strs = [str(v) for v in nonempty if isinstance(v, str)]
    # 含比较符/区间/N/A/中文说明的写法
    special = [s for s in strs if any(k in s for k in ("<", ">", "≤", "≥", "~", "-", "N/A", "未", "约", "室温", "余量", "余", "～", "—"))]
    print(f"列{c} {headers[c-1]}: 非空{len(nonempty)} 数值{types['num']} 文本{types['str']} 特殊写法{len(special)}")
    for s in special[:6]:
        print("    ", s[:80])
