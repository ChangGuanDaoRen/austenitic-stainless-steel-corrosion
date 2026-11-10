# -*- coding: utf-8 -*-
"""核对当前 0.xlsx 的 Z 备注是否仍为旧长文，以及 1-26/28-63 两代风格。只读。"""
import json
import os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
wb = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"), data_only=True)
ws = wb["Sheet1"]
blocks = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "")]
print("blocks", len(blocks), "maxcol", ws.max_column, "merges", len(ws.merged_cells.ranges))
lens = []
for r in blocks:
    z = ws.cell(r, 26).value or ""
    lens.append((int(ws.cell(r, 1).value), r, len(z), z[:60].replace("\n", " ")))
print("len>200:", sum(1 for x in lens if x[2] > 200), "max", max(x[2] for x in lens))
for x in lens:
    if x[2] > 200:
        print("LONG", x[0], x[1], x[2], x[3])
print("--- first 5 ---")
for x in lens[:5]:
    print(x)
print("--- last 5 ---")
for x in lens[-5:]:
    print(x)
