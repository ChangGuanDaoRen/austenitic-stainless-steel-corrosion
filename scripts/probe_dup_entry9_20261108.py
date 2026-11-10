# -*- coding: utf-8 -*-
r"""核实第 9 批已回收卡片与 0.xlsx 现有块是否同篇（只读）。"""
import os
import sys
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
wb = load_workbook(os.path.join(ROOT, "0.xlsx"), data_only=True, read_only=True)
ws = wb["Sheet1"]

blocks = []
for r in range(3, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if v not in (None, ""):
        blocks.append((r, str(v)))

print("0.xlsx 现有 %d 块，逐块 A/Y/R/W 首行：" % len(blocks))
for i, (r, bid) in enumerate(blocks):
    end = blocks[i + 1][0] - 1 if i + 1 < len(blocks) else ws.max_row
    y = str(ws.cell(r, 25).value or "").replace("\n", " ")
    proc = str(ws.cell(r, 18).value or "").replace("\n", " ")[:70]
    print("块%s r%d-%d | Y: %s | R: %s" % (bid, r, end, y[:110], proc))
