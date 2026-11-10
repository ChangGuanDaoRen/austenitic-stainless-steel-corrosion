# -*- coding: utf-8 -*-
"""只读抽查 0.xlsx 若干文献块的工艺原文与介质备注，用于口径文件举证。"""
import os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ws = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"), data_only=True)["Sheet1"]

want = {9, 10, 12, 13, 26, 29, 30, 34, 36, 38, 40, 41, 53, 56, 61}
cur = None
seen = set()
for r in range(3, ws.max_row + 1):
    a = ws.cell(r, 1).value
    if a not in (None, ""):
        try:
            cur = int(a)
        except (TypeError, ValueError):
            continue
    if cur not in want or cur in seen:
        continue
    seen.add(cur)
    proc = str(ws.cell(r, 18).value or "")
    info = str(ws.cell(r, 25).value or "")
    note = str(ws.cell(r, 26).value or "")
    print("=== 块 %d (行 %d) 行数待补 ===" % (cur, r))
    print("工艺: %s" % proc[:300])
    print("文献信息: %s" % info[:260])
    print("备注: %s" % note[:420])
    print()
