# -*- coding: utf-8 -*-
"""核对第9批入库前备份与当前 0.xlsx 是否字节一致、结构一致。只读。"""
import hashlib
import os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
a = os.path.join(ROOT, "0.xlsx")
b = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
c = os.path.join(ROOT, "备份", "0_基准_加口径标注列前_20261108.xlsx")

def md5(p):
    with open(p, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

for p in (a, b, c):
    print(os.path.relpath(p, ROOT), os.path.getsize(p), md5(p), os.path.getmtime(p))
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb["Sheet1"]
    print("  dims", ws.max_row, ws.max_column, "AA1=", ws["AA1"].value if ws.max_column >= 27 else None,
          "merges", len(ws.merged_cells.ranges))
