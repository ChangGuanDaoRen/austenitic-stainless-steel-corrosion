# -*- coding: utf-8 -*-
"""只读：核对 Fe 列（Q=17）在 43/45/47/48/54 等块的现值，
并全表扫描 A–Z 是否还有其它「换算说明」（→SCE / SHE / 0.1989）。"""
import openpyxl

wb = openpyxl.load_workbook("0.xlsx", data_only=True)
ws = wb["Sheet1"]
blocks = [(r, int(ws.cell(r, 1).value)) for r in range(3, ws.max_row + 1)
          if ws.cell(r, 1).value not in (None, "")]

print("=== Fe 列（Q）在关注块的值 ===")
for r, b in blocks:
    if b in (43, 45, 47, 48, 54, 28, 29, 40, 43, 51, 52, 53, 56, 57, 58, 59, 60, 61, 62, 63):
        print("block", b, "r", r, "Fe(Q)=", repr(ws.cell(r, 17).value),
              "T=", repr(ws.cell(r, 20).value))

print("\n=== 全表 A–Z 含换算说明的单元格 ===")
for r in range(1, ws.max_row + 1):
    for c in range(1, 27):
        v = ws.cell(r, c).value
        if isinstance(v, str) and (("→" in v) or ("SHE" in v) or ("0.1989" in v)):
            print("cell", ws.cell(r, c).coordinate, repr(v))

print("\n=== Z 列合并区（备注）数量 ===")
zm = [str(m) for m in ws.merged_cells.ranges
      if m.min_col == 26 and m.max_col == 26]
print("Z merges", len(zm))
print("first5", zm[:5], "last5", zm[-5:])
print("AA merges", [str(m) for m in ws.merged_cells.ranges if m.min_col == 27])
