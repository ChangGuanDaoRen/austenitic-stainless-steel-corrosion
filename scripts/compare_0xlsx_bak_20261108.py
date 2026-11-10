# -*- coding: utf-8 -*-
"""只读对比：加列前备份 vs 当前 0.xlsx（A–Z 值与样式），并统计 Z 备注长度。"""
import openpyxl
import statistics

cur = openpyxl.load_workbook("0.xlsx", data_only=True)["Sheet1"]
bak = openpyxl.load_workbook(r"备份\0_基准_加口径标注列前_20261108.xlsx", data_only=True)["Sheet1"]
print("cur dims", cur.max_row, cur.max_column, "merges", len(cur.merged_cells.ranges))
print("bak dims", bak.max_row, bak.max_column, "merges", len(bak.merged_cells.ranges))

diff = 0
style_diff = 0
for r in range(1, max(cur.max_row, bak.max_row) + 1):
    for c in range(1, 27):  # A..Z
        a = cur.cell(r, c); b = bak.cell(r, c)
        if a.value != b.value:
            diff += 1
        # 仅比较字体/边框/对齐/数字格式这些核心样式
        fa, fb = a.font, b.font
        if (fa.name, fa.sz, fa.b, fa.i) != (fb.name, fb.sz, fb.b, fb.i):
            style_diff += 1
print("A-Z value diff:", diff, "core font diff:", style_diff)

blocks = [(r, cur.cell(r, 1).value) for r in range(3, cur.max_row + 1)
          if cur.cell(r, 1).value not in (None, "")]
lens = [(len(str(cur.cell(r, 26).value)) if cur.cell(r, 26).value else 0, bid) for r, bid in blocks]
print("Z len: min", min(l for l, _ in lens), "median", statistics.median(l for l, _ in lens),
      "max", max(l for l, _ in lens), "count>200:", sum(1 for l, _ in lens if l > 200))
print("long(>200):", [str(b) + ":" + str(l) for l, b in lens if l > 200])

# 原始示例表头（若存在）
import os
src = r"备份\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx"
if os.path.exists(src):
    o = openpyxl.load_workbook(src, data_only=True)
    s = o[o.sheetnames[0]]
    print("ORIG dims", s.max_row, s.max_column, "sheets", o.sheetnames)
    print("ORIG head1", [s.cell(1, c).value for c in range(1, s.max_column + 1)])
    print("ORIG head2", [s.cell(2, c).value for c in range(1, s.max_column + 1)])
    for r in range(3, min(8, s.max_row + 1)):
        print("ORIG r", r, "last2", [s.cell(r, c).value for c in range(s.max_column - 1, s.max_column + 1)])
