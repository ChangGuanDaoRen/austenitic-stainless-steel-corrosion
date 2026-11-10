# -*- coding: utf-8 -*-
"""只读：导出 0.xlsx 的 T 列（参比电极）全部取值与所在行，
以及需要按特殊情况写备注的块的 W 列（Epit）实值，供改备注/清理参比列使用。"""
import openpyxl
from collections import defaultdict

wb = openpyxl.load_workbook("0.xlsx", data_only=True)
ws = wb["Sheet1"]
print("dims", ws.max_row, ws.max_column, "merges", len(ws.merged_cells.ranges))

# 块首
blocks = [(r, ws.cell(r, 1).value) for r in range(3, ws.max_row + 1)
          if ws.cell(r, 1).value not in (None, "")]
b2row = {str(b): r for r, b in blocks}
print("blocks", len(blocks))

# T 列（第 20 列）全部非空取值及所在行（物理单元格）
tvals = defaultdict(list)
for r in range(3, ws.max_row + 1):
    v = ws.cell(r, 20).value
    if v not in (None, ""):
        tvals[str(v)].append(r)
print("\n=== T 列（参比电极）取值分布 ===")
for v, rs in tvals.items():
    print(repr(v), "rows", rs)

# 含换算说明的单元格
print("\n=== T 列含 '→' / 'SHE' / '0.1989' 的单元格 ===")
for v, rs in tvals.items():
    if ("→" in v) or ("SHE" in v) or ("0.1989" in v):
        print(repr(v), "rows", rs)

# 关心块的 Epit（W 列 23）、样本名（B 列 2）、以及 Z 备注长度
focus = [28, 30, 34, 36, 38, 40, 41, 43, 49, 54, 55, 60, 61, 62, 63]
print("\n=== 关注块 B/W/样本行数 ===")
for b in focus:
    r = b2row.get(str(b))
    if not r:
        print("block", b, "MISSING"); continue
    # 找到下一块首行
    nxt = ws.max_row + 1
    for r2, _ in blocks:
        if r2 > r:
            nxt = r2; break
    samples = [(rr, ws.cell(rr, 2).value, ws.cell(rr, 23).value) for rr in range(r, nxt)]
    print("block", b, "rows", r, "-", nxt - 1, "count", len(samples))
    for rr, name, ep in samples:
        print("   r%d %s | Epit=%r" % (rr, str(name)[:36], ep))
