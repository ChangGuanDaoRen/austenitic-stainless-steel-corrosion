# -*- coding: utf-8 -*-
"""
按 2026-10-28 核查报告 + 用户裁决修改 0.xlsx（基准数据表）

用户裁决：
  1) 文献22 的 P、S（原文 P、S 合并列）—— 不管，保持原样
  2) 删除 Lewmet55 样本（文献22）
  3) 删除 Sample A、Sample B（文献24，XRD 证实为奥氏体+δ双相）
  4) 删除文献27 整篇（C002，溶液为 3.56% NaCl，非 3.5%）
同时修正核查发现的笔误：
  文献26 Si 0.346 -> 0.34（24 行）
  文献26 Epit：550×360min 0.145->0.045；650×180min 0.043->0.004；750×60min -0.111->-0.011
  文献23 N 未提及 -> 0.033（5 行）

用法：python scripts/fix_0xlsx.py [dry|write]
"""
import sys
import openpyxl

MODE = sys.argv[1] if len(sys.argv) > 1 else "dry"
SRC = "0.xlsx"

COL_SAMPLE = 2
COL_SI = 4
COL_N = 5
COL_EPIT = 24

DELETE_EXACT = {"Lewmet55"}
DELETE_PREFIX = ("Sample A (", "Sample B (")
DELETE_PAPER27 = {"C1", "C2", "C3", "C4", "C5", "C6", "C7"}

wb = openpyxl.load_workbook(SRC)
ws = wb["Sheet1"]

max_row = ws.max_row
# 1) 收集待删行
to_delete = []
for r in range(3, max_row + 1):
    name = ws.cell(row=r, column=COL_SAMPLE).value
    if not name:
        continue
    name = str(name)
    if name in DELETE_EXACT or name.startswith(DELETE_PREFIX) or name in DELETE_PAPER27:
        to_delete.append((r, name))

# 2) 修正数值
edits = []
for r in range(3, max_row + 1):
    name = ws.cell(row=r, column=COL_SAMPLE).value
    if not name:
        continue
    name = str(name)
    if name.startswith("HNASS"):
        if ws.cell(row=r, column=COL_SI).value == 0.346:
            edits.append((r, "Si", ws.cell(row=r, column=COL_SI).value, 0.34))
        if name == "HNASS-550℃-360min":
            edits.append((r, "Epit", ws.cell(row=r, column=COL_EPIT).value, 0.045))
        elif name == "HNASS-650℃-180min":
            edits.append((r, "Epit", ws.cell(row=r, column=COL_EPIT).value, 0.004))
        elif name == "HNASS-750℃-60min":
            edits.append((r, "Epit", ws.cell(row=r, column=COL_EPIT).value, -0.011))
    if name.startswith("Z3CN20.09M (") and ws.cell(row=r, column=COL_N).value == "未提及":
        edits.append((r, "N", ws.cell(row=r, column=COL_N).value, 0.033))

print("待删除行：")
for r, n in to_delete:
    print("  行", r, n)
print("待修正格：")
for e in edits:
    print("  行", e[0], e[1], e[2], "->", e[3])

if MODE != "write":
    print("\n[dry] 未写盘")
    sys.exit(0)

# 3) 先记录合并区，再取消合并（避免 delete_rows 破坏合并区）
merges = [str(m) for m in ws.merged_cells.ranges]
for m in merges:
    ws.unmerge_cells(m)

# 4) 自下而上删行
for r, _ in sorted(to_delete, key=lambda x: -x[0]):
    ws.delete_rows(r)

# 5) 重建合并区：按删除行数计算新行号
del_before = lambda old: sum(1 for r, _ in to_delete if r < old)
rebuilt = {}
for m in merges:
    from openpyxl.worksheet.cell_range import CellRange
    cr = CellRange(m)
    new_min = cr.min_row - del_before(cr.min_row)
    new_max = cr.max_row - sum(1 for r, _ in to_delete if r <= cr.max_row)
    if new_max < new_min or new_min < 1:
        continue
    coord = f"{cr.min_col and openpyxl.utils.get_column_letter(cr.min_col)}{new_min}:{openpyxl.utils.get_column_letter(cr.max_col)}{new_max}"
    rebuilt[coord] = (new_min, new_max)
for coord in rebuilt:
    ws.merge_cells(coord)

# 6) 应用数值修正（按样本名重新定位，避免行号漂移）
for r in range(3, ws.max_row + 1):
    name = ws.cell(row=r, column=COL_SAMPLE).value
    if not name:
        continue
    name = str(name)
    if name.startswith("HNASS"):
        ws.cell(row=r, column=COL_SI).value = 0.34
        if name == "HNASS-550℃-360min":
            ws.cell(row=r, column=COL_EPIT).value = 0.045
        elif name == "HNASS-650℃-180min":
            ws.cell(row=r, column=COL_EPIT).value = 0.004
        elif name == "HNASS-750℃-60min":
            ws.cell(row=r, column=COL_EPIT).value = -0.011
    if name.startswith("Z3CN20.09M ("):
        ws.cell(row=r, column=COL_N).value = 0.033

wb.save(SRC)
print("\n[write] 已写盘，剩余行数（含表头2行）：", ws.max_row)
