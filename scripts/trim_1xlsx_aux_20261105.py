import shutil, csv, os, sys
import openpyxl

SRC = '1.xlsx'
BAK = r'备份\1_删辅助表前_20261105.xlsx'
DUMP = r'temp\1xlsx_辅助表_删前_20261105'
KEEP = 'Sheet1'

if not os.path.exists(BAK):
    shutil.copy2(SRC, BAK)
    print('[备份]', BAK)
else:
    print('[备份已存在]', BAK)

os.makedirs(DUMP, exist_ok=True)

# 1) 先把要删的 5 张表导成 CSV 留痕
wb = openpyxl.load_workbook(SRC, data_only=True)
for ws in wb.worksheets:
    if ws.title == KEEP:
        continue
    p = os.path.join(DUMP, ws.title + '.csv')
    with open(p, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            w.writerow(['' if c is None else c for c in row])
    print('[导出]', p, ws.max_row, 'x', ws.max_column)

# 2) 快照 Sheet1 全量值，供删后逐格比对
keep_vals = [['' if c is None else c for c in r]
             for r in wb[KEEP].iter_rows(min_row=1, max_row=wb[KEEP].max_row,
                                          max_col=wb[KEEP].max_column, values_only=True)]
keep_ranges = sorted(str(r) for r in wb[KEEP].merged_cells.ranges)
print('[快照]', len(keep_vals), '行 x', len(keep_vals[0]), '列 | 合并区', len(keep_ranges))
wb.close()

if '--commit' not in sys.argv:
    print('dry-run：未写盘')
    sys.exit(0)

# 3) 删除其余工作表
wb = openpyxl.load_workbook(SRC)
for name in list(wb.sheetnames):
    if name != KEEP:
        del wb[name]
        print('[删除]', name)
wb.save(SRC)
wb.close()

# 4) 校验
wb2 = openpyxl.load_workbook(SRC, data_only=True)
print('[结果] sheets =', wb2.sheetnames)
ws = wb2[KEEP]
print('[结果] Sheet1 dims =', ws.dimensions, '| 合并区', len(ws.merged_cells.ranges))
now = [['' if c is None else c for c in r]
       for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column, values_only=True)]
diff = 0
for i, (a, b) in enumerate(zip(keep_vals, now), 1):
    if a != b:
        diff += 1
        for j, (x, y) in enumerate(zip(a, b), 1):
            if x != y:
                print('  差异 r%dc%d: %r -> %r' % (i, j, x, y))
print('[校验] 行数 %d vs %d | 逐格差异行 %d | 合并区 %d vs %d' %
      (len(keep_vals), len(now), diff, len(keep_ranges), len(ws.merged_cells.ranges)))
wb2.close()
