# -*- coding: utf-8 -*-
"""0.xlsx 数值修订（2026-11-05）：改 3 格 Epit + 3 段块备注追加。

对应报告：0xlsx_数值问题说明_20261105.md 第二节「明确不符」3 条。
    row 162  文献28 Mo1N2   0.94   -> 0.87
    row  81  文献13 G 材    -0.242 -> -0.02
    row  72  文献12 SCS14-B  0.582 -> 0.96
备注列 AA 按块合并（AA160:AA162 / AA71:AA74 / AA75:AA82），只写合并区首行。

用法：
    python scripts/fix_num_20261105.py            # 干跑，写 temp/audit0/0_num预览.xlsx
    python scripts/fix_num_20261105.py --commit   # 备份后写回 0.xlsx
"""
import os
import shutil
import sys

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, '0.xlsx')
BACKUP = os.path.join(BASE, '备份', '0_基准_数值修订前_20261105.xlsx')
PREVIEW = os.path.join(BASE, 'temp', 'audit0', '0_num预览.xlsx')
SHEET = 'Sheet1'
COL_EPIT = 24
COL_NOTE = 27

# (行, 旧值, 新值)
CELLS = [
    (162, 0.94, 0.87),
    (81, -0.242, -0.02),
    (72, 0.582, 0.96),
]

# (合并区首行, 追加文字)  追加在备注末尾
NOTES = [
    (160, '\n【2026-11-05 复核修订】Mo1N2 原记 0.940 V 与 Fig.3a 不符：三条曲线在 0.86–0.96 V 依次右折，'
          '黑 Mo2N1 0.94–0.96、红 Mo1N1 0.90、蓝 Mo1N2 0.86–0.88（最低）；0.940 已顶到 Mo2N1 上包络，'
          '排序与本图相反。改记 0.87 V（图读 ±0.02）。'),
    (75, '\n【2026-11-05 复核修订】G 原记 −0.242 V 与 Fig.4 不符：图例 G 为实心菱形（单点，δ≈1.3，'
         '纵轴约 −20 mV），H 为实心方块（约 −290 mV）；−0.242 与 H 几乎同高，属 G/H 混淆。'
         '改记 −0.02 V（图读 ±0.03）。'),
    (71, '\n【2026-11-05 复核修订】SCS14-B 原记 0.582 V 与 Fig.4 不符：B（空心圆）三点 '
         '0.69 / 0.98 / 1.22，按本块「三点取平均」口径应为 0.96；0.582 不落在任何一点上。改记 0.96 V。'),
]


def snapshot(path, sheet=SHEET):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    grid = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            grid[(r, c)] = ws.cell(r, c).value
    meta = {
        'rows': ws.max_row,
        'cols': ws.max_column,
        'merged': sorted(str(m) for m in ws.merged_cells.ranges),
        'sheets': wb.sheetnames,
        'widths': {k: v.width for k, v in ws.column_dimensions.items()},
    }
    wb.close()
    return grid, meta


def apply(path):
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    changed = []
    for row, old, new in CELLS:
        cur = ws.cell(row, COL_EPIT).value
        if cur != old:
            raise SystemExit('行 %d 现值 %r 与预期旧值 %r 不符，终止' % (row, cur, old))
        ws.cell(row, COL_EPIT).value = new
        changed.append((row, COL_EPIT, old, new))
    for row, text in NOTES:
        cur = ws.cell(row, COL_NOTE).value or ''
        if '2026-11-05 复核修订' in cur:
            print('  备注行 %d 已有 11-05 修订段，跳过' % row)
            continue
        ws.cell(row, COL_NOTE).value = cur + text
        changed.append((row, COL_NOTE, '（追加）', text[:24] + '…'))
    wb.save(path)
    wb.close()
    return changed


def main():
    commit = '--commit' in sys.argv
    sheet_name = SHEET  # noqa: F841  (apply 用全局)
    target = SRC if commit else PREVIEW
    ref = BACKUP if commit else SRC

    if commit:
        if not os.path.exists(BACKUP):
            os.makedirs(os.path.dirname(BACKUP), exist_ok=True)
            shutil.copy2(SRC, BACKUP)
            print('已备份 ->', os.path.relpath(BACKUP, BASE))
        else:
            print('备份已存在，沿用 ->', os.path.relpath(BACKUP, BASE))
    else:
        os.makedirs(os.path.dirname(PREVIEW), exist_ok=True)
        shutil.copy2(SRC, PREVIEW)

    g0, m0 = snapshot(ref)
    changed = apply(target)
    g1, m1 = snapshot(target)

    print('\n计划改格 %d 处：' % len(changed))
    for c in changed:
        print('   r%-4d c%-2d  %r -> %r' % c)

    diff = [(k, g0.get(k), g1.get(k)) for k in set(g0) | set(g1) if g0.get(k) != g1.get(k)]
    print('\n逐格 diff：%d 格变化' % len(diff))
    for k, a, b in sorted(diff):
        print('   r%d c%d  %r -> %r' % (k[0], k[1], a, b))

    # 结构校验
    ok = True
    for key in ('rows', 'cols', 'sheets'):
        if m0[key] != m1[key]:
            ok = False
            print('!! 结构不一致 %s: %r -> %r' % (key, m0[key], m1[key]))
    if len(m0['merged']) != len(m1['merged']):
        ok = False
        print('!! 合并区数量变化 %d -> %d' % (len(m0['merged']), len(m1['merged'])))
    lost = set(m0['merged']) - set(m1['merged'])
    if lost:
        ok = False
        print('!! 丢失合并区', lost)
    if m0['widths'] != m1['widths']:
        print('   注意：列宽映射有差异（openpyxl 重写常见，需人工复核）')

    expect = len(CELLS) + len([1 for r, _ in NOTES])
    if len(diff) != expect:
        ok = False
        print('!! 预期 %d 格变化，实际 %d 格' % (expect, len(diff)))

    print('\n校验：%s' % ('通过' if ok else '未通过'))
    print('目标文件：', os.path.relpath(target, BASE))
    if not commit:
        print('（干跑，0.xlsx 未动；加 --commit 写回）')


if __name__ == '__main__':
    main()
