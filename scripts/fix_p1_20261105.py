# -*- coding: utf-8 -*-
"""按 0xlsx_文献核对_20261104.md 的 P1 清单（24 项 / 53 行）修订 0.xlsx。

用法：
    python scripts/fix_p1_20261105.py            # 干跑：写 temp/audit0/0_P1预览.xlsx 并逐格比对
    python scripts/fix_p1_20261105.py --commit   # 备份后写回 0.xlsx

原则：
  1. 每个待改格都带「旧值」，不一致即中止，不盲写。
  2. 只改清单内的格；改完全表 266×27 逐格比对，差异必须与清单完全相等。
  3. 合并区、列宽、行数、列数、sheet 名必须与原件一致。
"""
import os
import sys
import shutil
import hashlib
import openpyxl
from openpyxl.utils import column_index_from_string

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, '0.xlsx')
PREVIEW = os.path.join(BASE, 'temp/audit0/0_P1预览.xlsx')
BACKUP = os.path.join(BASE, '备份/0_基准_P1修订前_20261105.xlsx')
SHEET = 'Sheet1'

E = '未提及'

# (行, 列字母, 旧值, 新值, 说明)
EDITS = []


def add(rows, col, old, new, why):
    for r in rows:
        EDITS.append((r, col, old, new, why))


# ---- 文献14 C15 行成分右移一格 ----
add([84], 'J', 12.38, 20.23, '14 Cr 误取 Ni 列')
add([84], 'L', 10.24, 12.38, '14 Ni 误取 Mo 列')
add([84], 'K', 3.16, 10.24, '14 Mo 误取 Cu 列')
add([84], 'H', E, 3.16, '14 Cu 漏读')

# ---- 文献46 Epit 误取 Eprot / NaBr ----
add([224], 'X', 0.8808, 1.409, '46 T1 NaCl Epit')
add([225], 'X', 1.171, 0.8302, '46 T2 NaCl Epit')

# ---- 文献23 S ----
add(range(126, 131), 'G', 0.003, 0.0039, '23 S')

# ---- 文献48 第二行成分与电极 ----
add(range(231, 235), 'K', E, 0.0254, '48 Mo')
add(range(231, 235), 'F', E, 0.034, '48 P')
add(range(231, 235), 'G', E, 0.0076, '48 S')
add(range(231, 235), 'U', E, '高纯度石墨片', '48 CE')
add(range(231, 235), 'Y', E, 'OCP 30 min', '48 稳定时间')

# ---- 文献45 正文成分与电极 ----
add(range(218, 223), 'C', E, 0.07, '45 C')
add(range(218, 223), 'D', E, 0.33, '45 Si')
add(range(218, 223), 'I', E, 1.13, '45 Mn')
add(range(218, 223), 'J', E, 18.09, '45 Cr')
add(range(218, 223), 'L', E, 8.06, '45 Ni')
add(range(218, 223), 'G', E, 0.022, '45 S')
add(range(218, 223), 'F', E, 0.039, '45 P')
add(range(218, 223), 'S', E, '10×10×2 mm 环氧封装', '45 WE')
add(range(218, 223), 'U', E, '铂片', '45 CE')
add(range(218, 223), 'Y', E, '预浸10 min + OCP 3600 s', '45 稳定时间')

# ---- 文献54 pH（四行同为 pH 6 一组）----
add(range(261, 265), 'W', '7（中性3.5% NaCl）', '6', '54 pH')

# ---- Fe 实测值 ----
add([171], 'Q', '余量', 71.8, '32 Fe')
add(range(173, 180), 'Q', '余量', 32.45, '34 Fe')
add(range(201, 204), 'Q', '余量', 68, '41 Fe 321')
add(range(204, 207), 'Q', '余量', 54, '41 Fe 310s')

# ---- 工艺与电极 ----
add([122], 'R', '真空精密铸造，双向搅拌150s(正向4s+反向4s，循环交替)，未固溶处理',
    '真空精密铸造，双向搅拌150s(正向3s+反向3s，循环交替)，未固溶处理', '21 6号单向6s按半时记')
add([123], 'R', '真空精密铸造，双向搅拌150s(正向8s+反向8s，循环交替)，未固溶处理',
    '真空精密铸造，双向搅拌150s(正向4s+反向4s，循环交替)，未固溶处理', '21 7号单向8s按半时记')
add([258], 'R', '304-SA + 应变退火(SEN)', '304-SA + 敏化 945 K×1 h(SEN)', '53 SEN=敏化')
add(range(237, 239), 'R', '真空感应熔炼铸坯（铸态）',
    '真空感应熔炼铸坯→1200℃×2 h 随炉冷却', '50 补均匀化')
add(range(185, 191), 'U', '铂片', '铂丝', '38 CE=Pt wire')
add(range(107, 113), 'J', 18.5, 18.51, '19 Cr 精度')
add([131], 'R', '热轧80%，1180℃均匀化1h，空冷',
    '1180℃均匀化1 h→热轧80%→空冷', '24 工艺顺序')

# ---- 备注改写（整段替换，旧串必须命中）----
NOTE_EDITS = [
    (218,
     '【数据不全】全文无化学成分表，也未给逐元素成分 → 15 个元素列全部记「未提及」，绝不用 304 标准 18Cr-8Ni 回填。Epit 取自 p4 Table 1（供货态与 R10/R20/R30/R50-T200 五条）。',
     '【成分】p2 §2 实验段正文括号内已给 C 0.07、Si 0.33、Mn 1.13、Cr 18.09、Ni 8.06、S 0.022、P 0.039 wt.%；N/Cu/Mo/Ti/Nb/Co/Sn 原文未给，记「未提及」，未用 304 标准值回填。Epit 取自 p4 Table 1（供货态与 R10/R20/R30/R50-T200 五条）。'),
    (218,
     '【用途限制】本篇通篇只给牌号 type 304、无任何化学成分数据，15 个元素列全为「未提及」。这 5 条可用于 Epit 与工艺分析，**不可用于成分—性能建模**；清洗时不可按 Fe = 100−Σ 反算（会得到 Fe≈100% 的假值）。',
     '【用途】7 个元素有原文出处，可用于成分—性能建模；Fe 原文未单列，记「未提及」，清洗时不要按 Fe = 100−Σ 反算。'),
    (261,
     '【重要】该文在多个 pH 下测点蚀电位，本 4 条对应中性 3.5 wt% NaCl、25℃；碱性 pH 数据未入。',
     '【重要】该文在 pH 3/6/8/10 四组下测点蚀电位；本 4 条（0.22/0.31/0.65/1.25 V）都是 pH 6（准中性）一组，25℃，依据 p4 正文「at pH 6 for example」与 Fig.2。pH 3/8/10 数据未入表。'),
]
# 备注列按块合并（AA261:AA264、AA218:AA222），只在合并区首行写
NOTE_APPEND = {}


def md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def eq(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 1e-9
    return a == b


def apply_edits(ws):
    seen = set()
    for r, col, old, new, why in EDITS:
        c = column_index_from_string(col)
        if (r, c) in seen:
            raise SystemExit('重复格 %s%d' % (col, r))
        seen.add((r, c))
        cur = ws.cell(r, c).value
        if not eq(cur, old):
            raise SystemExit('旧值不符 %s%d：表内=%r 期望=%r（%s）' % (col, r, cur, old, why))
        ws.cell(r, c).value = new
    for r, old, new in NOTE_EDITS:
        cell = ws.cell(r, 27)
        txt = cell.value or ''
        if old not in txt:
            raise SystemExit('备注旧串未命中 row%d：%r' % (r, old[:30]))
        cell.value = txt.replace(old, new)
    for r, extra in NOTE_APPEND.items():
        cell = ws.cell(r, 27)
        txt = cell.value or ''
        if extra.split('】')[0] in txt:
            continue
        cell.value = (txt + '\n' + extra).strip()
    return seen


def snapshot(path):
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    grid = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for c in row:
            grid[(c.row, c.column)] = c.value
    meta = {
        'sheets': wb.sheetnames,
        'max_row': ws.max_row,
        'max_col': ws.max_column,
        'merged': sorted(str(m) for m in ws.merged_cells.ranges),
        'widths': {k: v.width for k, v in ws.column_dimensions.items()},
    }
    return grid, meta


def main():
    commit = '--commit' in sys.argv
    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]
    planned = apply_edits(ws)
    out = SRC if commit else PREVIEW
    if commit:
        if not os.path.exists(BACKUP):
            shutil.copy2(SRC, BACKUP)
        assert md5(SRC) == md5(BACKUP), '备份 MD5 与原件不一致'
        print('备份 ->', BACKUP, md5(BACKUP))
    else:
        os.makedirs(os.path.dirname(PREVIEW), exist_ok=True)
    wb.save(out)
    print('已写 ->', out)

    # 与原件比对（commit 时备份即原件副本）
    ref = BACKUP if commit else SRC
    g0, m0 = snapshot(ref)
    g1, m1 = snapshot(out)

    diffs = {k for k in set(g0) | set(g1) if not eq(g0.get(k), g1.get(k))}
    note_cells = {(r, 27) for r, _, _ in NOTE_EDITS} | {(r, 27) for r in NOTE_APPEND}
    expect = planned | note_cells
    extra = diffs - expect
    missing = expect - diffs
    print('计划改格 %d（含备注 %d），实际差异 %d' % (len(expect), len(note_cells), len(diffs)))
    if extra:
        print('!! 多改的格：', sorted(extra)[:20])
    if missing:
        print('!! 未生效的格：', sorted(missing)[:20])
    for key in ('sheets', 'max_row', 'max_col', 'merged', 'widths'):
        if m0[key] != m1[key]:
            print('!! 结构不一致：', key)
    if not extra and not missing and all(m0[k] == m1[k] for k in m0):
        print('校验通过：仅清单内的格变化，结构与样式元数据一致')
    else:
        raise SystemExit('校验未通过')

    # 打印改动明细
    for r, col, old, new, why in EDITS:
        print('  %s%-3d %r -> %r   [%s]' % (col, r, old, new, why))


if __name__ == '__main__':
    main()
