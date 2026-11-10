# -*- coding: utf-8 -*-
"""
2026-11-05  删除「人为加酸/加碱」样本 + 删除 pH 列

依据：师兄原话（对话/总对话_转写.md:205）「找3.5质量分数的氯化钠，
      不要有其他介质例如酸啊硫化氢之类的」。
      只删真正人为调酸/调碱的 5 条：
        文献6  pH 1 / 3 / 5.24 / 9（pH 系列，HN-steel-1/2/3/5）
        文献39 pH 3.0（原文 pH adjusted to 3.0）
      保留纯 3.5% NaCl 自然 pH 的 38 条（文献10 pH5.5-5.8、文献40 pH6.4、文献54 pH6）
      以及 211 条「未提及」。

用法：
  python drop_acidic_ph_20261105.py dry      # 只预览
  python drop_acidic_ph_20261105.py write    # 写盘（先备份）
"""
import io
import os
import shutil
import sys

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
C_PH = 23          # W 列 = 腐蚀pH
DEL_ROWS_0 = [26, 27, 28, 30, 191]     # 原行号（文献6 四条 + 文献39 一条）
DEL_ROWS_1 = [25, 26, 27, 29, 183]     # 原行号（同上，1.xlsx 的行位）


def merged_map(ws):
    """返回 (r,c) -> 合并区左上角的值"""
    fill = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                fill[(r, c)] = v
    return fill


def cell(ws, fill, r, c):
    v = ws.cell(r, c).value
    return fill.get((r, c), v)


def snapshot_and_unmerge(ws):
    """删除前先记下所有合并区坐标并全部解除（openpyxl 的 delete_rows 会自行改写合并区）

    关键：合并区只在左上角存值。若直接删掉块首行，值会一起消失。
    所以先把左上角的值填满整块，再解除合并。
    """
    old = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
    vals = []
    for (r1, c1, r2, c2) in old:
        vals.append((ws.cell(r1, c1).value, r1, c1, r2, c2))
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    # 解除后才能写值（MergedCell.value 只读）
    for v, r1, c1, r2, c2 in vals:
        if v is None:
            continue
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = v
    return old


def rebuild_merged(ws, old, deleted_rows, deleted_cols):
    """按删除后的坐标重建合并区"""
    def nr(r, is_end=False):
        n = sum(1 for d in deleted_rows if (d <= r if is_end else d < r))
        return r - n

    def nc(c, is_end=False):
        n = sum(1 for d in deleted_cols if (d <= c if is_end else d < c))
        return c - n

    kept, dropped = 0, 0
    for (r1, c1, r2, c2) in old:
        # 整块落在被删列上
        if all(c in deleted_cols for c in range(c1, c2 + 1)):
            dropped += 1
            continue
        n1, n2, m1, m2 = nr(r1), nr(r2, True), nc(c1), nc(c2, True)
        if n2 < n1 or m2 < m1 or (n1 == n2 and m1 == m2):
            dropped += 1
            continue
        ws.merge_cells(start_row=n1, start_column=m1, end_row=n2, end_column=m2)
        kept += 1
    return kept, dropped


def plan(wb, fname, del_rows):
    ws = wb['Sheet1']
    fill = merged_map(ws)
    print('== %s  Sheet1 %d 行 x %d 列' % (fname, ws.max_row, ws.max_column))
    for r in del_rows:
        lid = cell(ws, fill, r, 1)
        sample = cell(ws, fill, r, 2)
        ph = ws.cell(r, C_PH).value
        epit = ws.cell(r, 24).value
        print('   删 行%-4d 文%-3s %-26s pH=%-8r Epit=%r' % (r, lid, str(sample)[:26], ph, epit))
    print('   删列 %d（%s），删后 %d 列' % (C_PH, ws.cell(1, C_PH).value, ws.max_column - 1))


def do_sheet1(ws, del_rows):
    old = snapshot_and_unmerge(ws)
    # 倒序逐段删（先大行号，避免位移）
    segs = []
    start = prev = del_rows[0]
    for r in del_rows[1:]:
        if r == prev + 1:
            prev = r
        else:
            segs.append((start, prev - start + 1))
            start = prev = r
    segs.append((start, prev - start + 1))
    for s, n in sorted(segs, reverse=True):
        ws.delete_rows(s, n)
    ws.delete_cols(C_PH, 1)
    kept, dropped = rebuild_merged(ws, old, sorted(del_rows), [C_PH])
    return kept, dropped


def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else 'dry').lower()
    commit = mode == 'write'

    wb0 = openpyxl.load_workbook(os.path.join(BASE, '0.xlsx'))
    wb1 = openpyxl.load_workbook(os.path.join(BASE, '1.xlsx'))

    plan(wb0, '0.xlsx', DEL_ROWS_0)
    plan(wb1, '1.xlsx', DEL_ROWS_1)
    print()

    # ---- 0.xlsx ----
    ws0 = wb0['Sheet1']
    k, d = do_sheet1(ws0, DEL_ROWS_0)
    print('0.xlsx Sheet1 → %d 行 x %d 列（合并区保留 %d / 移除 %d）' % (ws0.max_row, ws0.max_column, k, d))

    # ---- 1.xlsx Sheet1 ----
    ws1 = wb1['Sheet1']
    k, d = do_sheet1(ws1, DEL_ROWS_1)
    print('1.xlsx Sheet1 → %d 行 x %d 列（合并区保留 %d / 移除 %d）' % (ws1.max_row, ws1.max_column, k, d))

    # ---- 1.xlsx 清洗标记：删同样 5 行 + 删 pH 列 ----
    cm = wb1['清洗标记']
    cm_fill = merged_map(cm)
    hits = []
    for r in range(2, cm.max_row + 1):
        rid = cell(cm, cm_fill, r, 3)   # C 列 = 0.xlsx 行号
        try:
            rid = int(rid)
        except (TypeError, ValueError):
            continue
        if rid in DEL_ROWS_0:
            hits.append(r)
    print('1.xlsx 清洗标记：命中 %d 行（应为 5）→ %s' % (len(hits), hits))
    cm_old = snapshot_and_unmerge(cm)
    for r in sorted(hits, reverse=True):
        cm.delete_rows(r, 1)
    # 找 pH 列
    ph_col = None
    for c in range(1, cm.max_column + 1):
        if str(cm.cell(1, c).value).strip() == 'pH':
            ph_col = c
    if ph_col:
        print('1.xlsx 清洗标记：删 pH 列 = 第 %d 列，删后 %d 列' % (ph_col, cm.max_column - 1))
        cm.delete_cols(ph_col, 1)
        rebuild_merged(cm, cm_old, sorted(hits), [ph_col])

    # ---- 1.xlsx 统计：改数字 + 删 pH 相关行 ----
    st = wb1['统计']
    updates = {'总条数': 251, '覆盖文献数': 51}
    to_drop = []
    for r in range(2, st.max_row + 1):
        a, b = st.cell(r, 1).value, st.cell(r, 2).value
        key = str(a).strip() if a is not None else ''
        if key in updates:
            print('统计 行%-3d %-16s %s → %s' % (r, key, b, updates[key]))
            st.cell(r, 2).value = updates[key]
        if key == '默认7':
            to_drop.append(r)
            print('统计 行%-3d 删除「默认7」行（pH 列已删）' % r)
        if key == '6' and str(b) == '5':
            print('统计 行%-3d 文献6 条数 5 → 1' % r)
            st.cell(r, 2).value = 1
        if key == '39' and str(b) == '1':
            to_drop.append(r)
            print('统计 行%-3d 删除文献39 行（该文唯一一条已删）' % r)
        if key == '合计' and str(b) == '256':
            print('统计 行%-3d 合计 256 → 251' % r)
            st.cell(r, 2).value = 251
    for r in sorted(to_drop, reverse=True):
        st.delete_rows(r, 1)

    if commit:
        bak = os.path.join(BASE, '备份')
        os.makedirs(bak, exist_ok=True)
        for f in ('0.xlsx', '1.xlsx'):
            dst = os.path.join(bak, '%s_删pH前_20261105.xlsx' % f.replace('.xlsx', ''))
            if not os.path.exists(dst):
                shutil.copy2(os.path.join(BASE, f), dst)
                print('备份 →', dst)
        wb0.save(os.path.join(BASE, '0.xlsx'))
        wb1.save(os.path.join(BASE, '1.xlsx'))
        print('\n[已写盘] 0.xlsx / 1.xlsx')
    else:
        wb0.save(os.path.join(BASE, 'temp', 'audit0', '0_删pH预览.xlsx'))
        wb1.save(os.path.join(BASE, 'temp', 'audit0', '1_删pH预览.xlsx'))
        print('\n[预览] 已写 temp\\audit0\\0_删pH预览.xlsx / 1_删pH预览.xlsx')


if __name__ == '__main__':
    main()
