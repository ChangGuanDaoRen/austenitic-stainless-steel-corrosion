# -*- coding: utf-8 -*-
"""
2026-11-05  整篇删除文献 6（HN-steel 系列）

依据：用户决定。文献 6 的介质是用 HCl / NaOH 人为调出 pH 1/3/5.24/7/9 的
      3.5% NaCl 体系，属「有其他介质」，与师兄口径（:205「不要有其他介质例如酸啊硫化氢之类」）
      冲突。上一轮只删了 4 条、留下 pH 7 那条，用户现要求整篇删除。

影响面（一并处理，保持 0 / 1 两表与各分节合计自洽）：
  0.xlsx  Sheet1 删 r26（文献6 唯一一行）
  1.xlsx  Sheet1 删 r25
          清洗标记 删 mr24；「0.xlsx行号」列 >26 的全部 −1
          未纳入   增一行「文献6 整篇剔除」；「0.xlsx行号」列 >26 的全部 −1
          统计     全部重算（口径：只统计 1.xlsx 的 251→250 条，不含未纳入）
          热处理分组 删 r29 / r91 / r192；改 r128（≈1080℃ 22→21，去 6）
          工艺分组   删 r97 / r155；改 r4（铸态 46→45）、r20（铸态+固溶 32→31），均去 6

用法：
  python drop_paper6_20261105.py dry      # 只预览，写 temp\\audit0\\*.xlsx
  python drop_paper6_20261105.py write    # 写盘（先备份）
"""
import os
import re
import shutil
import sys
from collections import Counter

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAPER = 6
DEL_ROW_0 = 26          # 0.xlsx 中文献6 的行号
DEL_ROW_1 = 25          # 1.xlsx 中同一行的行号（0 表行号 −1）


# ---------- 通用工具 ----------
def snapshot_and_unmerge(ws):
    old = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
    vals = [(ws.cell(r1, c1).value, r1, c1, r2, c2) for (r1, c1, r2, c2) in old]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    for v, r1, c1, r2, c2 in vals:
        if v is None:
            continue
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = v
    return old


def rebuild_merged(ws, old, deleted_rows):
    def nr(r, is_end=False):
        n = sum(1 for d in deleted_rows if (d <= r if is_end else d < r))
        return r - n

    kept = dropped = 0
    for (r1, c1, r2, c2) in old:
        n1, n2 = nr(r1), nr(r2, True)
        if n2 < n1 or (n1 == n2 and c1 == c2):
            dropped += 1
            continue
        ws.merge_cells(start_row=n1, start_column=c1, end_row=n2, end_column=c2)
        kept += 1
    return kept, dropped


def drop_list(v, target):
    """字符串逗号列表里去掉 target，返回 (新串, 是否变了)"""
    if v is None:
        return v, False
    s = str(v)
    parts = [p.strip() for p in re.split(r'([,，])', s)]
    out, changed = [], False
    for p in parts:
        if p.strip() == str(target):
            changed = True
            continue
        out.append(p)
    ns = ''.join(out)
    ns = re.sub(r',+', ',', ns).strip(',').strip('，')
    return ns, changed


# ---------- 0.xlsx ----------
def do_0(wb0):
    ws = wb0['Sheet1']
    old = snapshot_and_unmerge(ws)
    print('0.xlsx 合并区删前 %d 个；被删行 %d 是否落在某合并区内：%s' % (
        len(old), DEL_ROW_0,
        [x for x in old if x[0] <= DEL_ROW_0 <= x[2]] or '否'))
    ws.delete_rows(DEL_ROW_0, 1)
    kept, dropped = rebuild_merged(ws, old, [DEL_ROW_0])
    print('0.xlsx Sheet1 → %d 行 x %d 列；合并区保留 %d / 移除 %d' % (
        ws.max_row, ws.max_column, kept, dropped))
    return kept, dropped


# ---------- 1.xlsx ----------
def do_1_sheet1(wb1):
    ws = wb1['Sheet1']
    old = snapshot_and_unmerge(ws)
    ws.delete_rows(DEL_ROW_1, 1)
    kept, dropped = rebuild_merged(ws, old, [DEL_ROW_1])
    print('1.xlsx Sheet1 → %d 行 x %d 列；合并区保留 %d / 移除 %d' % (
        ws.max_row, ws.max_column, kept, dropped))


def do_clean_mark(wb1):
    cm = wb1['清洗标记']
    hits = [r for r in range(2, cm.max_row + 1) if cm.cell(r, 3).value == DEL_ROW_0]
    print('清洗标记：命中 0.xlsx行号=%d 的行 %s（应为 1 行）' % (DEL_ROW_0, hits))
    for r in sorted(hits, reverse=True):
        cm.delete_rows(r, 1)
    # 行号列整体 −1（0.xlsx 删了一行）
    n = 0
    for r in range(2, cm.max_row + 1):
        v = cm.cell(r, 3).value
        if isinstance(v, int) and v > DEL_ROW_0:
            cm.cell(r, 3).value = v - 1
            n += 1
        elif isinstance(v, str) and v.isdigit() and int(v) > DEL_ROW_0:
            cm.cell(r, 3).value = int(v) - 1
            n += 1
    print('清洗标记：行号列 >%d 已 −1 的有 %d 行；现 %d 行' % (DEL_ROW_0, n, cm.max_row - 1))


def do_not_included(wb1):
    u = wb1['未纳入']
    n = 0
    for r in range(2, u.max_row + 1):
        v = u.cell(r, 3).value
        if isinstance(v, int) and v > DEL_ROW_0:
            u.cell(r, 3).value = v - 1
            n += 1
        elif isinstance(v, str) and v.isdigit() and int(v) > DEL_ROW_0:
            u.cell(r, 3).value = int(v) - 1
            n += 1
    print('未纳入：行号列 >%d 已 −1 的有 %d 行' % (DEL_ROW_0, n))
    r = u.max_row + 1
    u.cell(r, 1).value = PAPER
    u.cell(r, 2).value = 'HN-steel-4（整篇 5 条全删）'
    u.cell(r, 3).value = DEL_ROW_0
    u.cell(r, 4).value = '文献6 整篇剔除'
    u.cell(r, 5).value = ('介质为 3.5% NaCl 但用 HCl/NaOH 人为调出 pH 1/3/5.24/7/9，'
                          '属「其他介质」，按师兄口径整篇剔除；原 5 条已全部移出，'
                          'PubMed 文献 6.pdf 仍保留在 文献\\ 备用')
    print('未纳入：追加 r%d 文献6 记录' % r)


def cell_sum(ws, r, cols):
    return sum(ws.cell(r, c).value for c in cols if isinstance(ws.cell(r, c).value, (int, float)))


def do_heat_groups(wb1):
    g = wb1['热处理分组']
    # 先改汇总（行号小的先改，删行从大到小）
    for r, want_col2, want_ids in [(128, 21, '1,2,3,12,13,16,18,43')]:
        old2, oldids = g.cell(r, 2).value, g.cell(r, 3).value
        g.cell(r, 2).value = want_col2
        g.cell(r, 3).value = want_ids
        print('热处理分组 r%d：「%s」条数 %s→%s；文献 %s→%s' % (
            r, g.cell(r, 1).value, old2, want_col2, oldids, want_ids))
    for r in sorted([29, 91, 192], reverse=True):
        print('热处理分组：删 r%-3d %s' % (r, [g.cell(r, c).value for c in range(1, 6)]))
        g.delete_rows(r, 1)


def do_proc_groups(wb1):
    g = wb1['工艺分组']
    # (行, 条数列, 文献列, 新条数, 新文献串)
    # A 节列序：成形方式 | 条数 | 文献编号        → 条数 B(2)、文献 C(3)
    # B 节列序：成形方式 | 热处理类型 | 条数 | 文献 → 条数 C(3)、文献 D(4)
    jobs = [(4, 2, 3, 45, '1,2,4,5,12,13,14,15,16,18,49,53'),
            (20, 3, 4, 31, '1,2,12,13,14,15,16,18')]
    for r, c_cnt, c_ids, new_cnt, new_ids in jobs:
        o1, o2 = g.cell(r, c_cnt).value, g.cell(r, c_ids).value
        label = ' | '.join(str(g.cell(r, c).value) for c in (1, 2) if g.cell(r, c).value is not None)
        g.cell(r, c_cnt).value = new_cnt
        g.cell(r, c_ids).value = new_ids
        print('工艺分组 r%d：「%s」条数(%s列) %s→%s；文献(%s列) %s→%s' % (
            r, label[:44], c_cnt, o1, new_cnt, c_ids, o2, new_ids))
    for r in sorted([97, 155], reverse=True):
        print('工艺分组：删 r%-3d %s' % (r, [g.cell(r, c).value for c in range(1, 10)]))
        g.delete_rows(r, 1)


def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else 'dry').lower()
    commit = mode == 'write'

    wb0 = openpyxl.load_workbook(os.path.join(BASE, '0.xlsx'))
    wb1 = openpyxl.load_workbook(os.path.join(BASE, '1.xlsx'))

    print('=' * 60)
    do_0(wb0)
    print('=' * 60)
    do_1_sheet1(wb1)
    do_clean_mark(wb1)
    do_not_included(wb1)
    do_heat_groups(wb1)
    do_proc_groups(wb1)
    print('=' * 60)
    # 统计用内存里的 wb0（已删行）重算，不落中间文件
    do_stats_from(wb1, wb0['Sheet1'])

    if commit:
        bak = os.path.join(BASE, '备份')
        os.makedirs(bak, exist_ok=True)
        for f in ('0.xlsx', '1.xlsx'):
            dst = os.path.join(bak, '%s_删文献6前_20261105.xlsx' % f[:-5])
            if not os.path.exists(dst):
                shutil.copy2(os.path.join(BASE, f), dst)
                print('备份 →', dst)
        wb0.save(os.path.join(BASE, '0.xlsx'))
        wb1.save(os.path.join(BASE, '1.xlsx'))
        print('\n[已写盘] 0.xlsx / 1.xlsx')
    else:
        d = os.path.join(BASE, 'temp', 'audit0')
        os.makedirs(d, exist_ok=True)
        wb0.save(os.path.join(d, '0_删文献6预览.xlsx'))
        wb1.save(os.path.join(d, '1_删文献6预览.xlsx'))
        print('\n[预览] temp\\audit0\\0_删文献6预览.xlsx / 1_删文献6预览.xlsx')


def do_stats_from(wb1, z):
    """统计重算（0.xlsx 用内存对象，避免中间落盘）"""
    ws = wb1['Sheet1']
    cm = wb1['清洗标记']
    u = wb1['未纳入']
    st = wb1['统计']
    total = ws.max_row - 2
    ids = [int(ws.cell(r, 1).value) for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value is not None]
    npaper = len(set(ids))
    cat = Counter(str(u.cell(r, 4).value) for r in range(2, u.max_row + 1))
    no_comp = sum(v for k, v in cat.items() if '无成分' in k)
    no_epit = cat.get('无 Epit 数值', 0)
    zrows = [cm.cell(r, 3).value for r in range(2, cm.max_row + 1) if isinstance(cm.cell(r, 3).value, int)]
    c_un = c_up = c_rng = c_ppm = 0
    for zr in zrows:
        for c in range(3, 17):
            v = z.cell(zr, c).value
            if v == '未提及':
                c_un += 1
            elif isinstance(v, str):
                s = v.strip()
                if s.startswith(('<', '≤')):
                    c_up += 1
                elif 'ppm' in s:
                    c_ppm += 1
                elif re.match(r'^[\d.]+\s*[-–~]\s*[\d.]+$', s):
                    c_rng += 1
    fe = Counter(cm.cell(r, 6).value for r in range(2, cm.max_row + 1))
    off = sum(1 for r in range(3, ws.max_row + 1) if abs(cell_sum(ws, r, list(range(3, 18))) - 100) > 1)
    ref = Counter(cm.cell(r, 7).value for r in range(2, cm.max_row + 1))
    dfl = Counter(cm.cell(r, 9).value for r in range(2, cm.max_row + 1))
    stab = Counter(cm.cell(r, 10).value for r in range(2, cm.max_row + 1))
    ce = Counter(cm.cell(r, 11).value for r in range(2, cm.max_row + 1))

    start_paper = None
    seen = False
    for r in range(2, st.max_row + 1):
        a, b = st.cell(r, 1).value, st.cell(r, 2).value
        key = str(a).strip() if a is not None else ''
        if key == '总条数':
            seen = True
        if seen and key == '1' and isinstance(b, int):
            start_paper = r
            break

    NEW = {
        '总条数': total, '覆盖文献数': npaper,
        '未纳入：文献20（原文无成分）': no_comp, '未纳入：无 Epit 数值': no_epit,
        '未提及/空 → 0': c_un, '<x / ≤x → 取上限': c_up,
        'a-b 区间 → 取中值': c_rng, 'ppm → wt.%（÷1e4）': c_ppm,
        '原文实测': fe.get('实测', 0), '100 − Σ14元素 反算': fe.get('反算', 0),
        '成分和偏离 100 超过 1': off,
        'SCE': ref.get('SCE', 0),
        'Ag/AgCl(未注明->按3M)': ref.get('Ag/AgCl(未注明->按3M)', 0),
        'Ag/AgCl(3M KCl)': ref.get('Ag/AgCl(3M KCl)', 0),
        'Ag/AgCl(饱和KCl)': ref.get('Ag/AgCl(饱和KCl)', 0),
        'Ag/AgCl(饱和,0.1989V vs SHE)': ref.get('Ag/AgCl(饱和,0.1989V vs SHE)', 0),
        'Ag/AgCl(饱和) 原文已换算': ref.get('Ag/AgCl(饱和) 原文已换算', 0),
        '合计（统一记 SCE）': total,
        '默认25': dfl.get('默认25', 0), '默认30': stab.get('默认30', 0),
        '默认铂': ce.get('默认铂', 0),
    }
    changed, drop_rows = [], []
    for r in range(2, st.max_row + 1):
        a, b = st.cell(r, 1).value, st.cell(r, 2).value
        key = str(a).strip() if a is not None else ''
        if start_paper and r >= start_paper:
            if key == str(PAPER):
                drop_rows.append(r); continue
            if key == '合计':
                if b != total:
                    changed.append((r, key, b, total)); st.cell(r, 2).value = total
                continue
            cnt = sum(1 for i in ids if str(i) == key)
            if isinstance(b, int) and b != cnt:
                changed.append((r, key, b, cnt)); st.cell(r, 2).value = cnt
            continue
        if key in NEW and b != NEW[key]:
            changed.append((r, key, b, NEW[key])); st.cell(r, 2).value = NEW[key]
    for r in sorted(drop_rows, reverse=True):
        print('统计：删除文献编号 6 那一行（r%d）' % r)
        st.delete_rows(r, 1)
    for x in changed:
        print('统计：r%-3d %-24s %s → %s' % x)
    print('统计：共改 %d 格，删 %d 行；新总条数 %d / 覆盖文献 %d' % (len(changed), len(drop_rows), total, npaper))


if __name__ == '__main__':
    main()
