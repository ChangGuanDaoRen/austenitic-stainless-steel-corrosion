# -*- coding: utf-8 -*-
"""只读：按列名对齐，比对 师兄原始示例表 / 0.xlsx 前段(块1-26) / 新批次(块28-63) 各列写法差异。
目的：定位 28-63 批次引入的、不符合既有习惯的标注方式。"""
import openpyxl
from collections import Counter, defaultdict

# ---------- 工具 ----------
def col_index_by_header(ws):
    """返回 {列名: 列号}，两行表头：第1行大类/第2行元素符号，元素列用第2行名。"""
    m = {}
    for c in range(1, ws.max_column + 1):
        h1 = ws.cell(1, c).value
        h2 = ws.cell(2, c).value
        name = str(h2).strip() if h2 else (str(h1).strip() if h1 else f"col{c}")
        m[name] = c
    return m

def block_of_rows(ws):
    """返回 [(起行, 块号)]，A 列非空为块首；返回 row->bid 映射与块起止。"""
    blocks = []
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip() != "":
            blocks.append((r, int(v) if str(v).strip().isdigit() else str(v).strip()))
    row2bid = {}
    for i, (r, bid) in enumerate(blocks):
        end = blocks[i + 1][0] - 1 if i + 1 < len(blocks) else ws.max_row
        for rr in range(r, end + 1):
            row2bid[rr] = bid
    return row2bid

def strvals(ws, row2bid, col, seg):
    """取某列、某分段(front/new/orig)的非数值文本取值集合。"""
    out = []
    for r, bid in row2bid.items():
        if seg == "front" and not (isinstance(bid, int) and bid <= 26):
            continue
        if seg == "new" and not (isinstance(bid, int) and bid >= 28):
            continue
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip() != "":
            out.append(v.strip())
    return out

# ---------- 当前 0.xlsx ----------
cur = openpyxl.load_workbook("0.xlsx", data_only=True)["Sheet1"]
cur_h = col_index_by_header(cur)
r2b = block_of_rows(cur)
front_bids = sorted({b for b in r2b.values() if isinstance(b, int) and b <= 26})
new_bids = sorted({b for b in r2b.values() if isinstance(b, int) and b >= 28})
print("前段块:", front_bids)
print("新批次块:", new_bids)
print("当前表头(列名->列号):", cur_h)

# ---------- 师兄原始示例表（按列名对齐） ----------
orig = openpyxl.load_workbook(r"备份\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx", data_only=True)["Sheet1"]
orig_h = col_index_by_header(orig)
o2b = block_of_rows(orig)
print("原始表头(列名->列号):", orig_h)

# 对齐：只比较两边都有的数据列（排除 Y/Z/文献信息/备注/口径标注/腐蚀pH）
EXCLUDE = {"文献信息", "备注", "口径标注", "腐蚀pH", "文献编号", "样本（原文命名）"}
common = [name for name in cur_h if name in orig_h and name not in EXCLUDE]

print("\n=== 各列：仅新批次出现的文本写法（前段与原始表都没有） ===")
for name in common:
    cc, oc = cur_h[name], orig_h[name]
    fset = set(strvals(cur, r2b, cc, "front"))
    oset = set(strvals(orig, o2b, oc, "orig"))
    nset = set(strvals(cur, r2b, cc, "new"))
    only_new = sorted(nset - fset - oset)
    if only_new:
        # 去重后按频次
        cnt = Counter(strvals(cur, r2b, cc, "new"))
        print(f"[{name}] 仅新批次 {len(only_new)} 种:")
        for s in only_new[:25]:
            print(f"    x{cnt[s]}  {s[:90]}")

print("\n=== 备注(Z) 工作日志式套话词频（36 块新批次） ===")
new_notes = strvals(cur, r2b, cur_h["备注"], "new")
tags = ["【本轮入库", "【提取方式】", "【口径存疑", "【数据不全", "【Fe 记法】", "【用途】",
        "【重要】", "【参比", "【去重】", "【文件名", "【S 列单位", "【Nb00", "复核修订",
        "AA 备注", "未提及", "余量", "WebPlotDigitizer", "取证见"]
for t in tags:
    n = sum(1 for s in new_notes if t in s)
    if n:
        print(f"    {t}: {n} 块")

print("\n=== 备注含『复核修订』的块（前段也有） ===")
front_notes = defaultdict(list)
for r, bid in r2b.items():
    if isinstance(bid, int) and bid <= 26:
        v = cur.cell(r, cur_h["备注"]).value
        if v and "复核修订" in str(v):
            front_notes[bid].append(str(v)[:80])
for bid, ss in sorted(front_notes.items()):
    print(f"    块{bid}: {ss[0]}")
