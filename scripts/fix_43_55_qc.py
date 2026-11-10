# -*- coding: utf-8 -*-
r"""按质检结论修订 43-55 块：Fe 改「余量」+ 三处备注加注（2026-11-04 19:3x）。

修订内容
  1. 文献 43/47/48/54 共 21 行，Fe（第 17 列）由「未提及」改「余量」——
     这四篇的成分表已把合金元素列全、未单列 Fe，钢基体按余量记，与全表 205/206 行口径一致。
     **文献 45 不改**：它 15 个元素全空，改 Fe 会变成 Fe=100%。
  2. 文献 43 的 AA 备注追加：Nb00@1300℃ 的 0.240 是正文推断值（Fig.6 红线无 0 wt% 点）。
  3. 文献 45 的 AA 备注追加：无成分数据，不可用于成分—性能建模。
  4. 文献 54 的 S 列「13 ppm」字符串**按原样保留**，不换算（换算是替作者改单位）。

用法：
  python scripts/fix_43_55_qc.py            只校验，写 temp\lit\entry2\0_质检修订_校验.xlsx
  python scripts/fix_43_55_qc.py --commit   校验通过后写回 0.xlsx
"""
import hashlib
import os
import shutil
import sys

import openpyxl

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
BACKUP = os.path.join(WS, "备份", "0_基准_质检修订前_20261104.xlsx")
TARGET = os.path.join(WS, "0.xlsx")
STAGE = os.path.join(WS, "temp", "lit", "entry2", "0_质检修订_校验.xlsx")

COL_FE, COL_AA = 17, 27
NROW, NCOL = 266, 27

# 块首行 -> (文献号, 条数, 是否改 Fe)
BLOCKS = {209: (43, 8, True), 218: (45, 5, False), 226: (47, 5, True),
          231: (48, 4, True), 261: (54, 4, True)}

NOTE_FE = ("\n【Fe 记法】原文成分表已列全合金元素但未单列 Fe，钢基体按「余量」记，"
           "与本表其余文献口径一致；未做 Fe = 100−Σ 的数值折算。")
NOTE_43 = ("\n【Nb00@1300℃ 说明】Fig.6 的 1300℃ 曲线只有 0.23/0.48/0.99 三个数据点，0 wt% 处无独立点；"
           "正文写「raising STT ... nearly had no influence on the stable pitting for Nb-free ones」，"
           "故该条按 1100℃ 同值 0.240 V 记录，属正文推断值而非图上读出，引用时请注意。"
           "另：1300℃ 三点读图精度约 ±0.02 V。")
NOTE_45 = ("\n【用途限制】本篇通篇只给牌号 type 304、无任何化学成分数据，15 个元素列全为「未提及」。"
           "这 5 条可用于 Epit 与工艺分析，**不可用于成分—性能建模**；"
           "清洗时不可按 Fe = 100−Σ 反算（会得到 Fe≈100% 的假值）。")


def md5(p):
    with open(p, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def sig(c):
    f, a, b = c.font, c.alignment, c.border
    return (f.name, f.sz, f.b, f.i, a.horizontal, a.vertical, a.wrap_text,
            c.number_format, b.left.style, b.right.style, b.top.style, b.bottom.style)


def plan():
    """返回预期改动的格集合与明细，便于校验。"""
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if (ws.max_row, ws.max_column) != (NROW, NCOL):
        raise SystemExit("原件结构变了：%d 行 / %d 列" % (ws.max_row, ws.max_column))
    expect, detail = set(), []
    for anchor, (no, n, fix_fe) in BLOCKS.items():
        if ws.cell(anchor, 1).value != no:
            raise SystemExit("块首行 %d 的编号是 %r，应为 %d" % (anchor, ws.cell(anchor, 1).value, no))
        if fix_fe:
            for r in range(anchor, anchor + n):
                if ws.cell(r, COL_FE).value != "未提及":
                    raise SystemExit("文献 %d r%d 的 Fe 不是「未提及」：%r" % (no, r, ws.cell(r, COL_FE).value))
                expect.add((r, COL_FE))
            detail.append("文献 %d：%d 行 Fe 未提及 → 余量" % (no, n))
    for anchor in BLOCKS:
        expect.add((anchor, COL_AA))
    return wb, expect, detail


def build(wb):
    ws = wb["Sheet1"]
    for anchor, (no, n, fix_fe) in BLOCKS.items():
        if fix_fe:
            for r in range(anchor, anchor + n):
                c = ws.cell(r, COL_FE)
                keep = c._style
                c.value = "余量"
                c._style = keep
            ws.cell(anchor, COL_AA).value = (ws.cell(anchor, COL_AA).value or "") + NOTE_FE
    ws.cell(209, COL_AA).value += NOTE_43
    ws.cell(218, COL_AA).value += NOTE_45
    wb.save(STAGE)


def verify(expect):
    problems = []
    old = openpyxl.load_workbook(BACKUP)["Sheet1"]
    new = openpyxl.load_workbook(STAGE)["Sheet1"]
    if (new.max_row, new.max_column) != (NROW, NCOL):
        problems.append("行/列数变了：%d/%d" % (new.max_row, new.max_column))
    actual = set()
    for r in range(1, NROW + 1):
        for c in range(1, NCOL + 1):
            a, b = old.cell(r, c), new.cell(r, c)
            if a.value != b.value:
                actual.add((r, c))
            elif sig(a) != sig(b):
                problems.append("样式变了：r%d c%d" % (r, c))
    if actual != expect:
        problems.append("改动格不符：多改 %r / 漏改 %r" % (sorted(actual - expect), sorted(expect - actual)))
    for r in [x[0] for x in expect if x[1] == COL_FE]:
        if new.cell(r, COL_FE).value != "余量":
            problems.append("r%d Fe 未写入" % r)
    for anchor in BLOCKS:
        if "【" not in (new.cell(anchor, COL_AA).value or ""):
            problems.append("r%d 备注未追加" % anchor)
    om = set(str(x) for x in old.merged_cells.ranges)
    nm = set(str(x) for x in new.merged_cells.ranges)
    if om != nm:
        problems.append("合并区变了")
    if {k: v.width for k, v in old.column_dimensions.items()} != {k: v.width for k, v in new.column_dimensions.items()}:
        problems.append("列宽变了")
    return problems


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if not os.path.exists(BACKUP):
        raise SystemExit("缺备份：" + BACKUP)
    wb, expect, detail = plan()
    build(wb)
    problems = verify(expect)
    for d in detail:
        print("  " + d)
    print("  文献 43 备注追加 Nb00@1300℃ 推断说明；文献 45 备注追加「不可用于成分建模」")
    print("预期改动 %d 格（Fe %d 格 + 备注 %d 格）" % (len(expect), len(expect) - len(BLOCKS), len(BLOCKS)))
    print("校验件：" + STAGE)
    if problems:
        print("校验未通过：")
        for x in problems[:15]:
            print("  -", x)
        raise SystemExit(2)
    print("校验通过：除预期格外逐格零差异 / 合并区与列宽不变")
    if "--commit" in sys.argv:
        if md5(TARGET) not in (md5(BACKUP), md5(STAGE)):
            raise SystemExit("0.xlsx 已被改动，停止")
        shutil.copy2(STAGE, TARGET)
        print("[commit] 已写回 0.xlsx，MD5=%s" % md5(TARGET))
    else:
        print("[dry] 未改 0.xlsx")


if __name__ == "__main__":
    main()
