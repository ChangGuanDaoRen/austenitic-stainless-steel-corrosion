# -*- coding: utf-8 -*-
r"""第 9 批最终闭环校验：0.xlsx 新增块 ↔ 文献\64.pdf..146.pdf ↔ 台账源 MD5。

只读；在 commit 与移动重编号都完成后跑（准\ 原件已移入 文献\，故只比台账源 MD5）。
"""
import csv
import hashlib
import json
import os
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
FINAL = os.path.join(WORK, "final_rows.json")
LEDGER = os.path.join(ROOT, "准_入0与重编号对照_20261108.csv")
TARGET = os.path.join(ROOT, "0.xlsx")
BASE = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
LIT = os.path.join(ROOT, "文献")
OLD_MAX_ROW = 285
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    with open(FINAL, encoding="utf-8") as f:
        data = json.load(f)
    with open(LEDGER, encoding="utf-8-sig") as f:
        ledger = {int(r["新编号"]): r for r in csv.DictReader(f)}
    base = openpyxl.load_workbook(BASE)["Sheet1"]
    cur = openpyxl.load_workbook(TARGET)["Sheet1"]
    problems = []

    total = sum(len(p["rows"]) for p in data)
    if cur.max_row != OLD_MAX_ROW + total:
        problems.append("0.xlsx 行数 %d（应为 %d）" % (cur.max_row, OLD_MAX_ROW + total))
    if cur.max_column != 26:
        problems.append("0.xlsx 列数 %d（应为 26）" % cur.max_column)
    for rr in range(1, OLD_MAX_ROW + 1):
        for c in range(1, 27):
            if base.cell(rr, c).value != cur.cell(rr, c).value:
                problems.append("原有格值差异 r%d c%d" % (rr, c))
                break
        if problems and problems[-1].startswith("原有格值差异"):
            break

    r = OLD_MAX_ROW
    merges = set(str(x) for x in cur.merged_cells.ranges)
    for p in data:
        no, n = p["no"], len(p["rows"])
        start, end = r + 1, r + n
        r = end
        if cur.cell(start, 1).value != no:
            problems.append("文献 %d 块首编号=%r" % (no, cur.cell(start, 1).value))
        if cur.cell(start, 25).value != p["citation"]:
            problems.append("文献 %d 题录不符" % no)
        if cur.cell(start, 26).value != p["notes"]:
            problems.append("文献 %d 备注不符" % no)
        if n > 1:
            for col, letter in ((1, "A"), (25, "Y"), (26, "Z")):
                rng = "%s%d:%s%d" % (letter, start, letter, end)
                if rng not in merges:
                    problems.append("文献 %d 缺合并区 %s" % (no, rng))
        for i, row in enumerate(p["rows"]):
            rr = start + i
            for j, k in enumerate(FIELDS):
                if cur.cell(rr, 2 + j).value != row[k]:
                    problems.append("文献 %d r%d %s 值不符" % (no, rr, k))
                    break
        dst = os.path.join(LIT, "%02d.pdf" % no)
        lg = ledger.get(no)
        if not os.path.exists(dst):
            problems.append("缺 文献\\%02d.pdf" % no)
        elif not lg:
            problems.append("台账缺文献 %d" % no)
        elif md5(dst) != lg["源MD5"]:
            problems.append("文献 %d 的 PDF MD5 与台账不符" % no)

    pdfs = [x for x in os.listdir(LIT) if x.lower().endswith(".pdf")]
    want = set("%02d.pdf" % i for i in range(1, 147))
    got = set(pdfs)
    if got != want:
        problems.append("文献目录不是 01–146 连续：缺 %s / 多 %s"
                        % (sorted(want - got)[:5], sorted(got - want)[:5]))

    print("闭环校验：0.xlsx %d 行 / %d 列；新增 %d 篇 / %d 条；文献目录 %d 篇；问题 %d 项"
          % (cur.max_row, cur.max_column, len(data), total, len(pdfs), len(problems)))
    for x in problems[:40]:
        print("  -", x)
    if len(problems) > 40:
        print("  ... 其余 %d 项省略" % (len(problems) - 40))
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
