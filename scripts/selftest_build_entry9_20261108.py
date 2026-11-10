# -*- coding: utf-8 -*-
r"""build_entry9 的离线自测：用 83 篇假数据在临时目录走一遍 build + verify。

不触碰 0.xlsx / 备份\entry_20261108\0_追加校验9.xlsx / final_rows.json。
"""
import json
import os
import shutil
import sys
import tempfile

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_entry9_20261108 as m

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")

FIELDS = m.FIELDS
data = []
for no in range(64, 147):
    n = 2 if no % 3 == 0 else 1
    rows = []
    for i in range(n):
        r = {k: "未提及" for k in FIELDS}
        r["sample"] = "SELFT-%d-%d" % (no, i)
        r["C"] = "0.03"
        r["Cr"] = "18.2"
        r["Fe"] = "余量"
        r["process"] = "真空感应熔炼铸锭，1100℃固溶2h水淬"
        r["RE"] = "SCE"
        r["Epit"] = "0.321"
        r["temp"] = "25℃"
        rows.append(r)
    data.append({"no": no, "idx": "%03d" % no, "file": "selftest_%d.pdf" % no,
                 "verdict": "合规", "citation": "Selftest %d. Journal. 2026." % no,
                 "notes": "数据取自 Table 1；仅取基准组。" if no % 7 == 0 else "数据取自 Table 1。",
                 "flag": "", "rows": [{k: m.cellval(v, k) for k, v in r.items()} for r in rows]})

tmp = tempfile.mkdtemp(prefix="entry9_selftest_")
try:
    base = os.path.join(tmp, "base.xlsx")
    stage = os.path.join(tmp, "stage.xlsx")
    shutil.copy2(BACKUP, base)
    old_backup, old_stage = m.BACKUP, m.STAGE
    m.BACKUP, m.STAGE = base, stage
    blocks = m.build(data)
    problems = m.verify(data, blocks, stage)
    total = sum(len(p["rows"]) for p in data)
    ws = openpyxl.load_workbook(stage)["Sheet1"]
    old = openpyxl.load_workbook(base)["Sheet1"]
    multi = sum(1 for b in blocks if b["n"] > 1)
    print("selftest: %d 篇 / %d 条 / 多行块 %d" % (len(data), total, multi))
    print("stage dims %d x %d / merges %d (base %d)"
          % (ws.max_row, ws.max_column, len(ws.merged_cells.ranges), len(old.merged_cells.ranges)))
    print("首块 r286: A=%r B=%r Y=%r Z=%r"
          % (ws.cell(286, 1).value, ws.cell(286, 2).value,
             ws.cell(286, 25).value, ws.cell(286, 26).value))
    print("末块 r%d: A=%r 编号连续=%s"
          % (ws.max_row - 1 if False else blocks[-1]["start"], ws.cell(blocks[-1]["start"], 1).value,
             [ws.cell(b["start"], 1).value for b in blocks] == list(range(64, 147))))
    print("problems = %d" % len(problems))
    for p in problems[:15]:
        print("  -", p)
    m.BACKUP, m.STAGE = old_backup, old_stage
    sys.exit(1 if problems else 0)
finally:
    shutil.rmtree(tmp, ignore_errors=True)
