# -*- coding: utf-8 -*-
r"""第10批流水线离线自测：13 篇假卡片走 patched()->build()->verify()。

只在工作区 temp 的临时目录里造底本/校验件，不碰 0.xlsx / 备份\entry10_20261108。
"""
import importlib.util
import json
import os
import shutil
import sys
import tempfile

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
import entry10_20261108 as e10

FIELDS = e10.FIELDS
m, base_rows, anchor = e10.patched()
data = []
for no in range(147, 160):
    n = 2 if no % 4 == 0 else 1
    rows = []
    for i in range(n):
        r = {k: "未提及" for k in FIELDS}
        r["sample"] = "SELFT10-%d-%d" % (no, i)
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
                 "notes": "数据取自 Table 1。", "flag": "",
                 "rows": [{k: None for k in FIELDS} for _ in rows]})
    for r0, r1 in zip(rows, data[-1]["rows"]):
        for k in FIELDS:
            r1[k] = m.cellval(r0[k], k)

tmp = tempfile.mkdtemp(prefix="entry10_selftest_", dir=os.path.join(ROOT, "temp"))
try:
    base = os.path.join(tmp, "base.xlsx")
    stage = os.path.join(tmp, "stage.xlsx")
    shutil.copy2(e10.BACKUP, base)
    m.BACKUP, m.STAGE = base, stage
    blocks = m.build(data)
    problems = m.verify(data, blocks, stage)
    total = sum(len(p["rows"]) for p in data)
    ws = openpyxl.load_workbook(stage)["Sheet1"]
    old = openpyxl.load_workbook(base)["Sheet1"]
    multi = sum(1 for b in blocks if b["n"] > 1)
    print("selftest10: %d 篇 / %d 条 / 多行块 %d / 底本 %d 行 / 模板块首 r%d"
          % (len(data), total, multi, base_rows, anchor))
    print("stage dims %d x %d / merges %d (base %d)"
          % (ws.max_row, ws.max_column, len(ws.merged_cells.ranges), len(old.merged_cells.ranges)))
    print("首块 r%d: A=%r B=%r Y=%r Z=%r"
          % (blocks[0]["start"], ws.cell(blocks[0]["start"], 1).value,
             ws.cell(blocks[0]["start"], 2).value, ws.cell(blocks[0]["start"], 25).value,
             ws.cell(blocks[0]["start"], 26).value))
    print("末块 r%d: A=%r 编号连续=%s"
          % (blocks[-1]["start"], ws.cell(blocks[-1]["start"], 1).value,
             [ws.cell(b["start"], 1).value for b in blocks] == list(range(147, 160))))
    print("problems = %d" % len(problems))
    for p in problems[:20]:
        print("  -", p)
    sys.exit(1 if problems else 0)
finally:
    shutil.rmtree(tmp, ignore_errors=True)
