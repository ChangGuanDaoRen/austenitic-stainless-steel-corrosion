# -*- coding: utf-8 -*-
import csv
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry10_20261108")
ZHUN = os.path.join(ROOT, "准")
with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
    manifest = json.load(f)
by_no = {m["newno"]: m for m in manifest}
with open(os.path.join(WORK, "numbering.csv"), encoding="utf-8-sig", newline="") as f:
    num = {int(r["新编号"]): r["原文件名"] for r in csv.DictReader(f)}
with open(os.path.join(WORK, "cards", "156.json"), encoding="utf-8") as f:
    card = json.load(f)
disk = sorted(x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf"))
hit = [x for x in disk if "Tensile" in x]
for label, s in [("manifest", by_no[156]["file"]), ("numbering", num[156]),
                 ("card", card.get("file")), ("disk", hit[0] if hit else "")]:
    print(label, "ascii=", ascii(s))
    print(label, "codes=", " ".join("%04X" % ord(x) for x in s))
print("counts:", len(hit), "manifest==disk", by_no[156]["file"] == (hit[0] if hit else None))
