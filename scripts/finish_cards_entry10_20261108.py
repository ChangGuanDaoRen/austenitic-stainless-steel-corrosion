# -*- coding: utf-8 -*-
r"""补完第10批卡片：用 manifest 的精确文件名修 156，并人工按全文表 1 补 159。

只读 manifest/text；只写 备份\entry10_20261108\cards\156.json 与 159.json。
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry10_20261108")
CARD_DIR = os.path.join(WORK, "cards")
MANIFEST = os.path.join(WORK, "manifest.json")
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]

with open(MANIFEST, encoding="utf-8") as f:
    manifest = json.load(f)
by_no = {m["newno"]: m for m in manifest}

# 1) 156：卡片 file 多了一个连字符，改用 manifest/disk 的精确字符串。
p156 = os.path.join(CARD_DIR, "156.json")
with open(p156, encoding="utf-8") as f:
    c156 = json.load(f)
old156 = c156.get("file")
c156["file"] = by_no[156]["file"]
with open(p156, "w", encoding="utf-8") as f:
    json.dump(c156, f, ensure_ascii=False, indent=2)
print("[fix156] 原 file =", repr(old156))
print("[fix156] 新 file =", repr(c156["file"]))
print("[fix156] 与 manifest/disk 一致 =", c156["file"] == by_no[156]["file"])

# 2) 159：按 text\159.txt 的 1 实验部分 + Table 1 逐字段照录。
BASE_ELEMS = {
    "C": "0.14×10-1%",
    "Si": "0.60%",
    "N": "未提及",
    "P": "0.13×10-1%",
    "S": "0.07%",
    "Cu": "未提及",
    "Mn": "0.80%",
    "Cr": "17.14%",
    "Mo": "未提及",
    "Ni": "12.58%",
    "Ti": "未提及",
    "Nb": "未提及",
    "Co": "未提及",
    "Sn": "未提及",
    "Fe": "余量",
}
BASE = {
    "process": "未提及",
    "WE": "10 mm×10 mm×3 mm 片状，环氧树脂密封，留 1 mm²，SiC 2000#打磨",
    "RE": "SCE",
    "CE": "大块石墨",
    "stable": "阴极极化-1.3 V 3 min；OCP 1 800 s 后稳定",
}
TEMPS = ["5°C", "10°C", "15°C", "25°C"]
EPITS = ["892.3 mV", "864.4 mV", "365.7 mV", "-262.4 mV"]
rows = []
for temp, epit in zip(TEMPS, EPITS):
    r = {"sample": "316L-" + temp.replace("°C", "℃")}
    r.update(BASE_ELEMS)
    r.update(BASE)
    r["temp"] = temp
    r["Epit"] = epit
    rows.append({k: r[k] for k in FIELDS})

c159 = {
    "newno": 159,
    "file": by_no[159]["file"],
    "citation": "林海波, 张巨伟, 李思雨. 温度对316L不锈钢在3.5%NaCl溶液中腐蚀行为的影响[J]. 辽宁石油化工大学学报, 2019, 39(2): 54-58. DOI: 10.3969/j.issn.1672⁃6952.2019.02.010",
    "notes": "成分取自1实验部分；Epit取自Table 1（Eb/mV vs SCE）；参比句有模板残留。",
    "rows": rows,
}
p159 = os.path.join(CARD_DIR, "159.json")
with open(p159, "w", encoding="utf-8") as f:
    json.dump(c159, f, ensure_ascii=False, indent=2)

with open(p159, encoding="utf-8") as f:
    check = json.load(f)
print("[write159] file =", repr(check["file"]))
print("[write159] rows =", len(check["rows"]), " / 23键 =",
      all(list(r.keys()) == FIELDS and all(isinstance(v, str) and v.strip() for v in r.values())
          for r in check["rows"]))
print("[write159] flag =", repr(check.get("flag")), " / notes 字数 =", len(check["notes"]))
print("[done] 156 与 159 已按 manifest 精确写回")
