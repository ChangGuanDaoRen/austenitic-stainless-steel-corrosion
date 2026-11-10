# -*- coding: utf-8 -*-
r"""为 temp\lit\entry\*.txt 生成页级索引：每页标记是否含成分表/电势/介质/工艺关键词。"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
D = os.path.join(WS, "temp", "lit", "entry")

KEYS = {
    "COMP": r"Table\s*1|chemical composition|composition of|成分|mass%|wt\.?%",
    "EPIT": r"pitting potential|Epit|E\s*pit|Eb\b|breakdown potential|击穿电位|点蚀电位",
    "TABLE": r"Table\s*[2-9]|Table\s*1[0-9]",
    "MEDIA": r"3[.]5\s*(wt|mass|%)|3[.]56|0[.]6\s*(M|mol)|NaCl",
    "REF": r"SCE|Ag/AgCl|calomel|甘汞",
    "HT": r"hot.roll|热轧|cold.roll|冷轧|anneal|退火|solution|固溶|as.cast|铸态|quench|水淬",
}

for fn in sorted(os.listdir(D)):
    if not fn.endswith(".txt"):
        continue
    txt = open(os.path.join(D, fn), encoding="utf-8").read()
    pages = re.split(r"=== PAGE (\d+) ===\n", txt)
    print("\n###", fn)
    for i in range(1, len(pages), 2):
        num, body = pages[i], pages[i + 1]
        hits = [k for k, p in KEYS.items() if re.search(p, body, re.I)]
        head = re.sub(r"\s+", " ", body[:150])
        print("  p%-2s [%-28s] %s" % (num, ",".join(hits), head[:110]))
