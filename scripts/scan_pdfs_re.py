# -*- coding: utf-8 -*-
"""扫描 21 篇文献 PDF，提取参比电极与溶液浓度原文，用于核对表格里的 RE 与浓度记录。"""
import re
import sys
from pathlib import Path

import pymupdf

sys.stdout.reconfigure(encoding="utf-8")
DIR = Path(r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\文献")

PAT = re.compile(
    r"(reference\s+electrode|Ag\s*/\s*AgCl|Ag/AgCl|saturated\s+calomel|\bSCE\b|"
    r"3\.5\s*(wt\.?\s*%|%|mass)|0\.6\s*mol|NaCl\s+solution|natural\s+seawater)",
    re.I,
)

for pdf in sorted(DIR.glob("*.pdf"), key=lambda p: int(p.stem)):
    try:
        doc = pymupdf.open(pdf)
    except Exception as e:  # noqa: BLE001
        print(f"[{pdf.stem}] 打开失败: {e}")
        continue
    hits = []
    for pno, page in enumerate(doc):
        txt = page.get_text()
        if not txt:
            continue
        flat = re.sub(r"\s+", " ", txt)
        for m in PAT.finditer(flat):
            s = max(0, m.start() - 90)
            snippet = flat[s:m.end() + 110].strip()
            hits.append((pno + 1, snippet))
    doc.close()
    # 去重后最多保留 4 条
    seen, out = set(), []
    for pno, sn in hits:
        key = sn[:60]
        if key in seen:
            continue
        seen.add(key)
        out.append((pno, sn))
        if len(out) >= 4:
            break
    print(f"\n===== {pdf.stem}.pdf =====")
    if not out:
        print("  （未检出关键词，可能是扫描件/日文PDF）")
    for pno, sn in out:
        print(f"  p{pno}: {sn}")
