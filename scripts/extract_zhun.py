# -*- coding: utf-8 -*-
r"""把 准\ 内 15 篇 PDF 逐页抽成文本，供录入 0.xlsx 时逐条回原文核定。

输出：temp\lit\entry\<文件名主干>.txt，页间以 === PAGE n === 分隔。
仅做文本抽取（NFKC 归一），不改任何 PDF。
"""
import os
import re
import sys
import unicodedata

import pymupdf

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
SRC = os.path.join(WS, "准")
DST = os.path.join(WS, "temp", "lit", "entry")
os.makedirs(DST, exist_ok=True)


def main():
    names = sorted(f for f in os.listdir(SRC) if f.lower().endswith(".pdf"))
    for fn in names:
        stem = os.path.splitext(fn)[0]
        out = os.path.join(DST, stem + ".txt")
        doc = pymupdf.open(os.path.join(SRC, fn))
        parts = []
        for i, page in enumerate(doc):
            txt = unicodedata.normalize("NFKC", page.get_text())
            parts.append("=== PAGE %d ===\n%s" % (i + 1, txt))
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(parts))
        total = sum(len(p) for p in parts)
        print("%-76s %2d页 %7d字符 -> %s" % (fn[:76], doc.page_count, total, os.path.basename(out)))
    print("共", len(names), "篇")


if __name__ == "__main__":
    main()
