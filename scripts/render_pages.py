# -*- coding: utf-8 -*-
"""渲染指定 PDF 的指定页为 PNG，供人工/OCR 核对表格。
用法: python render_pages.py <PDF相对准目录的文件名> <页码,页码,...> <输出前缀>
"""
import os
import sys

import pymupdf

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
SRC = os.path.join(WS, "准")
DST = os.path.join(WS, "temp", "lit", "entry", "img")
os.makedirs(DST, exist_ok=True)

fn = sys.argv[1]
pages = [int(x) for x in sys.argv[2].split(",")]
prefix = sys.argv[3]
doc = pymupdf.open(os.path.join(SRC, fn))
for p in pages:
    pix = doc[p - 1].get_pixmap(dpi=200)
    out = os.path.join(DST, "%s_p%d.png" % (prefix, p))
    pix.save(out)
    print(out, pix.width, pix.height)
