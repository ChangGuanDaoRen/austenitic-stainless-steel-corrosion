# -*- coding: utf-8 -*-
"""对第9批里 2 篇无文字层的中文扫描件做 RapidOCR，重写 备份\\entry_20261108\\text\\NNN.txt。
只读 准\\*.pdf；只覆盖对应的 text 文件（保留原有题头）。
"""
import json
import os

import pymupdf
from rapidocr_onnxruntime import RapidOCR

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
TEXT_DIR = os.path.join(WORK, "text")
MARK = "================ 论文全文 ================\n"


def ocr_pdf(path, dpi=300):
    ocr = RapidOCR()
    doc = pymupdf.open(path)
    out = []
    tmp = os.path.join(WORK, "_tmp_page.png")
    for pno in range(len(doc)):
        pix = doc[pno].get_pixmap(dpi=dpi)
        pix.save(tmp)
        res, _ = ocr(tmp)
        out.append("===== 第 %d 页（OCR） =====\n" % (pno + 1))
        if res:
            out.append("\n".join(r[1] for r in res))
        out.append("\n")
        print("  page", pno + 1, "/", len(doc), "chars", sum(len(x) for x in out), flush=True)
    doc.close()
    try:
        os.remove(tmp)
    except OSError:
        pass
    return "\n".join(out)


def main():
    with open(os.path.join(WORK, "ocr_needed.json"), encoding="utf-8") as f:
        targets = json.load(f)
    print("待 OCR", len(targets), "篇", flush=True)
    for t in targets:
        src = os.path.join(ROOT, "准", t["file"])
        txt = os.path.join(TEXT_DIR, t["idx"] + ".txt")
        print("OCR", t["idx"], t["file"], flush=True)
        text = ocr_pdf(src)
        with open(txt, encoding="utf-8") as f:
            old = f.read()
        hdr = old.split(MARK)[0]
        with open(txt, "w", encoding="utf-8") as f:
            f.write(hdr + MARK + text)
        print("  chars=%d written=%s" % (len(text), txt), flush=True)


if __name__ == "__main__":
    main()
