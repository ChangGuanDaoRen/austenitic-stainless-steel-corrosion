# -*- coding: utf-8 -*-
"""对 3 篇文字层抽取失败（chars<2000）的 PDF 做 RapidOCR，重写其 digest。

只读 PDF，只覆盖 temp\\judge_20261108\\digest\\pNNN.txt。
用法：D:\\Python\\python.exe -X utf8 scripts\\ocr_failed_20261108.py
"""
import json
import os
import re
import sys

import pymupdf
from rapidocr_onnxruntime import RapidOCR

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JUDGE = os.path.join(ROOT, "temp", "judge_20261108")
MAN = os.path.join(JUDGE, "manifest.json")

KW = re.compile(
    r"NaCl|氯化钠|氯离子|chloride|3\.5|0\.6|35\s*g/L|600\s*mM"
    r"|海水|seawater|人工海水|D1141|模拟"
    r"|pitting|点蚀|孔蚀|点蚀电位|Epit|Eb\b|breakdown"
    r"|SCE|Ag/AgCl|calomel|甘汞|参比|reference"
    r"|固溶|退火|时效|敏化|稳定化|铸|cast|锻|轧|roll"
    r"|duplex|双相|铁素体|马氏体|奥氏体|austenit"
    r"|成分|composition|Cr|Ni|Mo|Mn|Si|Cu"
)


def ocr_pdf(path, dpi=300):
    ocr = RapidOCR()
    doc = pymupdf.open(path)
    out = []
    for pno in range(len(doc)):
        pix = doc[pno].get_pixmap(dpi=dpi)
        tmp = os.path.join(JUDGE, "_tmp_page.png")
        pix.save(tmp)
        res, _ = ocr(tmp)
        if res:
            out.append("\n".join(r[1] for r in res))
        out.append("\n")
    doc.close()
    try:
        os.remove(os.path.join(JUDGE, "_tmp_page.png"))
    except OSError:
        pass
    return "\n".join(out)


def build_digest(rec, text):
    head = text[:3000]
    mid, used, seen = [], 0, set()
    for p in re.split(r"\n\s*\n|\n", text):
        p = p.strip()
        if used >= 13000:
            break
        if not p or not KW.search(p):
            continue
        k = p[:80]
        if k in seen:
            continue
        seen.add(k)
        mid.append(p[:900])
        used += len(p[:900])
    tail = text[-1500:]
    return (
        "【文件】%s\n【目录】%s\n【页数】%s\n【字符数】%s（RapidOCR 300dpi 重抽）\n【DOI】%s\n\n"
        "===== 标题与摘要（OCR 开头） =====\n%s\n\n"
        "===== 关键证据段 =====\n%s\n\n===== 原文结尾 =====\n%s\n"
        % (rec["file"], rec["dir_at_extract"], rec["pages"], len(text), rec.get("doi") or "",
           head, "\n---\n".join(mid), tail)
    )


def main():
    with open(MAN, encoding="utf-8") as f:
        man = json.load(f)
    targets = [m for m in man if (m.get("chars") or 0) < 2000]
    print("待 OCR %d 篇" % len(targets))
    for m in targets:
        src = os.path.join(ROOT, m["dir_at_extract"], m["file"])
        if not os.path.exists(src):
            src = os.path.join(ROOT, "准", m["file"])
        print("OCR:", m["idx"], m["file"], flush=True)
        try:
            text = ocr_pdf(src)
        except Exception as e:
            print("  FAIL", type(e).__name__, e, flush=True)
            continue
        dg = build_digest(m, text)
        with open(os.path.join(JUDGE, "digest", m["idx"] + ".txt"), "w", encoding="utf-8") as f:
            f.write(dg)
        print("  chars=%d digest=%d" % (len(text), len(dg)), flush=True)


if __name__ == "__main__":
    main()
