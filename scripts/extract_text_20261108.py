# -*- coding: utf-8 -*-
"""2026-11-08 只读文本提取：待\\ 全部 PDF + 文献\\ 全部 PDF -> 单个 JSONL。
只读 PDF，不改任何原文件。输出 temp\\rescreen_20261108\\text_all.jsonl
用法: D:\\Python\\python.exe -X utf8 scripts\\extract_text_20261108.py
"""
import json
import os
import re
import sys
import unicodedata

import fitz

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "temp", "rescreen_20261108", "text_all.jsonl")


def nfkc(s):
    return unicodedata.normalize("NFKC", s or "")


def squeeze(s):
    s = nfkc(s).replace("\u00ad", "")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


def read_pdf(path):
    doc = fitz.open(path)
    pages = []
    for p in doc:
        try:
            pages.append(p.get_text("text"))
        except Exception:
            pages.append("")
    npages = doc.page_count
    doc.close()
    return squeeze("\n".join(pages)), npages


def doi_of(txt):
    m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", txt)
    return m.group(0).rstrip(".,;)").lower() if m else ""


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    jobs = []
    for sub, tag in (("待", "DAI"), ("文献", "WEN")):
        d = os.path.join(ROOT, sub)
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".pdf"):
                jobs.append((tag, sub, f, os.path.join(d, f)))
    print("pdf count", len(jobs), flush=True)
    done = 0
    with open(OUT, "w", encoding="utf-8") as w:
        for tag, sub, name, path in jobs:
            try:
                txt, npages = read_pdf(path)
            except Exception as e:
                txt, npages = "", 0
                print("ERR", name, repr(e), flush=True)
            rec = {
                "tag": tag,
                "dir": sub,
                "file": name,
                "pages": npages,
                "chars": len(txt),
                "doi": doi_of(txt[:8000]),
                "head": txt[:1200],
                "text": txt,
            }
            w.write(json.dumps(rec, ensure_ascii=False) + "\n")
            done += 1
            if done % 20 == 0:
                w.flush()
                print("done", done, flush=True)
    print("wrote", OUT, done, flush=True)


if __name__ == "__main__":
    main()
