# -*- coding: utf-8 -*-
"""建立「已见」索引：待\\ + 文献\\ + 库内题录

目的：新检索池去重。既要按 DOI/PII 精确比，也要按标题模糊比（历史发现 4 例名实不符，
文件名不可信，所以标题必须从 PDF 首页抽）。

产物：temp/lit/oa6/seen.json  [{src, file, doi, pii, title}]
"""
import json
import os
import re

import fitz  # PyMuPDF
import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6")


def doi_from_name(name):
    """从文件名提取 DOI / PII。"""
    doi, pii = "", ""
    m = re.search(r"(1[0-9]\.[0-9]{4,5})[_-]([A-Za-z0-9().\-_]+)", name)
    if m:
        tail = m.group(2).replace("_", ".", 0)
        # 逐段还原：SD_ 系列无 DOI；SP_/CH_/WY_/CR_ 后面接 DOI 尾段
        doi = m.group(1) + "/" + m.group(2)
    m2 = re.search(r"\bS\d{4}\d{4}X\d{5,}\b", name)
    if m2:
        pii = m2.group(0)
    return doi.lower(), pii.lower()


def title_from_pdf(path):
    try:
        doc = fitz.open(path)
        if doc.page_count == 0:
            return ""
        page = doc[0]
        h = page.rect.height
        # 只取上半页，字号最大的几行
        blocks = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for line in b["lines"]:
                txt = "".join(s["text"] for s in line["spans"]).strip()
                if not txt:
                    continue
                size = max(s["size"] for s in line["spans"])
                y = line["bbox"][1]
                if y > h * 0.55:
                    continue
                blocks.append((size, y, txt))
        if not blocks:
            return ""
        blocks.sort(key=lambda x: (-x[0], x[1]))
        top = blocks[0][0]
        lines = [b for b in sorted(blocks, key=lambda x: x[1]) if b[0] >= top - 1.5]
        return " ".join(t[2] for t in lines)[:300]
    except Exception as e:
        return "ERR:" + str(e)[:40]


def main():
    seen = []
    for d in ("待", "文献", "准"):
        p = os.path.join(BASE, d)
        if not os.path.isdir(p):
            continue
        for f in sorted(os.listdir(p)):
            if not f.lower().endswith(".pdf"):
                continue
            doi, pii = doi_from_name(f)
            t = title_from_pdf(os.path.join(p, f))
            seen.append({"src": d, "file": f, "doi": doi, "pii": pii, "title": t})

    # 库内题录
    wb = openpyxl.load_workbook(os.path.join(BASE, "0.xlsx"), data_only=True)
    ws = wb["Sheet1"]
    cur, lib = None, {}
    for r in range(3, ws.max_row + 1):
        p = ws.cell(row=r, column=1).value
        if p not in (None, ""):
            try:
                cur = int(str(p).strip())
            except ValueError:
                pass
        y = ws.cell(row=r, column=25).value
        if y and cur and cur not in lib:
            lib[cur] = str(y)
    for k, v in lib.items():
        m = re.search(r"10\.\d{4,5}/[^\s，,；;]+", v)
        seen.append({
            "src": "0.xlsx",
            "file": "p%02d" % k,
            "doi": (m.group(0).rstrip(".").lower() if m else ""),
            "pii": "",
            "title": v[:300],
        })

    os.makedirs(OUT, exist_ok=True)
    json.dump(seen, open(os.path.join(OUT, "seen.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("已见条目 %d（待/文献/准 PDF %d + 库内题录 %d）"
          % (len(seen), sum(1 for s in seen if s["src"] != "0.xlsx"), len(lib)))
    bad = [s["file"] for s in seen if not s["title"] or s["title"].startswith("ERR:")]
    if bad:
        print("首页抽不出标题 %d 篇：%s" % (len(bad), ", ".join(bad[:8])))


if __name__ == "__main__":
    main()
