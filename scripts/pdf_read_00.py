#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""00.xlsx 重录用的 PDF 读取工具（只读，不写任何文件）。

用法（工作目录任意，路径写绝对或相对均可）：
  python pdf_read_00.py brief  <pdf>            首页/标题/摘要 + 命中页清单
  python pdf_read_00.py grep   <pdf> <正则>      按行列出命中内容并标页码
  python pdf_read_00.py pages  <pdf> 3-5,8      打印指定页文本（1 起）
  python pdf_read_00.py tables <pdf>            抽取所有页的表格（find_tables）
  python pdf_read_00.py chars  <pdf>            字符统计，判断是否扫描件
"""
import sys, re, os
import pymupdf


def open_doc(p):
    d = pymupdf.open(p)
    return d


def cmd_brief(p):
    d = open_doc(p)
    print(f"[FILE] {os.path.basename(p)}  pages={d.page_count}")
    t0 = d[0].get_text()
    print("[PAGE1]")
    print(t0[:2500])
    keys = re.compile(r"3\.5|NaCl|pitting|breakdown|Table\s*\d|Table\s*[IVX]|composition|wt\.?\s*%|potentiodynamic|polarization", re.I)
    print("[HITS] page:count")
    for i in range(d.page_count):
        c = len(keys.findall(d[i].get_text()))
        if c:
            print(f"  p{i+1}: {c}")
    d.close()


def cmd_grep(p, pat):
    d = open_doc(p)
    rx = re.compile(pat, re.I)
    for i in range(d.page_count):
        lines = d[i].get_text().split("\n")
        for ln in lines:
            s = ln.strip()
            if s and rx.search(s):
                print(f"p{i+1}\t{s[:300]}")
    d.close()


def cmd_pages(p, spec):
    d = open_doc(p)
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            pages.extend(range(int(a), int(b) + 1))
        elif part:
            pages.append(int(part))
    for n in pages:
        if 1 <= n <= d.page_count:
            print(f"===== PAGE {n} =====")
            print(d[n - 1].get_text())
    d.close()


def cmd_tables(p):
    d = open_doc(p)
    for i in range(d.page_count):
        page = d[i]
        try:
            tabs = page.find_tables()
        except Exception as e:
            print(f"[p{i+1}] find_tables error: {e}")
            continue
        for k, tb in enumerate(tabs.tables, 1):
            rows = tb.extract()
            if not rows:
                continue
            print(f"===== PAGE {i+1} TABLE {k} rows={len(rows)} =====")
            for r in rows[:40]:
                cells = [(c or "").replace("\n", " ").strip() for c in r]
                print(" | ".join(cells)[:600])
    d.close()


def cmd_chars(p):
    d = open_doc(p)
    total = 0
    for i in range(d.page_count):
        n = len(d[i].get_text().strip())
        total += n
        print(f"p{i+1}: {n}")
    print(f"[TOTAL] pages={d.page_count} chars={total} per_page={total/max(d.page_count,1):.0f}")
    d.close()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    mode, path = sys.argv[1], sys.argv[2]
    if mode == "brief":
        cmd_brief(path)
    elif mode == "grep":
        cmd_grep(path, sys.argv[3])
    elif mode == "pages":
        cmd_pages(path, sys.argv[3])
    elif mode == "tables":
        cmd_tables(path)
    elif mode == "chars":
        cmd_chars(path)
    else:
        print(__doc__)
        sys.exit(1)
