# -*- coding: utf-8 -*-
"""2026-11-08 只读：导出「存疑」类论文的加长证据，供逐篇人工定级。
输出 temp\\rescreen_20261108\\review_doubt_A.txt / _B.txt（各约一半）
用法: D:\\Python\\python.exe -X utf8 scripts\dump_doubt_20261108.py
"""
import csv
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.join(ROOT, "temp", "rescreen_20261108")
SRC = os.path.join(BASE, "triage_all.csv")


def main():
    rows = [r for r in csv.DictReader(open(SRC, encoding="utf-8-sig"))
            if r["预判定"].startswith("存疑")]
    print("doubt", len(rows))
    half = (len(rows) + 1) // 2
    for i, chunk in enumerate((rows[:half], rows[half:]), start=1):
        path = os.path.join(BASE, "review_doubt_%s.txt" % "AB"[i - 1])
        with open(path, "w", encoding="utf-8") as f:
            for k, r in enumerate(chunk, start=1 + (0 if i == 1 else half)):
                f.write("[%d] %s  (%s页/%s字)\n" % (k, r["文件"], r["页数"], r["字数"]))
                f.write("  标题: %s\n" % r["标题"])
                f.write("  预判: %s | %s\n" % (r["预判定"], r["触发条款"]))
                f.write("  介质: %s\n" % r["介质证据"])
                f.write("  成形: %s\n" % r["成形证据"])
                f.write("  组织: %s\n" % r["组织证据"])
                f.write("  热处理: %s\n" % r["热处理证据"])
                f.write("  电化学: %s\n\n" % r["电化学证据"])
        print("wrote", path, len(chunk))


if __name__ == "__main__":
    main()
