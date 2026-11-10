# -*- coding: utf-8 -*-
"""2026-11-08 把判「合规」的 PDF 从 待\\ 移入 准\\（可回滚：写台账 CSV + 失败即停）。
用法: D:\\Python\\python.exe -X utf8 scripts\\move_compliant_20261108.py [dry|run]
"""
import csv
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "待_复筛判定_20261108.csv")
DAI = os.path.join(ROOT, "待")
ZHUN = os.path.join(ROOT, "准")
LOG = os.path.join(ROOT, "准_移入台账_20261108.csv")


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    rows = [r for r in csv.DictReader(open(SRC, encoding="utf-8-sig")) if r["判定"] == "合规"]
    print("合规 %d 篇, mode=%s" % (len(rows), mode))
    os.makedirs(ZHUN, exist_ok=True)
    log = []
    ok = 0
    for r in rows:
        name = r["文件名"]
        s = os.path.join(DAI, name)
        d = os.path.join(ZHUN, name)
        if not os.path.isfile(s):
            print("MISS", name)
            continue
        if os.path.exists(d):
            print("EXISTS", name)
            continue
        if mode == "run":
            shutil.move(s, d)
            ok += 1
        log.append({"文件名": name, "原路径": "待\\" + name, "新路径": "准\\" + name,
                    "判定": r["判定"], "依据": r["依据"], "回滚方式": "从 准\\ 移回 待\\"})
    if mode == "run":
        with open(LOG, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["文件名", "原路径", "新路径", "判定", "依据", "回滚方式"])
            w.writeheader()
            w.writerows(log)
        print("moved", ok, "->", ZHUN)
        print("log", LOG)
        print("待 剩余 pdf", len([f for f in os.listdir(DAI) if f.lower().endswith(".pdf")]))
        print("准 现有 pdf", len([f for f in os.listdir(ZHUN) if f.lower().endswith(".pdf")]))
    else:
        for x in log[:5]:
            print("  would move", x["文件名"])


if __name__ == "__main__":
    main()
