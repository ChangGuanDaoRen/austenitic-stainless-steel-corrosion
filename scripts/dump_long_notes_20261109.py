# -*- coding: utf-8 -*-
r"""导出 0.xlsx 块级 Z 备注（长度>120）全文，便于人工精简规则取样。

写 备份\清洗1_20261109\long_notes_full.txt，不写 temp。
"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "备份", "清洗1_20261109", "audit_inputs.json")
OUT = os.path.join(ROOT, "备份", "清洗1_20261109", "long_notes_full.txt")


def main():
    with open(SRC, "r", encoding="utf-8") as f:
        d = json.load(f)
    blocks = d["blocks"]
    out = []
    for k in sorted(blocks, key=lambda x: int(x)):
        v = blocks[k]
        if v["z_len"] <= 120:
            continue
        out.append("=" * 100)
        out.append("块 %s | 行 %s | 数据行 %s | 长度 %d" % (k, v["first_row"], v["count"], v["z_len"]))
        out.append("Y: " + v["y"].replace("\n", " / "))
        out.append("Z 原文：")
        out.append(v["z"])
    txt = "\n".join(out) + "\n"
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(txt)
    print("dump ->", OUT, os.path.getsize(OUT), "B")


if __name__ == "__main__":
    main()
