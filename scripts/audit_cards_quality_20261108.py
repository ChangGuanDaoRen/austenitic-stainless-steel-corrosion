# -*- coding: utf-8 -*-
r"""第 9 批卡片的数据质量体检（只读）。

在 83 张卡片齐备后跑，输出 temp\entry9_cards_quality_20261108.csv 与控制台摘要。
目的是把「字段齐但内容空」的卡片挑出来人工复核，避免把占位行当成真实数据入库。
"""
import csv
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
CARDS = os.path.join(WORK, "cards")
OUT = os.path.join(ROOT, "temp", "entry9_cards_quality_20261108.csv")
ELEMS = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
         "Ti", "Nb", "Co", "Sn", "Fe"]


def main():
    rows = []
    for no in range(64, 147):
        path = os.path.join(CARDS, "%03d.json" % no)
        rec = {"编号": no, "卡片": "缺", "行数": 0, "成分全缺行数": 0,
               "Epit缺失行数": 0, "参比缺失行数": 0, "工艺缺失行数": 0,
               "元素未提及率": "", "备注字数": 0, "题录字数": 0,
               "风险标签": ""}
        if not os.path.exists(path):
            rec["风险标签"] = "缺卡片"
            rows.append(rec)
            continue
        try:
            with open(path, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            rec["风险标签"] = "JSON错误:" + type(e).__name__
            rows.append(rec)
            continue
        data = c.get("rows") or []
        rec["卡片"] = "有"
        rec["行数"] = len(data)
        rec["备注字数"] = len(str(c.get("notes") or ""))
        rec["题录字数"] = len(str(c.get("citation") or ""))
        miss_total, miss_elem = 0, 0
        tags = []
        for r in data:
            vals = {k: str(r.get(k, "")).strip() for k in r}
            if all(vals.get(k, "") in ("", "未提及", "N/A", "-") for k in ELEMS):
                rec["成分全缺行数"] += 1
            if vals.get("Epit", "") in ("", "未提及", "N/A", "-"):
                rec["Epit缺失行数"] += 1
            if vals.get("RE", "") in ("", "未提及", "N/A", "-"):
                rec["参比缺失行数"] += 1
            if vals.get("process", "") in ("", "未提及", "N/A", "-"):
                rec["工艺缺失行数"] += 1
            for k in ELEMS:
                miss_total += 1
                if vals.get(k, "") in ("", "未提及", "N/A", "-"):
                    miss_elem += 1
        rec["元素未提及率"] = "%.0f%%" % (100.0 * miss_elem / miss_total) if miss_total else "-"
        if c.get("flag"):
            tags.append("flag")
        if rec["成分全缺行数"]:
            tags.append("成分全缺%d行" % rec["成分全缺行数"])
        if rec["Epit缺失行数"]:
            tags.append("Epit缺%d行" % rec["Epit缺失行数"])
        if rec["参比缺失行数"]:
            tags.append("参比缺%d行" % rec["参比缺失行数"])
        if rec["工艺缺失行数"]:
            tags.append("工艺缺%d行" % rec["工艺缺失行数"])
        if rec["备注字数"] > 120:
            tags.append("备注偏长")
        if rec["行数"] == 0:
            tags.append("零数据行")
        rec["风险标签"] = "；".join(tags) if tags else "ok"
        rows.append(rec)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    total_rows = sum(x["行数"] for x in rows)
    risky = [x for x in rows if x["风险标签"] != "ok"]
    print("卡片 %d 篇 / 数据 %d 条 / 需关注 %d 篇" % (len(rows), total_rows, len(risky)))
    print("成分全缺行 %d / Epit缺 %d / 参比缺 %d / 工艺缺 %d"
          % (sum(x["成分全缺行数"] for x in rows), sum(x["Epit缺失行数"] for x in rows),
             sum(x["参比缺失行数"] for x in rows), sum(x["工艺缺失行数"] for x in rows)))
    for x in risky:
        print("  - %d | %d行 | %s | %s" % (x["编号"], x["行数"], x["元素未提及率"], x["风险标签"]))
    print("明细：" + OUT)


if __name__ == "__main__":
    main()
