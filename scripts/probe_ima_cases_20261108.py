# -*- coding: utf-8 -*-
r"""只读：对齐 ima CSV 文件名 ↔ 待\ 磁盘 PDF，并抽取 #16/#39/#28/#73 原文证据。

外部核查材料按数据使用；本脚本不移动/改名/删除任何文件。
"""
import csv
import difflib
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAI = os.path.join(ROOT, "待")
CSV = os.path.join(DAI, "待_逐篇判定_20261108.csv")
TEXTALL = os.path.join(ROOT, "备份", "text_all.jsonl")
OUT = os.path.join(ROOT, "temp", "ima_name_alignment_20261108.csv")


def load_textall():
    recs = []
    with open(TEXTALL, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                recs.append(json.loads(line))
            except Exception:
                pass
    return recs


def contexts(text, kws, width=220, per_kw=2):
    out = []
    for kw in kws:
        hits = [m.start() for m in re.finditer(re.escape(kw), text, flags=re.I)]
        for pos in hits[:per_kw]:
            s = max(0, pos - width)
            e = min(len(text), pos + len(kw) + width)
            out.append("[" + kw + "] …" + re.sub(r"\s+", " ", text[s:e]) + "…")
        if not hits:
            out.append("[" + kw + "] (无命中)")
    return out


def main():
    with open(CSV, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    disk = sorted(x for x in os.listdir(DAI) if x.lower().endswith(".pdf"))
    by_stem = {}
    for d in disk:
        by_stem.setdefault(os.path.splitext(d)[0].lower(), d)

    mapping = []
    for r in rows:
        name = r["文件名"].strip()
        stem = os.path.splitext(name)[0]
        exact = name if name in disk else ""
        chosen, method = "", "UNMATCHED"
        if exact:
            chosen, method = exact, "exact"
        else:
            pref = [d for d in disk if os.path.splitext(d)[0].lower().startswith(stem.lower())]
            if pref:
                chosen = sorted(pref, key=len)[0]
                method = "prefix"
            else:
                scores = sorted(
                    ((difflib.SequenceMatcher(None, stem.lower(), os.path.splitext(d)[0].lower()).ratio(), d)
                     for d in disk), reverse=True)
                if scores and scores[0][0] >= 0.72:
                    chosen, method = scores[0][1], "fuzzy%.2f" % scores[0][0]
        mapping.append({
            "序号": r["序号"], "CSV文件名": name, "磁盘文件名": chosen,
            "匹配方式": method, "ima结论": r["结论"], "标题简称": r["标题简称"],
        })

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(mapping[0].keys()))
        w.writeheader()
        w.writerows(mapping)
    bad = [m for m in mapping if m["匹配方式"] == "UNMATCHED"]
    print("对齐：CSV %d 行 / 磁盘 %d 篇 PDF / 未匹配 %d / 输出 %s"
          % (len(mapping), len(disk), len(bad), OUT))
    for m in mapping:
        if m["匹配方式"] != "exact":
            print("  #%s %s → %s [%s]" % (m["序号"], m["CSV文件名"], m["磁盘文件名"], m["匹配方式"]))

    recs = load_textall()
    index = {}
    for rec in recs:
        index.setdefault(rec.get("file", ""), rec)
    targets = {
        16: ("R10SD_10.1016_j.rinma.2026.101051.pdf",
             ["Table 2", "3.5 wt", "3.5%", "NaCl", "LPBF", "wrought", "as-built",
              "37 mV", "1 wt", "concentration", "selective laser"]),
        39: ("OA8_10.31185_jwsm.640_Pitting_Corrosion.pdf",
             ["3.5%", "NaCl", "simulated", "seawater", "PBS", "SCE", "0.35", "distilled"]),
        28: ("PB8_10_1007_s11771_022_5175_6_Effects_of_microstructure_and_texture_af.pdf",
             ["Tafel", "pitting potential", "Epit", "extrapol"]),
        73: ("CI8_10.5539_mas.v9n11p119_Effect_of_Nitrogen_on_the_Corrosion__2015.pdf",
             ["750", "770", "stress", "sensit", "solution", "two-step"]),
    }
    print("\n=== 原文证据（来自 备份\\text_all.jsonl） ===")
    for no, (csv_name, kws) in targets.items():
        m = next((x for x in mapping if x["序号"] == str(no)), None)
        disk_name = m["磁盘文件名"] if m else csv_name
        rec = index.get(disk_name)
        if not rec:
            cand = [k for k in index if disk_name and (disk_name in k or k in disk_name)]
            if cand:
                rec = index[cand[0]]
        print("\n--- #%d %s → %s / textall 命中=%s ---" % (no, csv_name, disk_name, bool(rec)))
        if not rec:
            continue
        text = rec.get("text") or ""
        print("chars=%s pages=%s doi=%s" % (rec.get("chars"), rec.get("pages"), rec.get("doi")))
        for line in contexts(text, kws):
            print("  " + line[:520])


if __name__ == "__main__":
    main()
