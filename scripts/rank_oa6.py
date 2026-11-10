# -*- coding: utf-8 -*-
"""OpenAlex 池去重 + 打分 + 排序

输入：temp/lit/oa6/pool.json、seen.json
输出：temp/lit/oa6/ranked.csv（按分降序，含重复标记与命中理由）
      temp/lit/oa6/ranked.md（前 80 条速览，供人工过目）
"""
import csv
import difflib
import json
import os
import re

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6")

# 硬排除：材料体系或介质不是我们要的
BAD_HARD = [
    ("duplex stainless", 3), ("super duplex", 3), ("双相", 3),
    ("martensitic stainless", 3), ("马氏体", 3),
    ("ferritic stainless", 3), ("铁素体不锈钢", 3),
    ("precipitation hardening", 2), ("17-4 ph", 3), ("17-4ph", 3),
    ("high entropy", 3), ("高熵", 3),
    ("additive manufactur", 3), ("selective laser melting", 3), ("增材", 3),
    ("laser powder bed", 3), ("directed energy deposition", 3),
    ("welded", 2), ("weld metal", 2), ("welding", 2), ("焊缝", 3), ("焊接接头", 3),
    ("powder metallurgy", 2), ("spark plasma sintering", 3), ("粉末冶金", 2),
    ("inhibitor", 2), ("缓蚀剂", 3),
    ("h2s", 3), ("硫化氢", 3), ("sulfuric acid", 3), ("h2so4", 3), ("硫酸", 3),
    ("hydrochloric acid", 3), ("盐酸", 3),
    ("cavitation", 2), ("空蚀", 3), ("tribocorrosion", 3), ("摩擦腐蚀", 3),
    ("stress corrosion", 2), ("应力腐蚀", 2),
    ("microbiologically", 2), ("微生物", 2),
    ("carbon steel", 3), ("carbon steel", 3), ("low alloy steel", 3), ("低合金钢", 3),
    ("aluminium", 2), ("aluminum alloy", 3), ("铝合金", 3), ("titanium alloy", 3),
]

# 加分项
GOOD = [
    ("austenitic", 3), ("奥氏体", 3),
    ("pitting potential", 4), ("点蚀电位", 4), ("击穿电位", 4),
    ("breakdown potential", 3), ("critical pitting", 3),
    ("pitting corrosion", 2), ("点蚀", 2),
    ("3.5%", 3), ("3.5 wt", 3), ("3.5wt", 3), ("3.56", 2), ("0.6 mol", 2),
    ("nacl", 2), ("sodium chloride", 2), ("氯化钠", 2),
    ("seawater", 2), ("sea water", 2), ("海水", 2),
    ("cast", 3), ("casting", 3), ("铸造", 3), ("铸态", 3),
    ("cf8m", 5), ("cf3m", 4), ("cf8", 3), ("cd4mcu", 5), ("cn7m", 5),
    ("ck3mcun", 5), ("cf10", 3), ("cf16f", 3), ("s31254", 3), ("n08029", 3),
    ("solution anneal", 2), ("固溶", 2), ("annealed", 2), ("退火", 2),
    ("composition", 1), ("化学成分", 1), ("wt.%", 1),
]


def norm(s):
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", (s or "").lower()).strip()


def load_seen():
    p = os.path.join(OUT, "seen.json")
    seen = json.load(open(p, encoding="utf-8"))
    dois = {s["doi"] for s in seen if s["doi"]}
    piis = {s["pii"] for s in seen if s["pii"]}
    titles = [norm(s["title"]) for s in seen if s["title"] and len(s["title"]) > 12]
    return dois, piis, titles


def main():
    pool = json.load(open(os.path.join(OUT, "pool.json"), encoding="utf-8"))
    dois, piis, titles = load_seen()
    print("池 %d 条；已见 DOI %d / PII %d / 标题 %d" % (len(pool), len(dois), len(piis), len(titles)))
    # 词倒排索引：只与共享 >=1 个实词、长度相近的已见标题做 difflib，避免 100 万次全比对
    global INV
    INV = {}
    for t in set(titles):
        for w in set(x for x in t.split() if len(x) >= 4):
            INV.setdefault(w, set()).add(t)

    rows = []
    for w in pool:
        text = norm(w["title"] + " " + w["abs"])
        score = 0
        why = []
        for kw, v in GOOD:
            if kw in text:
                score += v
                why.append("+" + kw)
        for kw, v in BAD_HARD:
            if kw in text:
                score -= v
                why.append("-" + kw)

        ntitle = norm(w["title"])
        dup = ""
        if w["doi"] and w["doi"] in dois:
            dup = "DOI"
        elif any(p and p in (w["oa_url"] or "").lower() for p in piis):
            dup = "PII-url"
        else:
            if len(ntitle) > 12:
                toks = {w for w in ntitle.split() if len(w) >= 4}
                cand = set()
                for tok in toks:
                    cand |= INV.get(tok, set())
                best = 0
                for t in cand:
                    if abs(len(t) - len(ntitle)) > 45:
                        continue
                    r = difflib.SequenceMatcher(None, ntitle, t).ratio()
                    if r > best:
                        best = r
                if best >= 0.90:
                    dup = "标题%.2f" % best

        rows.append({
            "score": score, "dup": dup,
            "title": w["title"], "year": w["year"], "journal": w["journal"],
            "doi": w["doi"], "cites": w["cites"], "oa": w["oa"], "oa_url": w["oa_url"],
            "hit": ",".join(w["hit"]), "why": " ".join(why),
            "abs": w["abs"][:400].replace("\n", " "),
        })

    rows.sort(key=lambda r: (-(r["score"]), -(r["cites"] or 0)))
    with open(os.path.join(OUT, "ranked.csv"), "w", encoding="utf-8-sig", newline="") as f:
        wtr = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        wtr.writeheader()
        wtr.writerows(rows)

    new = [r for r in rows if not r["dup"]]
    hot = [r for r in new if r["score"] >= 12]
    print("去重后新候选 %d 条；其中高分（>=12）%d 条" % (len(new), len(hot)))

    with open(os.path.join(OUT, "ranked.md"), "w", encoding="utf-8") as f:
        f.write("# OpenAlex 新候选（去重后前 80，按分排序）\n\n")
        f.write("池 %d / 去重后 %d / 高分 %d\n\n" % (len(pool), len(new), len(hot)))
        for i, r in enumerate(hot[:80], 1):
            f.write("%d. **[%d]** %s (%s) *%s*\n   - DOI `%s` | 被引 %s | %s\n   - %s\n\n"
                    % (i, r["score"], r["title"], r["year"], r["journal"],
                       r["doi"] or "—", r["cites"], "OA" if r["oa"] else "闭源",
                       r["abs"][:260]))
    print("→ %s" % os.path.join(OUT, "ranked.md"))


if __name__ == "__main__":
    main()
