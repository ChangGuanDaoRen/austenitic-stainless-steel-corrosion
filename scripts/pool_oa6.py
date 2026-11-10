# -*- coding: utf-8 -*-
"""OpenAlex 题录检索（2026-11-05 第 6 轮检索）

用途：按用户「成形工艺=铸造即可，排除增材/焊接/粉末冶金」的口径，重新做一轮
覆盖式题录检索，重点是历史上没走过的维度：
  - 铸造牌号定向（CF8M / CD4MCu / CN7M / CK3MCuN / CF3MN / HK40 ...）
  - 中文关键词（点蚀电位/击穿电位/铸造奥氏体不锈钢）
  - 海水介质（师兄 :91-92 天然海水可进主库）
  - 前向引文（库内 52 篇的被引文献，历史只做过参考文献后向回溯）

产物：temp/lit/oa6/raw/<slug>.json（原样存，便于复查命中了什么）
      temp/lit/oa6/pool.json（去重合并后的候选池）

用法：python pool_oa6.py scan [--only slug1,slug2]
      python pool_oa6.py pool
"""
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6")
RAW = os.path.join(OUT, "raw")
UA = "cast-ASS-pitting-review/1.0 (mailto:3144668926@qq.com)"

SELECT = ",".join([
    "id", "doi", "title", "publication_year", "type", "cited_by_count",
    "abstract_inverted_index", "primary_location", "biblio",
    "open_access", "authorships",
])

# (slug, 检索式, 说明)
QUERIES = [
    # A 组：介质+电位（主检索式）
    ("q01", "pitting potential austenitic stainless steel sodium chloride", "主检索式"),
    ("q02", "pitting corrosion austenitic stainless steel 3.5% NaCl", "3.5% 显式"),
    ("q03", "critical pitting temperature austenitic stainless steel chloride", "CPT"),
    ("q04", "pitting potential stainless steel seawater chloride", "天然海水"),
    # B 组：成形工艺（铸造/铸造牌号）
    ("q05", "pitting corrosion cast austenitic stainless steel chloride", "铸造"),
    ("q06", "cast stainless steel pitting potential solution annealing", "铸造+固溶"),
    ("q07", "CF8M pitting corrosion chloride", "CF8M"),
    ("q08", "CD4MCu corrosion pitting chloride", "CD4MCu"),
    ("q09", "CN7M alloy 20 corrosion pitting", "CN7M/20 号合金"),
    ("q10", "CK3MCuN pitting corrosion", "254SMO 铸造"),
    ("q11", "CF3M CF8 microstructure corrosion chloride", "CF3M/CF8"),
    ("q12", "investment casting stainless steel pitting corrosion", "熔模铸造"),
    # C 组：牌号/组织维度（锻轧也收）
    ("q13", "super austenitic stainless steel pitting corrosion chloride 254SMO", "超级奥氏体"),
    ("q14", "high nitrogen austenitic stainless steel pitting potential chloride", "高氮"),
    ("q15", "cold rolled annealed austenitic stainless steel pitting resistance chloride", "冷轧+退火"),
    ("q16", "hot rolled solution treated austenitic stainless steel pitting chloride", "热轧+固溶"),
    ("q17", "delta ferrite austenitic stainless steel pitting corrosion", "δ 铁素体"),
    ("q18", "molybdenum nitrogen synergy pitting austenitic stainless steel", "Mo/N 协同"),
    # D 组：中文
    ("q19", "点蚀电位 奥氏体不锈钢 氯化钠", "中文点蚀电位"),
    ("q20", "铸造奥氏体不锈钢 点蚀", "中文铸造"),
    ("q21", "不锈钢 击穿电位 3.5%NaCl", "中文击穿电位"),
    ("q22", "铸造不锈钢 耐蚀性 氯化钠溶液", "中文铸造耐蚀"),
    # E 组：11-05 第二轮补充（铸造牌号 / 中文期刊 / NACE 会议 / 引文向）
    ("q23", "cast superaustenitic stainless steel pitting chloride seawater", "铸造超级奥氏体"),
    ("q24", "solution annealed cast stainless steel pitting potential delta ferrite", "δ铁素体+固溶"),
    ("q25", "breakdown potential stainless steel 3.5 percent sodium chloride alloying", "NACE老论文"),
    ("q26", "CORROSION journal pitting potential austenitic stainless steel chloride", "NACE/CORROSION"),
    ("q27", "奥氏体不锈钢 3.5% NaCl 点蚀电位 固溶 成分", "中文固溶"),
    ("q28", "铸造不锈钢 点蚀 海水 固溶 腐蚀电位", "中文铸造海水"),
    ("q29", "CF8 CF8M CF3M stainless steel casting corrosion electrochemical", "CF系列"),
    ("q30", "high nitrogen austenitic stainless steel pitting potential chloride solution treatment", "高氮固溶"),
]

SLEEP = 5.0


def fetch(url, tries=6):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=40) as r:
                return json.loads(r.read())
        except Exception as e:
            code = getattr(e, "code", "")
            wait = 8 * (i + 1) if code == 429 else 3 * (i + 1)
            print("    retry %d (%s) wait %ds" % (i + 1, code or type(e).__name__, wait))
            time.sleep(wait)
    return None


def scan(only=None):
    os.makedirs(RAW, exist_ok=True)
    for slug, q, note in QUERIES:
        if only and slug not in only:
            continue
        path = os.path.join(RAW, slug + ".json")
        if os.path.exists(path):
            d = json.load(open(path, encoding="utf-8"))
            print("%s 跳过（已有 %d 条）%s" % (slug, len(d.get("results", [])), note))
            continue
        url = ("https://api.openalex.org/works?search="
               + urllib.parse.quote(q)
               + "&per-page=200&select=" + SELECT)
        d = fetch(url)
        if d is None:
            print("%s 失败" % slug)
            continue
        json.dump(d, open(path, "w", encoding="utf-8"), ensure_ascii=False)
        print("%s %-22s 命中 %5d 条，取回 %3d  (%s)"
              % (slug, note, d["meta"]["count"], len(d["results"]), q))
        time.sleep(SLEEP)


def abstract(work):
    idx = work.get("abstract_inverted_index")
    if not idx:
        return ""
    pos = {}
    for w, ps in idx.items():
        for p in ps:
            pos[p] = w
    return " ".join(pos[k] for k in sorted(pos))


def pool():
    works = {}
    for slug, q, note in QUERIES:
        path = os.path.join(RAW, slug + ".json")
        if not os.path.exists(path):
            continue
        d = json.load(open(path, encoding="utf-8"))
        for w in d.get("results", []):
            wid = w["id"]
            if wid in works:
                works[wid]["hit"].append(slug)
                continue
            loc = w.get("primary_location") or {}
            src = (loc.get("source") or {}).get("display_name", "")
            works[wid] = {
                "id": wid,
                "doi": (w.get("doi") or "").replace("https://doi.org/", "").lower(),
                "title": (w.get("title") or "").strip(),
                "year": w.get("publication_year"),
                "type": w.get("type"),
                "cites": w.get("cited_by_count"),
                "journal": src,
                "oa": (w.get("open_access") or {}).get("is_oa"),
                "oa_url": (w.get("open_access") or {}).get("oa_url") or "",
                "abs": abstract(w)[:1600],
                "hit": [slug],
            }
    out = os.path.join(OUT, "pool.json")
    json.dump(list(works.values()), open(out, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("池内 %d 条（去重后）→ %s" % (len(works), out))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "scan"
    if cmd == "scan":
        only = None
        if "--only" in sys.argv:
            only = set(sys.argv[sys.argv.index("--only") + 1].split(","))
        scan(only)
    elif cmd == "pool":
        pool()
    else:
        print("用法: scan [--only slug] | pool")
