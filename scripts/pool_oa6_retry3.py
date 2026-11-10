# -*- coding: utf-8 -*-
"""q23/q24/q30 补跑：OpenAlex 匿名 search 集群 429，改用 filter + 长退避"""
import os, json, time, urllib.request, urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "temp", "lit", "oa6", "raw")
UA = "cast-ASS/1.0 (mailto:3144668926@qq.com)"
SELECT = "id,doi,title,publication_year,type,cited_by_count,primary_location,open_access,best_oa_location,abstract_inverted_index,authorships"
QUERIES = [
    ("q23", "cast superaustenitic stainless steel pitting chloride seawater"),
    ("q24", "solution annealed cast stainless steel pitting potential delta ferrite"),
    ("q30", "high nitrogen austenitic stainless steel pitting potential chloride solution treatment"),
]


def build(q, mode, page=1):
    if mode == "filter":
        return ("https://api.openalex.org/works?filter="
                + urllib.parse.quote("title_and_abstract.search:" + q, safe=":")
                + "&per-page=200&select=" + SELECT)
    return ("https://api.openalex.org/works?search=" + urllib.parse.quote(q)
            + "&per-page=200&select=" + SELECT)


def fetch(url):
    r = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(r, timeout=45) as resp:
        return json.loads(resp.read())


def run(code, q):
    for attempt in range(10):
        for mode in ("filter", "search"):
            try:
                d = fetch(build(q, mode))
                json.dump(d, open(os.path.join(RAW, code + ".json"), "w", encoding="utf-8"),
                          ensure_ascii=False)
                print("%s OK mode=%s count=%s got=%d (attempt %d)" %
                      (code, mode, d["meta"]["count"], len(d["results"]), attempt + 1), flush=True)
                return True
            except Exception as e:
                print("%s %s try%d %s %s" % (code, mode, attempt + 1,
                                             getattr(e, "code", ""), type(e).__name__), flush=True)
                time.sleep(4)
        time.sleep(42)
    print(code + " GIVE UP", flush=True)
    return False


if __name__ == "__main__":
    for code, q in QUERIES:
        run(code, q)
        time.sleep(15)
