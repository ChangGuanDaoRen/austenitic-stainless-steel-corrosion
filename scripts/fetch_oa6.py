# -*- coding: utf-8 -*-
"""第 6 轮候选全文抓取（只走合规开放获取通道）

候选来自 temp/lit/oa6/ranked.csv。对每个 DOI：
  1) OpenAlex 给的 oa_url（best_oa_location 等价物）
  2) Unpaywall best_oa_location.url_for_pdf / url
  3) CORE / PMC / SciELO / MDPI / doaj 等常见 OA 直链（由 oa_url 派生）
不碰付费墙，不做登录态复用。

产物：temp/lit/oa6/pdf/<slug>.pdf
      temp/lit/oa6/fetch_ledger.csv
用法：python fetch_oa6.py [--limit N] [--min-score S]
"""
import csv
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6")
PDF = os.path.join(OUT, "pdf")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"


def http(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": UA,
                                               "Accept": "application/pdf,*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read(), r.geturl(), r.headers.get("Content-Type", "")


def unpaywall(doi):
    u = "https://api.unpaywall.org/v2/%s?email=3144668926@qq.com" % urllib.parse.quote(doi, safe="")
    try:
        b, _, _ = http(u, 20)
        d = json.loads(b)
        loc = d.get("best_oa_location") or {}
        return [x for x in (loc.get("url_for_pdf"), loc.get("url")) if x]
    except Exception:
        return []


def is_pdf(b):
    return len(b) > 12000 and b[:5] == b"%PDF-"


def pdf_links_from_html(b, base):
    """落地页 → PDF 直链（citation_pdf_url / 常见全文按钮）。"""
    try:
        h = b.decode("utf-8", "ignore")
    except Exception:
        return []
    out = []
    for pat in (r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)',
                r'<meta[^>]+content=["\']([^"\']+\.pdf[^"\']*)["\'][^>]+name=["\']citation_pdf_url',
                r'href=["\']([^"\']*(?:/pdf|\.pdf|pdfdirect|download)[^"\']*)["\']'):
        for m in re.finditer(pat, h, re.I):
            u = m.group(1).replace("&amp;", "&")
            out.append(urllib.parse.urljoin(base, u))
    return list(dict.fromkeys(out))[:6]


def main():
    limit = 250
    minscore = 6
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    if "--min-score" in sys.argv:
        minscore = int(sys.argv[sys.argv.index("--min-score") + 1])

    os.makedirs(PDF, exist_ok=True)
    rows = list(csv.DictReader(open(os.path.join(OUT, "ranked.csv"), encoding="utf-8-sig")))
    todo = [r for r in rows if not r["dup"] and r["doi"] and int(r["score"]) >= minscore]
    print("候选 %d 条（score>=%d，已去重，有 DOI）" % (len(todo), minscore))

    ledger_path = os.path.join(OUT, "fetch_ledger.csv")
    done = {}
    if os.path.exists(ledger_path):
        done = {r["doi"]: r for r in csv.DictReader(open(ledger_path, encoding="utf-8-sig"))}

    n_ok = 0
    for i, r in enumerate(todo[:limit], 1):
        doi = r["doi"]
        if doi in done and done[doi]["status"] == "OK":
            continue
        urls = []
        if r["oa_url"]:
            urls.append(r["oa_url"])
        urls += unpaywall(doi)
        urls.append("https://doi.org/" + doi)

        status, src, size = "NOPDF", "", 0
        tried = set()
        queue = list(dict.fromkeys(urls))
        while queue and status != "OK":
            u = queue.pop(0)
            if u in tried:
                continue
            tried.add(u)
            try:
                b, final, ctype = http(u, 25)
                if is_pdf(b):
                    fn = re.sub(r"[^A-Za-z0-9._-]", "_", doi)[:110] + ".pdf"
                    with open(os.path.join(PDF, fn), "wb") as f:
                        f.write(b)
                    status, src, size = "OK", final[:180], len(b)
                    n_ok += 1
                    break
                if b[:1] == b"<":
                    queue.extend([x for x in pdf_links_from_html(b, final) if x not in tried])
                elif "pdf" in ctype.lower():
                    status = "SMALL"
            except Exception as e:
                status = "ERR:" + type(e).__name__
            time.sleep(0.3)

        done[doi] = {"doi": doi, "title": r["title"][:120], "score": r["score"],
                     "status": status, "src": src, "size": size}
        print("[%3d/%d] %-9s %7dB %s" % (i, min(len(todo), limit), status, size,
                                          r["title"][:66]), flush=True)
        if i % 10 == 0:
            with open(ledger_path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=["doi", "title", "score", "status", "src", "size"])
                w.writeheader()
                w.writerows(done.values())

    with open(ledger_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["doi", "title", "score", "status", "src", "size"])
        w.writeheader()
        w.writerows(done.values())
    print("完成：新下 %d 篇 → %s" % (n_ok, PDF))


if __name__ == "__main__":
    main()
