# -*- coding: utf-8 -*-
"""第 6 轮检索 · 引文网络 OA 全文批量抓取（带站点专用直链规则）"""
import os, re, json, time, csv, urllib.request, urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6", "refs", "pdf")
os.makedirs(OUT, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"

# 站点专用：直接给 CDN 直链，绕开 www.mdpi.com 403 与落地页跳转
EXTRA = {
    "10.3390/met8080653": ["https://mdpi-res.com/d_attachment/metals/metals-08-00653/article_deploy/metals-08-00653.pdf"],
    "10.3390/ma16237397": ["https://mdpi-res.com/d_attachment/materials/materials-16-07397/article_deploy/materials-16-07397.pdf",
                           "https://mdpi-res.com/d_attachment/materials/materials-16-23797/article_deploy/materials-16-23797.pdf"],
    "10.3390/met13071244": ["https://mdpi-res.com/d_attachment/metals/metals-13-01244/article_deploy/metals-13-01244.pdf"],
    "10.1590/0370-44672024780044": [],  # 已有
    "10.1038/s41598-018-23273-x": ["https://www.nature.com/articles/s41598-018-23273-x.pdf"],
    "10.20964/2022.04.29": ["https://www.ijesi.org/papers/Vol17-4/17-4-29.pdf"],
    "10.20964/2021.11.05": ["https://www.ijesi.org/papers/Vol16-11/16-11-05.pdf"],
    "10.1016/j.corsci.2016.01.022": ["https://www.sciencedirect.com/science/article/am/pii/S0010938X16300221/pdfft"],
    "10.1016/j.jmrt.2026.05.006": ["https://www.sciencedirect.com/science/article/pii/S223878542601328X/pdfft"],
}


def http(u, t=28):
    r = urllib.request.Request(u, headers={
        "User-Agent": UA, "Accept": "application/pdf,text/html,*/*",
        "Referer": urllib.parse.urlsplit(u).scheme + "://" + urllib.parse.urlsplit(u).netloc + "/"})
    with urllib.request.urlopen(r, timeout=t) as resp:
        return resp.read(), resp.geturl(), resp.headers.get("Content-Type", "")


def pdf_name(doi):
    return re.sub(r"[^A-Za-z0-9._-]", "_", doi)[:100] + ".pdf"


def main():
    scored = json.load(open(os.path.join(BASE, "temp", "lit", "oa6", "refs", "scored.json"), encoding="utf-8"))
    # 只取 OA 且未见过、分数 >= 8 的；再加指定 DOI
    picks = [x for x in scored if x["oa"] and not x["dup"] and x["score"] >= 8][:45]
    bydoi = {x["doi"]: x for x in picks}
    for d in EXTRA:
        hit = next((x for x in scored if x["doi"] == d), None)
        if hit:
            bydoi[d] = hit
    ledger = []
    for doi, x in bydoi.items():
        dest = os.path.join(OUT, pdf_name(doi))
        if os.path.exists(dest) and os.path.getsize(dest) > 20000:
            ledger.append([doi, x["title"][:90], "SKIP", "", os.path.getsize(dest)])
            continue
        urls = list(EXTRA.get(doi, []))
        if x.get("oa_url"):
            urls.append(x["oa_url"])
        # MDPI 落地页 → CDN
        for u in list(urls):
            m = re.search(r"mdpi\.com/(\d{4}-\d{4})/(\d+)/(\d+)/(\d+)/pdf", u)
            if m:
                jmap = {"2075-4701": "metals", "1996-1944": "materials", "2079-6412": "coatings"}
                j = jmap.get(m.group(1))
                if j:
                    tag = "%s-%02d-%05d" % (j, int(m.group(2)), int(m.group(4)))
                    urls.insert(0, "https://mdpi-res.com/d_attachment/%s/%s/article_deploy/%s.pdf" % (j, tag, tag))
        status, src, size = "FAIL", "", 0
        tried = set()
        for u in urls:
            if u in tried or u.startswith("https://doi.org"):
                continue
            tried.add(u)
            try:
                b, final, ct = http(u)
                if len(b) > 20000 and b[:5] == b"%PDF-":
                    open(dest, "wb").write(b)
                    status, src, size = "OK", final[:170], len(b)
                    break
            except Exception as e:
                status = "ERR:" + type(e).__name__
            time.sleep(0.25)
        ledger.append([doi, x["title"][:90], status, src, size])
        print("%-5s %-42s %s" % (status, doi[:42], x["title"][:62]))
    p = os.path.join(BASE, "temp", "lit", "oa6", "refs", "fetch_ledger.csv")
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(["doi", "title", "status", "src", "size"]); w.writerows(ledger)
    print("OK %d / %d" % (sum(1 for r in ledger if r[2] == "OK"), len(ledger)))


if __name__ == "__main__":
    main()
