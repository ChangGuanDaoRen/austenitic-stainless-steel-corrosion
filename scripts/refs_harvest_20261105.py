# -*- coding: utf-8 -*-
"""第 6 轮检索 · 通道 7：引文网络挖掘

从 `文献\\01-55.pdf` 与本轮新下全文的「参考文献」段抽 DOI，
与已见索引比对，产出全新题录候选；再走 OpenAlex 批量补全元数据。

用法:
    python scripts/refs_harvest_20261105.py harvest   # 抽 DOI + 去重
    python scripts/refs_harvest_20261105.py meta      # OpenAlex 批量补元数据
"""
import os, re, sys, json, glob, time, urllib.request, urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "temp", "lit", "oa6", "refs")
os.makedirs(OUT, exist_ok=True)
UA = "cast-ASS/1.0 (mailto:3144668926@qq.com)"

DOI_START = re.compile(r"10\.\d{4,9}/", re.I)
# DOI 主体允许 字母数字 . - _ ( ) / + : ；遇到空白、中文、逗号分号、尖括号、引号即止
DOI_BODY = re.compile(r"[A-Za-z0-9._()/;:+-]+")


def clean_doi(d):
    d = d.strip().rstrip(".,;:)]}/-")
    d = re.sub(r"(https?://(dx\.)?doi\.org/)", "", d, flags=re.I)
    d = re.sub(r"(Abstract|FullText|CrossRef|Google|View|Article)$", "", d, flags=re.I)
    # 括号不成对 → 截断
    if d.count("(") != d.count(")"):
        i = min([x for x in (d.find("("), d.find(")")) if x >= 0] or [len(d)])
        d = d[:i].rstrip(".-/")
    return d.lower()


def pdf_dois(path):
    try:
        import fitz
        d = fitz.open(path)
    except Exception:
        return []
    txt = []
    for i in range(d.page_count):
        try:
            txt.append(d[i].get_text())
        except Exception:
            pass
    s = "\n".join(txt)
    # 把 DOI 常见的断行接回来：「j\n.corsci」「-\n2011」「(\n94」等
    s = re.sub(r"-\s*\n\s*", "-", s)
    # 只在「明显是 DOI 中间断开」时接行：下一字符为 . - ( ) 数字，或 DOI 尾停在 /j /s 等短前缀
    def _join(m):
        tail, nxt = m.group(1), m.group(2)
        if nxt in ".-()0123456789" or re.search(r"/[js]$|/[a-z]{1,3}$", tail[-8:]):
            return tail + nxt
        return m.group(0)
    for _ in range(3):
        s2 = re.sub(r"(10\.\d{4,9}/[A-Za-z0-9._()/;:+-]*)\s*\n\s*([A-Za-z0-9._()/;:+-])",
                    _join, s)
        if s2 == s:
            break
        s = s2
    out = []
    for m in DOI_START.finditer(s):
        tail = s[m.end():m.end() + 130]
        bm = DOI_BODY.match(tail)
        if not bm:
            continue
        d0 = clean_doi(m.group(0) + bm.group(0))
        if len(d0) > 10:
            out.append(d0)
    return out


def load_seen():
    p = os.path.join(BASE, "temp", "lit", "oa6", "seen.json")
    seen = json.load(open(p, encoding="utf-8"))
    dois, titles = set(), set()
    for x in seen:
        if isinstance(x, dict):
            if x.get("doi"):
                dois.add(clean_doi(str(x["doi"])))
            if x.get("title"):
                titles.add(re.sub(r"[^a-z0-9]+", " ", str(x["title"]).lower()).strip())
    # 0.xlsx 题录（Y 列）里的 DOI / 片段
    try:
        import openpyxl
        ws = openpyxl.load_workbook(os.path.join(BASE, "0.xlsx"), data_only=True)["Sheet1"]
        for r in range(3, ws.max_row + 1):
            v = ws.cell(row=r, column=25).value
            if v:
                for m in DOI_RE.finditer(str(v)):
                    dois.add(clean_doi(m.group(0)))
    except Exception:
        pass
    return dois, titles


def harvest():
    files = sorted(glob.glob(os.path.join(BASE, "文献", "*.pdf")))
    files += sorted(glob.glob(os.path.join(BASE, "temp", "lit", "oa6", "web", "*.pdf")))
    files += sorted(glob.glob(os.path.join(BASE, "temp", "lit", "oa6", "pdf", "*.pdf")))
    seen_doi, seen_title = load_seen()
    allmap = {}
    for f in files:
        ds = pdf_dois(f)
        for doi in ds:
            allmap.setdefault(doi, []).append(os.path.basename(f))
    print("PDF 数 %d，抽到唯一 DOI %d 条" % (len(files), len(allmap)))
    new = {d: v for d, v in allmap.items() if d not in seen_doi}
    print("已见 DOI %d；全新 DOI %d 条" % (len(seen_doi), len(new)))
    json.dump(allmap, open(os.path.join(OUT, "all_dois.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    json.dump(new, open(os.path.join(OUT, "new_dois.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    # 高频 DOI = 领域基石文献，优先补元数据
    top = sorted(new.items(), key=lambda x: -len(x[1]))[:30]
    print("\n--- 被最多库内文献引用的新 DOI（前 30）")
    for d, srcs in top:
        print("  %2d  %-62s  %s" % (len(srcs), d[:62], ",".join(srcs[:3])))


def fetch(url, tries=5):
    for i in range(tries):
        try:
            r = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(r, timeout=40) as resp:
                return json.loads(resp.read())
        except Exception as e:
            code = getattr(e, "code", "")
            if i == tries - 1:
                print("    FAIL", code, type(e).__name__)
                return None
            time.sleep((8 if code == 429 else 3) * (i + 1))
    return None


SELECT = "id,doi,title,publication_year,type,cited_by_count,primary_location,open_access,best_oa_location,authorships,abstract_inverted_index"


def meta():
    new = json.load(open(os.path.join(OUT, "new_dois.json"), encoding="utf-8"))
    dois = list(new)
    out = []
    for i in range(0, len(dois), 40):
        chunk = dois[i:i + 40]
        filt = "doi:" + "|".join(urllib.parse.quote("https://doi.org/" + d, safe="") for d in chunk)
        url = ("https://api.openalex.org/works?filter=" + filt +
               "&per-page=50&select=" + SELECT)
        d = fetch(url)
        if not d:
            continue
        for w in d.get("results", []):
            out.append({
                "doi": (w.get("doi") or "").replace("https://doi.org/", "").lower(),
                "title": w.get("title") or "",
                "year": w.get("publication_year"),
                "cited": w.get("cited_by_count"),
                "type": w.get("type"),
                "oa": bool((w.get("open_access") or {}).get("is_oa")),
                "oa_url": ((w.get("best_oa_location") or {}).get("pdf_url")
                           or (w.get("best_oa_location") or {}).get("landing_page_url")
                           or ((w.get("primary_location") or {}).get("landing_page_url")) or ""),
                "journal": (((w.get("primary_location") or {}).get("source") or {}).get("display_name")) or "",
                "src_refs": new.get((w.get("doi") or "").replace("https://doi.org/", "").lower(), []),
            })
        print("  %d/%d 已补 %d 条" % (min(i + 40, len(dois)), len(dois), len(out)))
        time.sleep(2)
    json.dump(out, open(os.path.join(OUT, "meta.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("写入 meta.json：%d 条" % len(out))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "harvest"
    {"harvest": harvest, "meta": meta}[cmd]()
