# -*- coding: utf-8 -*-
"""2026-11-08 只读初筛：从 text_all.jsonl 计算每篇证据 + 预判定，供人工复核。
只读，不改任何 PDF / xlsx。输出:
  temp\\rescreen_20261108\\triage_all.csv        262 行全字段
  temp\\rescreen_20261108\\triage_compact_1..4.txt  每行一篇的紧凑证据（分批复核用）
用法: D:\\Python\\python.exe -X utf8 scripts\\triage_20261108.py
"""
import csv
import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.join(ROOT, "temp", "rescreen_20261108")
SRC = os.path.join(BASE, "text_all.jsonl")
CSV_OUT = os.path.join(BASE, "triage_all.csv")
CMP = os.path.join(BASE, "triage_compact_{}.txt")

NEG = ("排除", "不收", "不包括", "不是", "非", "只", "仅")


def sq(s):
    return re.sub(r"\s+", " ", s or "").strip()


def snip(txt, m, before=70, after=110):
    if not m:
        return ""
    a = max(0, m.start() - before)
    return sq(txt[a:m.end() + after])


def first(txt, pats, flags=re.I):
    for p in pats:
        m = re.search(p, txt, flags)
        if m:
            return m, p
    return None, ""


# ---------------- 介质 ----------------
P_35 = [r"3\.5\s*(?:wt\.?\s*%|wt%|mass\s*%|weight\s*%|%)",
        r"3\.5\s*percent", r"35\s*g\s*[/·]?\s*L"]
P_06 = [r"0\.6\s*(?:mol\s*[/·]?\s*L|M\b|molar|kmol)", r"0\.61\s*mol",
        r"600\s*mM"]
P_SEA = [r"natural\s+seawater", r"real\s+seawater", r"seawater"]
P_ART = [r"artificial\s+seawater", r"synthetic\s+seawater", r"ASTM\s*D1141",
         r"simulated\s+seawater", r"\bSME\b"]
P_OTHER = [r"3\.56\s*%", r"3\.65\s*%", r"5\s*(?:wt\.?\s*%|wt%|mass\s*%)",
           r"5\s*%\s*NaCl", r"1\.4\s*%", r"double\s+concentration", r"wt\.?/v"]
P_ACID = [r"\bHCl\b", r"H2SO4|H₂SO₄", r"H2S|H₂S", r"sulfuric\s+acid",
          r"hydrochloric\s+acid", r"acidic\s+solution"]
P_FECL3 = [r"FeCl3|FeCl₃", r"ferric\s+chloride"]

# ---------------- 成形 ----------------
P_CAST = [r"cast\s+(?:austenitic\s+)?stainless", r"as-?cast", r"sand\s+cast",
          r"investment\s+cast", r"centrifugal\s+cast", r"casting", r"cast\b",
          r"铸", r"熔炼", r"熔模", r"砂型", r"离心铸造"]
P_WROUGHT = [r"hot[-\s]?rolled", r"cold[-\s]?rolled", r"forged", r"forging",
             r"rolled", r"轧", r"锻"]
P_ADD = [r"additive\s+manufactur", r"selective\s+laser\s+melt", r"\bSLM\b",
         r"\bLPBF\b", r"laser\s+powder\s+bed", r"\bDED\b", r"\bWAAM\b",
         r"wire\s+arc", r"3D\s+print", r"electron\s+beam\s+melt", r"\bEBM\b",
         r"directed\s+energy", r"增材", r"激光选区", r"电弧增材"]
P_POWDER = [r"powder\s+metallurg", r"\bSPS\b", r"spark\s+plasma\s+sinter",
            r"hot\s+isostatic", r"粉末冶金"]
P_WELD = [r"weld\s+(?:metal|overlay|joint|seam)", r"weld\s+overlay", r"cladding",
          r"welding", r"dissimilar\s+weld", r"堆焊", r"焊接"]
P_COAT = [r"coating", r"plasma\s+nitrid", r"nitriding", r"carburiz", r"渗氮",
          r"渗碳", r"镀层", r"涂层"]

# ---------------- 组织 ----------------
P_AUS = [r"austenitic", r"austenite", r"奥氏体", r"\bCF8M?\b", r"\bCF3M\b",
         r"\b316L?\b", r"\b304L?\b", r"\b321\b", r"\b347\b",
         r"super\s+austenitic", r"high[-\s]?nitrogen"]
P_DUP = [r"duplex\s+stainless", r"lean\s+duplex", r"双相"]
P_FER = [r"ferritic\s+stainless", r"martensitic", r"马氏体", r"铁素体不锈钢",
         r"17-4\s*PH"]
P_DELTA = [r"δ[-\s]?ferrite", r"delta\s+ferrite", r"δ\s*相",
           r"ferrite\s+content"]

# ---------------- 热处理 ----------------
P_SOL = [r"solution\s+(?:anneal|treat|heat)", r"固溶"]
P_ANN = [r"anneal", r"退火"]
P_SENS = [r"sensitiz", r"敏化"]
P_AGE = [r"aging|ageing", r"时效"]
P_STAB = [r"stabiliz", r"稳定化"]
P_AGED_HOT = [r"thermal\s+aging", r"热老化"]

# ---------------- 电化学 ----------------
P_EPIT = [r"pitting\s+potential", r"breakdown\s+potential", r"\bEpit\b",
          r"E\s*pit", r"点蚀电位", r"点蚀点位"]
P_REF = [r"\bSCE\b", r"saturated\s+calomel", r"Ag\s*/\s*AgCl",
         r"Hg\s*/\s*Hg2Cl2", r"reference\s+electrode", r"参比电极"]
P_TAFEL = [r"Tafel\s+extrapol", r"extrapolat", r"外推"]
P_CPT = [r"critical\s+pitting\s+temperature", r"\bCPT\b", r"临界点蚀温度"]


def judge(txt):
    ev = {}
    m35, _ = first(txt, P_35)
    m06, _ = first(txt, P_06)
    msea, _ = first(txt, P_SEA)
    mart, _ = first(txt, P_ART)
    moth, _ = first(txt, P_OTHER)
    macid, _ = first(txt, P_ACID)
    mfe, _ = first(txt, P_FECL3)
    ev["medium"] = " | ".join(x for x in (
        ("3.5%:" + snip(txt, m35, 30, 70)) if m35 else "",
        ("0.6M:" + snip(txt, m06, 30, 70)) if m06 else "",
        ("海水:" + snip(txt, msea, 40, 70)) if msea else "",
        ("人工海水:" + snip(txt, mart, 40, 70)) if mart else "",
        ("其它浓度:" + snip(txt, moth, 40, 80)) if moth else "",
        ("酸:" + snip(txt, macid, 30, 60)) if macid else "",
        ("FeCl3:" + snip(txt, mfe, 30, 60)) if mfe else "",
    ) if x)

    madd, _ = first(txt, P_ADD)
    mpow, _ = first(txt, P_POWDER)
    mweld, _ = first(txt, P_WELD)
    mcoat, _ = first(txt, P_COAT)
    mcast, _ = first(txt, P_CAST)
    mwro, _ = first(txt, P_WROUGHT)
    ev["form"] = " | ".join(x for x in (
        ("增材:" + snip(txt, madd, 30, 70)) if madd else "",
        ("粉末:" + snip(txt, mpow, 30, 70)) if mpow else "",
        ("焊接:" + snip(txt, mweld, 30, 70)) if mweld else "",
        ("涂层:" + snip(txt, mcoat, 30, 70)) if mcoat else "",
        ("铸:" + snip(txt, mcast, 30, 60)) if mcast else "",
        ("锻轧:" + snip(txt, mwro, 30, 60)) if mwro else "",
    ) if x)

    mdup, _ = first(txt, P_DUP)
    mfer, _ = first(txt, P_FER)
    maus, _ = first(txt, P_AUS)
    mdel, _ = first(txt, P_DELTA)
    ev["micro"] = " | ".join(x for x in (
        ("双相:" + snip(txt, mdup, 30, 70)) if mdup else "",
        ("铁素体/马氏体:" + snip(txt, mfer, 30, 70)) if mfer else "",
        ("δ:" + snip(txt, mdel, 30, 60)) if mdel else "",
        ("奥氏体:" + snip(txt, maus, 30, 60)) if maus else "",
    ) if x)

    msol, _ = first(txt, P_SOL)
    mann, _ = first(txt, P_ANN)
    msens, _ = first(txt, P_SENS)
    mage, _ = first(txt, P_AGE)
    mstab, _ = first(txt, P_STAB)
    mhot, _ = first(txt, P_AGED_HOT)
    ev["heat"] = " | ".join(x for x in (
        ("固溶:" + snip(txt, msol, 40, 80)) if msol else "",
        ("退火:" + snip(txt, mann, 40, 80)) if mann else "",
        ("敏化:" + snip(txt, msens, 40, 70)) if msens else "",
        ("时效:" + snip(txt, mage, 40, 70)) if mage else "",
        ("稳定化:" + snip(txt, mstab, 40, 70)) if mstab else "",
        ("热老化:" + snip(txt, mhot, 40, 60)) if mhot else "",
    ) if x)

    mep, _ = first(txt, P_EPIT)
    mref, _ = first(txt, P_REF)
    mtaf, _ = first(txt, P_TAFEL)
    mcpt, _ = first(txt, P_CPT)
    ev["ec"] = " | ".join(x for x in (
        ("Epit:" + snip(txt, mep, 40, 70)) if mep else "",
        ("参比:" + snip(txt, mref, 30, 60)) if mref else "",
        ("Tafel:" + snip(txt, mtaf, 40, 60)) if mtaf else "",
        ("CPT:" + snip(txt, mcpt, 30, 60)) if mcpt else "",
    ) if x)

    if madd:
        verdict, trig = "不符", "N1-增材"
    elif mpow:
        verdict, trig = "不符", "N1-粉末冶金"
    elif mcoat:
        verdict, trig = "不符", "N1-表面改性/涂层"
    elif mfe and not m35 and not m06:
        verdict, trig = "不符", "N3-FeCl3"
    elif mdup:
        verdict, trig = "不符", "N2-双相"
    elif mfer:
        verdict, trig = "存疑-待补证", "N2-铁素体/马氏体字样"
    elif not (m35 or m06 or msea or mart):
        if macid:
            verdict, trig = "不符", "N3-酸性介质"
        else:
            verdict, trig = "不符", "N3-无3.5%NaCl介质"
    elif mart:
        verdict, trig = "存疑-待补证", "N3-人工/模拟海水"
    elif moth and not (m35 or m06):
        verdict, trig = "存疑-待补证", "N3-非标准浓度"
    else:
        verdict, trig = "存疑-待补证", "待人工复核"
    if mweld and verdict != "不符":
        trig += "+焊接字样"
    return verdict, trig, ev


def main():
    recs = []
    with open(SRC, encoding="utf-8") as f:
        for line in f:
            recs.append(json.loads(line))
    print("loaded", len(recs), flush=True)

    lib = [r for r in recs if r["tag"] == "WEN"]
    lib_doi = {}
    for r in lib:
        if r["doi"]:
            lib_doi[r["doi"]] = r["file"]

    rows = []
    for r in recs:
        if r["tag"] != "DAI":
            continue
        v, t, ev = judge(r["text"])
        if r["doi"] and r["doi"] in lib_doi:
            v, t = "重复-已入库", "重复:DOI=" + lib_doi[r["doi"]]
        title = sq(r["head"].split("\n")[0])[:120]
        rows.append({
            "文件": r["file"],
            "页数": r["pages"],
            "字数": r["chars"],
            "标题": title,
            "DOI": r["doi"],
            "预判定": v,
            "触发条款": t,
            "介质证据": ev["medium"][:400],
            "成形证据": ev["form"][:300],
            "组织证据": ev["micro"][:240],
            "热处理证据": ev["heat"][:300],
            "电化学证据": ev["ec"][:300],
        })
    cols = list(rows[0].keys())
    with open(CSV_OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    print("wrote", CSV_OUT, len(rows), flush=True)

    n = len(rows)
    per = (n + 3) // 4
    for i in range(4):
        chunk = rows[i * per:(i + 1) * per]
        if not chunk:
            continue
        path = CMP.format(i + 1)
        with open(path, "w", encoding="utf-8") as f:
            for k, row in enumerate(chunk, start=i * per + 1):
                f.write("[%d] %s (%s页)\n  标题:%s\n  预判:%s | %s\n  介质:%s\n  成形:%s\n  组织:%s\n  热处理:%s\n  电化学:%s\n\n" % (
                    k, row["文件"], row["页数"], row["标题"], row["预判定"],
                    row["触发条款"], row["介质证据"], row["成形证据"],
                    row["组织证据"], row["热处理证据"], row["电化学证据"]))
        print("wrote", path, len(chunk), flush=True)

    from collections import Counter
    print("dist", dict(Counter(r["预判定"] for r in rows)), flush=True)


if __name__ == "__main__":
    main()
