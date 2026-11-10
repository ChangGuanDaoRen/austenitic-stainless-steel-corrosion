# -*- coding: utf-8 -*-
"""为「逐篇精读」准备材料（只读 PDF 缓存，不改任何既有文件）。

输入：temp\\rescreen_20261108\\text_all.jsonl（tag=DAI 的 262 篇全文）
输出：
  temp\\judge_20261108\\digest\\pNNN.txt   每篇一份证据摘要（标题/摘要 + 关键段 + 结尾）
  temp\\judge_20261108\\manifest.json      索引 -> 文件名/现位置/字符数/DOI
  temp\\judge_20261108\\verdicts\\          （空目录，子智能体写判定）
用法：D:\\Python\\python.exe -X utf8 scripts\\prep_judge_20261108.py
"""
import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "temp", "rescreen_20261108", "text_all.jsonl")
OUT = os.path.join(ROOT, "temp", "judge_20261108")
DIG = os.path.join(OUT, "digest")
VER = os.path.join(OUT, "verdicts")
for d in (OUT, DIG, VER):
    os.makedirs(d, exist_ok=True)

HEAD_N = 3000
TAIL_N = 1500
MID_MAX = 13000

KW = re.compile(
    r"NaCl|氯化钠|氯离子|chloride|3\.5\s*(?:wt|mass|%)|0\.6\s*(?:mol|M\b|kmol)|35\s*g/L|600\s*mM"
    r"|海水|seawater|人工海水|D1141|模拟"
    r"|pitting|点蚀|孔蚀|点蚀电位|Epit|E_pit|Eb\b|breakdown"
    r"|SCE|Ag/AgCl|calomel|甘汞|参比|reference electrode"
    r"|固溶|solution treat|退火|anneal|时效|aging|敏化|sensitiz|稳定化|stabiliz|热老化"
    r"|铸|cast|锻|forge|轧|roll|热轧|冷轧|增材|additive|SLM|LPBF|激光|焊接|weld|粉末|powder"
    r"|duplex|双相|铁素体|ferrite|马氏体|martensite|奥氏体|austenit"
    r"|composition|成分|wt\.\s*%|Cr\b|Ni\b|Mo\b|Mn\b|Si\b|Cu\b|N\b"
)


def split_paras(text):
    parts = re.split(r"\n\s*\n|\r\n\s*\r\n", text)
    if len(parts) < 5:
        parts = text.split("\n")
    return [p.strip() for p in parts if p.strip()]


def build_digest(rec):
    text = rec.get("text") or ""
    head = (rec.get("head") or "")[:HEAD_N]
    if not head:
        head = text[:HEAD_N]
    mid = []
    used = 0
    seen = set()
    for p in split_paras(text):
        if used >= MID_MAX:
            break
        if not KW.search(p):
            continue
        key = p[:80]
        if key in seen:
            continue
        seen.add(key)
        take = p[:900]
        mid.append(take)
        used += len(take)
    tail = text[-TAIL_N:] if len(text) > TAIL_N else ""
    return (
        "【文件】%s\n【目录】%s\n【页数】%s\n【字符数】%s\n【DOI】%s\n\n"
        "===== 标题与摘要（原文开头） =====\n%s\n\n"
        "===== 关键证据段（含介质/成分/电位/参比/工艺/组织关键词） =====\n%s\n\n"
        "===== 原文结尾 =====\n%s\n"
        % (
            rec["file"],
            rec["dir"],
            rec.get("pages"),
            rec.get("chars"),
            rec.get("doi") or "",
            head,
            "\n---\n".join(mid),
            tail,
        )
    )


def main():
    rows = []
    with open(SRC, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            if d.get("tag") != "DAI":
                continue
            rows.append(d)
    rows.sort(key=lambda r: r["file"])
    manifest = []
    for i, rec in enumerate(rows, 1):
        idx = "p%03d" % i
        dg = build_digest(rec)
        with open(os.path.join(DIG, idx + ".txt"), "w", encoding="utf-8") as f:
            f.write(dg)
        manifest.append(
            {
                "idx": idx,
                "file": rec["file"],
                "dir_at_extract": rec["dir"],
                "chars": rec.get("chars"),
                "pages": rec.get("pages"),
                "doi": rec.get("doi") or "",
                "digest_chars": len(dg),
            }
        )
    with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    sizes = [m["digest_chars"] for m in manifest]
    print("papers=%d digest_total=%.2f MB avg=%d max=%d min=%d"
          % (len(manifest), sum(sizes) / 1e6, sum(sizes) // len(sizes), max(sizes), min(sizes)))
    zero = [m["file"] for m in manifest if (m["chars"] or 0) < 2000]
    print("低字符数（抽取失败，需 OCR）: %d" % len(zero))
    for z in zero:
        print("  ", z)


if __name__ == "__main__":
    main()
