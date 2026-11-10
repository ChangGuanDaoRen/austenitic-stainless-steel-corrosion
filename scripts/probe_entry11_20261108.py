# -*- coding: utf-8 -*-
r"""第 11 批只读探针：列 准\ 18 篇精确文件名、全文索引匹配、建议编号 160-177。"""
import json
import os
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent.parent
ZHUN = ROOT / "准"
LIT = ROOT / "文献"
TEXT_ALL = ROOT / "备份" / "text_all.jsonl"
MARK = "__有图待读.pdf"
FIRST_NO, LAST_NO = 160, 177


def pdf_ok(path: Path) -> bool:
    with path.open("rb") as f:
        return f.read(5) == b"%PDF-"


def normalized_name(name: str) -> str:
    if name.endswith(MARK):
        return name[:-len(MARK)] + ".pdf"
    return name


def load_index():
    by_exact = {}
    by_norm = {}
    with TEXT_ALL.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            if rec.get("tag") != "DAI":
                continue
            fn = str(rec.get("file") or "")
            if not fn:
                continue
            by_exact.setdefault(fn, rec)
            by_norm.setdefault(re.sub(r"__有图待读(?=\.pdf$)", "", fn), rec)
    return by_exact, by_norm


def main():
    files = sorted(x.name for x in ZHUN.iterdir() if x.is_file() and x.suffix.lower() == ".pdf")
    if len(files) != LAST_NO - FIRST_NO + 1:
        raise SystemExit(f"准\\ PDF 数应为 {LAST_NO - FIRST_NO + 1}，实测 {len(files)}: {files}")
    by_exact, by_norm = load_index()
    rows = []
    for i, fn in enumerate(files):
        no = FIRST_NO + i
        src = ZHUN / fn
        target = LIT / f"{no:02d}.pdf"
        base = normalized_name(fn)
        rec = by_exact.get(fn) or by_norm.get(base) or by_exact.get(base) or by_norm.get(fn)
        rows.append({
            "编号": no,
            "准文件名": fn,
            "全文索引名": (rec or {}).get("file", ""),
            "需读图": fn.endswith(MARK),
            "全文匹配": bool(rec),
            "字符数": (rec or {}).get("chars", 0),
            "页数": (rec or {}).get("pages", 0),
            "DOI": (rec or {}).get("doi", ""),
            "PDF头": pdf_ok(src),
            "目标": target.name,
            "目标已存在": target.exists(),
        })
    print(f"准PDF={len(files)} 需读图={sum(1 for r in rows if r['需读图'])} "
          f"非读图={sum(1 for r in rows if not r['需读图'])} "
          f"全文命中={sum(1 for r in rows if r['全文匹配'])} "
          f"目标已存在={sum(1 for r in rows if r['目标已存在'])}")
    print("编号 | 需读图 | 全文 | 字符 | 页 | 目标 | 文件名")
    for r in rows:
        print(f"{r['编号']} | {'Y' if r['需读图'] else 'N'} | "
              f"{'Y' if r['全文匹配'] else 'N'} | {r['字符数']} | {r['页数']} | "
              f"{r['目标']} | {r['准文件名']}")
        if r["全文匹配"] and r["全文索引名"] != r["准文件名"]:
            print(f"    索引名 -> {r['全文索引名']}")
    miss = [r for r in rows if not r["全文匹配"] or not r["PDF头"] or r["目标已存在"]]
    if miss:
        print("需处理：", [r["编号"] for r in miss])
        raise SystemExit(2)
    print("probe11: OK")


if __name__ == "__main__":
    main()
