# -*- coding: utf-8 -*-
"""只读：核对 准\ 83 篇的入库底稿是否齐备（全文、判级、编号、重复）。"""
import os, json, csv, collections

ROOT = os.getcwd()
QUASI = os.path.join(ROOT, "准")
BAK = os.path.join(ROOT, "备份")
TXT = os.path.join(BAK, "text_all.jsonl")
LEDGER = os.path.join(ROOT, "逐篇判级_移动台账_20261108.csv")

qfiles = sorted(f for f in os.listdir(QUASI) if f.lower().endswith(".pdf"))
print("准 PDF:", len(qfiles))

# 1) 全文底稿
texts = {}
with open(TXT, "r", encoding="utf-8") as fh:
    for line in fh:
        try:
            d = json.loads(line)
        except Exception:
            continue
        if d.get("tag") == "DAI":
            texts[d["file"]] = (d.get("chars", 0), d.get("pages", 0), d.get("doi", ""))
print("text_all DAI 条数:", len(texts))
hit = [f for f in qfiles if f in texts]
miss = [f for f in qfiles if f not in texts]
print("准 83 篇里有全文:", len(hit), "无全文:", len(miss))
if miss:
    print("无全文名单:", miss)
short = [(f, texts[f][0]) for f in hit if texts[f][0] < 2000]
print("全文<2000字符(疑扫描件):", len(short), short)
lens = sorted(texts[f][0] for f in hit)
print("全文字符 min/中位/max:", lens[0], lens[len(lens)//2], lens[-1])

# 2) 判级台账
verd = {}
with open(LEDGER, "r", encoding="utf-8-sig", newline="") as fh:
    for r in csv.DictReader(fh):
        verd[r["文件"]] = r["判定"]
print("\n台账条数:", len(verd))
c = collections.Counter(verd.get(f, "台账缺") for f in qfiles)
print("准 83 篇判定分布:", dict(c))
base = [f for f in qfiles if verd.get(f) == "合规-基准组"]
print("合规-基准组篇数:", len(base))

# 3) 编号与文献目录
import openpyxl
wb = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"), data_only=True)
ws = wb["Sheet1"]
ids = [int(ws.cell(r, 1).value) for r in range(3, ws.max_row + 1)
       if ws.cell(r, 1).value not in (None, "")]
print("\n0.xlsx 行/列/块:", ws.max_row, ws.max_column, len(ids),
      "编号范围", min(ids), max(ids), "缺号",
      [i for i in range(1, max(ids) + 1) if i not in ids])
lit = sorted(f for f in os.listdir(os.path.join(ROOT, "文献")) if f.lower().endswith(".pdf"))
print("文献 PDF:", len(lit), "首末:", lit[0], lit[-1])
print("新编号建议: 64 –", 63 + len(qfiles))

# 4) 准 与 文献 是否同名重复
print("准/文献 同名交集:", len(set(qfiles) & set(lit)))
