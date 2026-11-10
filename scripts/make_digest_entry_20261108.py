# -*- coding: utf-8 -*-
"""为 83 篇生成紧凑 digest（head + 关键词命中段 + 尾部），供逐篇抽取 agent 先读；
text/NNN.txt 仍是完整（上限90k字符）原文，agent 不确定时再回读。
只读 备份\\text_all.jsonl 与 备份\\entry_20261108\\manifest.json；写 备份\\entry_20261108\\digest\\。
"""
import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
DIG = os.path.join(WORK, "digest")
os.makedirs(DIG, exist_ok=True)

KW = re.compile(
    r"3\.5\s*(wt|mass|%|g)|0\.6\s*(mol|M|kmol)|35\s*g/L|600\s*mM"
    r"|NaCl|氯化钠|氯离子|chloride|海水|seawater|人工海水|D1141|模拟海水"
    r"|点蚀电位|孔蚀电位|击穿电位|pitting|Epit|Eb\b|breakdown|polarization|极化曲线|阳极"
    r"|SCE|Ag/AgCl|calomel|甘汞|参比电极|reference electrode|饱和"
    r"|固溶|退火|时效|敏化|稳定化|均匀化|淬火|水淬|铸|cast|锻|热轧|冷轧|轧制|as[- ]received"
    r"|austenit|奥氏体|duplex|双相|铁素体|马氏体|δ|delta"
    r"|成分|composition|chemical|wt\.%|wt%|mass%|Table\s*\d|表\s*\d|Fig\.?\s*\d|图\s*\d"
    r"|Cr|Ni|Mo|Mn|Si|Cu|Nb|Ti|Co|Sn|Fe|balance|Bal\.|余量"
    r"|OCP|稳定时间|浸泡|immersion|scan rate|扫描速率|mV|V\b|SHE"
    r"|铸造|熔炼|感应炉|连铸|铸锭|砂型|离心|精密铸造"
)

by_file = {}
with open(os.path.join(ROOT, "备份", "text_all.jsonl"), encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if not line:
            continue
        try:
            rec = json.loads(line)
        except Exception:
            continue
        if rec.get("tag") == "DAI":
            by_file[rec.get("file")] = rec

with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
    manifest = json.load(f)

MARK = "================ 论文全文 ================\n"
sizes = []
for m in manifest:
    idx = m["idx"]
    fn = m["file"]
    full_path = os.path.join(WORK, "text", idx + ".txt")
    with open(full_path, encoding="utf-8") as f:
        file_text = f.read()
    hdr = file_text.split(MARK)[0]
    body = file_text.split(MARK, 1)[1] if MARK in file_text else file_text
    rec = by_file.get(fn)
    source = (rec or {}).get("text") or body
    head = source[:4500]
    tail = source[-2500:]
    mid, used, seen = [], 0, set()
    paras = re.split(r"\n\s*\n|\n", source)
    for p in paras:
        p = p.strip()
        if used >= 26000:
            break
        if len(p) < 8 or not KW.search(p):
            continue
        key = p[:100]
        if key in seen:
            continue
        seen.add(key)
        seg = p[:1600]
        mid.append(seg)
        used += len(seg)
    dg = (hdr + "【紧凑 digest：开头 4500 + 关键词命中段 ≤26000 + 结尾 2500 字符】\n"
          "【如需完整表格或数值，请再读 text/%s.txt（可用 read 的 offset/limit 分段）】\n\n"
          "===== 标题与摘要（开头） =====\n%s\n\n"
          "===== 关键证据段 =====\n%s\n\n===== 原文结尾 =====\n%s\n"
          % (idx, head, "\n---\n".join(mid), tail))
    with open(os.path.join(DIG, idx + ".txt"), "w", encoding="utf-8") as f:
        f.write(dg)
    sizes.append(len(dg))

print("digest files=", len(sizes), "min/median/max=", min(sizes), sorted(sizes)[len(sizes) // 2], max(sizes),
      "total MB=", round(sum(sizes) / 1024 / 1024, 2), flush=True)
