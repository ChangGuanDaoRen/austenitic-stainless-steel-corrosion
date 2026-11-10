r"""第11批卡片只读审计：检查 18 张卡片结构、编号、文件名、语义与读图标注。"""
import csv
import json
import os
import sys

ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
WORK = os.path.join(ROOT, "备份", "entry11_20261108")
CARDS = os.path.join(WORK, "cards")
MANIFEST = os.path.join(WORK, "manifest.json")
OUT = os.path.join(WORK, "cards_审计_第11批_20261108.csv")
FIRST, LAST = 160, 177
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]
BAD_WORDS = ["→SCE", "0.1989", "换算至 SCE", "清洗时", "1.xlsx", "2.xlsx",
             "【本轮入库", "【提取方式", "口径推演", "介质复述"]


def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


manifest = load_json(MANIFEST, [])
by_no = {int(x["newno"]): x for x in manifest}
rows = []
missing = []
block = []
remind = []
na = []
all_missing_elements = []

for no in range(FIRST, LAST + 1):
    path = os.path.join(CARDS, f"{no:03d}.json")
    if not os.path.exists(path):
        missing.append(no)
        continue
    try:
        c = load_json(path, {})
    except Exception as e:
        block.append(f"{no} JSON错误 {e}")
        continue
    mf = by_no.get(no, {})
    problems = []
    if c.get("newno") not in (no, str(no)):
        problems.append(f"newno={c.get('newno')!r}")
    if str(c.get("file") or "").strip() != str(mf.get("file") or "").strip():
        problems.append(f"file与台账不符：{c.get('file')!r}")
    if not str(c.get("citation") or "").strip():
        problems.append("citation空")
    if not str(c.get("notes") or "").strip():
        problems.append("notes空")
    if "flag" in c:
        problems.append("含禁字段flag")
    rr = c.get("rows")
    if not isinstance(rr, list) or not rr:
        problems.append("rows为空")
    notes = str(c.get("notes") or "")
    if len(notes) > 120:
        remind.append(f"{no} 备注{len(notes)}字(>120)")
    elif len(notes) > 80:
        remind.append(f"{no} 备注{len(notes)}字(>80，建议压缩)")
    blob = "\n".join([str(c.get("citation") or ""), notes,
                      *[str(v) for r in rr for v in r.values()]])
    hits = [w for w in BAD_WORDS if w in blob]
    if hits:
        problems.append("不该进0.xlsx：" + ",".join(hits))
    for i, r in enumerate(rr):
        if not isinstance(r, dict):
            problems.append(f"row{i}非对象")
            continue
        miss = [k for k in FIELDS if k not in r or r[k] is None or str(r[k]).strip() == ""]
        extra = [k for k in r if k not in FIELDS]
        if miss:
            problems.append(f"row{i}缺{','.join(miss)}")
        if extra:
            remind.append(f"{no} row{i}多字段{','.join(extra)}（已忽略）")
        for k, v in r.items():
            if k in FIELDS and not isinstance(v, str):
                problems.append(f"row{i}/{k}非字符串：{type(v).__name__}")
        if str(r.get("Epit") or "").strip() == "N/A（待读图）":
            na.append((no, i, str(r.get("sample") or "")))
        if all(str(r.get(k) or "").strip() == "未提及" for k in FIELDS[1:16]):
            all_missing_elements.append((no, i, str(r.get("sample") or "")))
    needs = bool(mf.get("needs_fig"))
    if needs and not any(str(r.get("Epit") or "").strip() == "N/A（待读图）" for r in rr):
        remind.append(f"{no} 需读图但Epit未标记N/A（待读图）；若正文有原文数值则合理")
    if not needs and any(str(r.get("Epit") or "").strip() == "N/A（待读图）" for r in rr):
        problems.append(f"{no} 非读图件却记N/A（待读图）")
    rows.append({"新编号": no, "原文件名": mf.get("file", ""), "需读图": "是" if needs else "否",
                 "数据行数": len(rr), "citation字数": len(str(c.get("citation") or "")),
                 "notes字数": len(notes), "状态": "ok" if not problems else "阻断",
                 "问题": "；".join(problems)})
    if problems:
        block.append(f"{no} " + "；".join(problems))

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["新编号", "原文件名", "需读图", "数据行数",
                                       "citation字数", "notes字数", "状态", "问题"])
    w.writeheader()
    w.writerows(rows)

print(f"已审计卡片 {len(rows)}/{LAST-FIRST+1} / 数据行 {sum(r['数据行数'] for r in rows)}")
if missing:
    print("[缺卡]", ",".join(map(str, missing)))
print(f"[阻断] {len(block)} 项")
for x in block:
    print(" -", x)
print(f"[提醒] {len(remind)} 项")
for x in remind:
    print(" -", x)
print(f"[Epit=N/A（待读图）] {len(na)} 行", na)
print(f"[15元素全未提及] {len(all_missing_elements)} 行", all_missing_elements)
print("审计表", OUT)
if missing or block:
    sys.exit(2)
print("audit11: OK")
