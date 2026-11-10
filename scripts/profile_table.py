# -*- coding: utf-8 -*-
"""对最后一次提交给师兄的原始表格做只读统计，支撑数据清洗决策。"""
import openpyxl
import sys
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")

PATH = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx"

wb = openpyxl.load_workbook(PATH, data_only=True)
ws = wb["Sheet1"]

# 列索引（1-based）
COL = {"A": 1, "B": 2, "RE": 20, "T": 22, "pH": 23, "Epit": 24, "stab": 25, "Z": 26, "note": 27}
ELEMENTS = {"C": 3, "Si": 4, "N": 5, "P": 6, "S": 7, "Cu": 8, "Mn": 9, "Cr": 10,
            "Mo": 11, "Ni": 12, "Ti": 13, "Nb": 14, "Co": 15, "Sn": 16, "Fe": 17}

rows = []
paper = None
for r in range(3, ws.max_row + 1):
    a = ws.cell(row=r, column=COL["A"]).value
    b = ws.cell(row=r, column=COL["B"]).value
    if a is not None and str(a).strip() != "":
        paper = str(a).strip()
    if b is None or str(b).strip() == "":
        continue
    rec = {"row": r, "paper": paper, "name": str(b).strip()}
    for k, c in COL.items():
        if k in ("A", "B"):
            continue
        rec[k] = ws.cell(row=r, column=c).value
    for el, c in ELEMENTS.items():
        rec[el] = ws.cell(row=r, column=c).value
    rows.append(rec)

print("== 数据行总数（不含表头）:", len(rows))
print("== 涉及文献编号数:", len(set(x["paper"] for x in rows)))
cnt = Counter(x["paper"] for x in rows)
print("== 每篇文献条数:", dict(sorted(cnt.items(), key=lambda kv: int(kv[0]))))

print("\n== 参比电极 RE 分布 ==")
for k, v in Counter(str(x["RE"]).strip() for x in rows).most_common():
    print(f"  {k}: {v}")
    for x in rows:
        if str(x["RE"]).strip() == k:
            print(f"      文献{x['paper']} {x['name']}")

print("\n== Epit 可用性 ==")
num, na, other = 0, 0, []
vals = []
for x in rows:
    v = x["Epit"]
    try:
        f = float(v)
        num += 1
        vals.append((x["paper"], x["name"], f))
    except (TypeError, ValueError):
        s = str(v).strip() if v is not None else ""
        if s.upper() == "N/A":
            na += 1
        else:
            other.append((x["paper"], x["name"], s))
print(f"  数值型: {num}   N/A: {na}   其他/空: {len(other)}")
if other:
    for o in other:
        print("     ", o)
print("  Epit 数值范围(V):", min(v[2] for v in vals), "~", max(v[2] for v in vals))

def is_number(s):
    try:
        float(str(s).replace("<", "").replace("≤", "").strip())
        return True
    except ValueError:
        return False


print("\n== 元素成分 缺失情况 ==")
for el, c in ELEMENTS.items():
    c2 = Counter(str(x[el]).strip() if x[el] is not None else "空" for x in rows)
    bad = [(k, v) for k, v in c2.items() if not is_number(k)]
    if bad:
        print(f"  {el}: " + ", ".join(f"{k}×{v}" for k, v in sorted(bad, key=lambda kv: -kv[1])))

print("\n== 温度分布 ==")
for k, v in Counter(str(x["T"]).strip() for x in rows).most_common():
    print(f"  {k}: {v}")

print("\n== pH 分布 ==")
for k, v in Counter(str(x["pH"]).strip() for x in rows).most_common():
    print(f"  {k}: {v}")

print("\n== 测试前稳定时间 ==")
for k, v in Counter(str(x["stab"]).strip() for x in rows).most_common():
    print(f"  {k}: {v}")
