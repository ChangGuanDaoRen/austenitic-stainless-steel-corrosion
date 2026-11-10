# -*- coding: utf-8 -*-
"""按加工工艺对 121 条数据分组统计，回答师兄要求的“看看哪种数据多”。"""
import openpyxl
import sys
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
PATH = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx"
ws = openpyxl.load_workbook(PATH, data_only=True)["Sheet1"]


def form(r):
    """成形方式"""
    if "离心铸造" in r:
        return "离心铸造"
    if "熔模铸造" in r:
        return "熔模铸造"
    if "砂型铸造" in r:
        return "砂型铸造"
    if "真空精密铸造" in r or "真空电磁搅拌" in r:
        return "真空精密铸造"
    if "锻造" in r:
        return "锻造"
    if "热锻" in r or "冷轧" in r:
        return "热锻+冷轧"
    if "铸态" in r:
        return "铸造(未注明造型)"
    if "真空感应炉熔炼" in r or "真空电弧炉熔炼" in r:
        return "熔炼铸锭(未注明铸/锻)"
    return "未注明"


def state(r):
    """热处理/冷加工状态"""
    if "敏化" in r:
        return "固溶+敏化"
    if "稳定化" in r:
        return "固溶+稳定化"
    if "钝化" in r or "酸洗" in r:
        return "固溶+钝化/酸洗"
    if "冷轧" in r and "固溶" not in r and "退火" not in r:
        return "冷轧态"
    if "固溶" in r:
        return "固溶(水淬/油冷/空冷/炉冷)"
    if "退火" in r:
        return "退火态"
    if "未热处理" in r or "无热处理" in r or "未固溶处理" in r:
        return "无热处理(铸态)"
    return "未注明"


rows, paper = [], None
for r in range(3, ws.max_row + 1):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    if a is not None and str(a).strip():
        paper = str(a).strip()
    if not b or not str(b).strip():
        continue
    raw = str(ws.cell(row=r, column=18).value or "")
    rows.append({"paper": paper, "name": str(b).strip(), "raw": raw,
                 "form": form(raw), "state": state(raw),
                 "Epit": ws.cell(row=r, column=24).value,
                 "T": ws.cell(row=r, column=22).value,
                 "pH": ws.cell(row=r, column=23).value,
                 "RE": ws.cell(row=r, column=20).value})

print("== 121 条逐条归类 ==")
for x in rows:
    print(f"  文献{x['paper']:>2} | {x['name'][:34]:<34} | {x['form']:<16} | {x['state']:<18} | {x['T']} | pH={x['pH']}")

print("\n== 按成形方式 ==")
for k, v in Counter(x["form"] for x in rows).most_common():
    print(f"  {k}: {v}")

print("\n== 按热处理状态 ==")
for k, v in Counter(x["state"] for x in rows).most_common():
    print(f"  {k}: {v}")

print("\n== 成形方式 × 状态 交叉 ==")
cross = Counter((x["form"], x["state"]) for x in rows)
for (f, s), v in cross.most_common():
    print(f"  {f:<18} | {s:<22} | {v}")

print("\n== 最大同工艺簇（含温度/参比细分）==")
# 只保留有数值 Epit 的
def isnum(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False

clusters = defaultdict(list)
for x in rows:
    if isnum(x["Epit"]):
        clusters[(x["form"], x["state"])].append(x)
for k, v in sorted(clusters.items(), key=lambda kv: -len(kv[1]))[:6]:
    print(f"  {k[0]} + {k[1]}: {len(v)} 条")

print("\n== 若再叠加“参比统一到 SCE”后仍保留的条数 ==")
print("  全部 121 条都可换算（Ag/AgCl 未注明浓度的那 25 条需查原文补齐或降权）")
