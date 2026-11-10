# -*- coding: utf-8 -*-
"""在“工艺 + 温度 + 参比 + 溶液”四维上统计同质簇大小，回答“看看哪种数据多”。"""
import sys
from collections import Counter, defaultdict

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
PATH = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx"
ws = openpyxl.load_workbook(PATH, data_only=True)["Sheet1"]

# 文献级已知信息（来自 PDF 原文扫描 + 师弟 10 月 16 日汇报）
SEAWATER = {"12", "13"}          # SCS14 / SCS13 用天然海水
UNKNOWN_CONC = {"6"}             # 文献6 为“不同浓度和 pH 的 NaCl”，浓度不固定
DUPLEX = {"18"}                  # Z3CN20-09M 为铸造双相不锈钢
COMPOSITE = {"17"}               # 304 + SiC 复合材料


def form(r):
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
        return "热锻/热轧+冷轧"
    if "铸态" in r:
        return "铸造(未注明造型)"
    if "熔炼" in r:
        return "熔炼铸锭(未注明铸/锻)"
    return "未注明"


def state(r):
    if "敏化" in r:
        return "固溶+敏化"
    if "稳定化" in r:
        return "固溶+稳定化"
    if "钝化" in r or "酸洗" in r:
        return "固溶+钝化"
    if "冷轧" in r and "固溶" not in r and "退火" not in r:
        return "冷轧态"
    if "固溶" in r:
        return "固溶"
    if "退火" in r:
        return "退火态"
    if "未热处理" in r or "无热处理" in r or "未固溶处理" in r:
        return "无热处理(铸态)"
    return "未注明"


def temp_group(t):
    t = str(t)
    if "未提及" in t:
        return "未提及"
    if any(k in t for k in ("室温", "约20", "20±2", "22±2", "22-23", "25℃", "28℃")):
        return "常温20-28℃"
    if "30" in t:
        return "30℃"
    if "40" in t:
        return "40℃"
    if "45" in t:
        return "45℃"
    if "14" in t:
        return "14℃"
    return t


def re_group(re):
    re = str(re).strip()
    if re == "SCE":
        return "SCE"
    if "3M" in re:
        return "Ag/AgCl(3M)"
    if "sat" in re:
        return "Ag/AgCl(sat)"
    return "Ag/AgCl(未注明)"


def solution(p):
    if p in SEAWATER:
        return "天然海水"
    if p in UNKNOWN_CONC:
        return "NaCl浓度不固定"
    return "3.5wt%NaCl"


rows, paper = [], None
for r in range(3, ws.max_row + 1):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    if a is not None and str(a).strip():
        paper = str(a).strip()
    if not b or not str(b).strip():
        continue
    raw = str(ws.cell(row=r, column=18).value or "")
    rows.append({
        "paper": paper, "name": str(b).strip(), "raw": raw,
        "form": form(raw), "state": state(raw),
        "T": temp_group(ws.cell(row=r, column=22).value),
        "RE": re_group(ws.cell(row=r, column=20).value),
        "sol": solution(paper),
        "Epit": ws.cell(row=r, column=24).value,
    })


def isnum(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


valid = [x for x in rows if isnum(x["Epit"])]
print(f"有数值 Epit 的条数: {len(valid)} / {len(rows)}")

print("\n== 维度分布（用于说明清洗后的损失）==")
for key, label in (("form", "成形方式"), ("state", "热处理状态"), ("T", "温度组"),
                   ("RE", "参比"), ("sol", "溶液")):
    print(f"\n-- {label} --")
    for k, v in Counter(x[key] for x in valid).most_common():
        print(f"   {k}: {v}")

def top(keys, n=8, title=""):
    c = Counter(tuple(x[k] for k in keys) for x in valid)
    print(f"\n== 簇大小：{' + '.join(keys)} {title} ==")
    for k, v in c.most_common(n):
        print(f"   {v:>3} 条 | " + " | ".join(k))

top(["form", "state"])
top(["form", "state", "T"])
top(["form", "state", "T", "RE"])
top(["form", "state", "T", "RE", "sol"])

print("\n== 推荐方案 A：只留铸造系 + 固溶态 ==")
cast = [x for x in valid if x["form"] in ("铸造(未注明造型)", "砂型铸造", "离心铸造", "熔模铸造", "真空精密铸造", "熔炼铸锭(未注明铸/锻)") and x["state"] == "固溶"]
print(f"   条数: {len(cast)}")
print("   其中温度组:", dict(Counter(x['T'] for x in cast)))
print("   其中参比:", dict(Counter(x['RE'] for x in cast)))

print("\n== 推荐方案 B：铸造系 全部热处理状态 ==")
cast_all = [x for x in valid if x["form"] in ("铸造(未注明造型)", "砂型铸造", "离心铸造", "熔模铸造", "真空精密铸造", "熔炼铸锭(未注明铸/锻)")]
print(f"   条数: {len(cast_all)}")

print("\n== 需要单独标记 / 可能剔除的 ==\n")
for tag, ps in (("天然海水", SEAWATER), ("浓度不固定", UNKNOWN_CONC), ("铸造双相", DUPLEX), ("SiC复合材料", COMPOSITE)):
    sub = [x for x in valid if x["paper"] in ps]
    print(f"   {tag}: {len(sub)} 条 —— " + ", ".join(sorted({f"文献{x['paper']}" for x in sub})))
