# -*- coding: utf-8 -*-
"""
0.xlsx 第二轮复核：把文献22–26 的条目与 PDF 原文提取值逐格比对。
参考值来源（2026-10-28 实测，PDF 页码见核查报告）：
  22.pdf=C021 p2 表1 / p4 表3；23.pdf=C034 p45 表3-1 / p59 表4-1 / p47-48；
  24.pdf=Chao2014 p2 表I / p4 表II；25.pdf=204Cu p2 表1 / p3 正文；
  26.pdf=HNASS p2 表1 / p14 表3
用法：python scripts/verify_0xlsx.py
"""
import openpyxl

# 列号：3 C, 4 Si, 5 N, 6 P, 7 S, 8 Cu, 9 Mn, 10 Cr, 11 Mo, 12 Ni, 15 Co, 24 Epit
COLS = {"C": 3, "Si": 4, "N": 5, "P": 6, "S": 7, "Cu": 8, "Mn": 9,
        "Cr": 10, "Mo": 11, "Ni": 12, "Co": 15, "Epit": 24}

# 文献26 的 Si 与 Epit 网格（Epit 单位 V，原文 mV vs SCE）
HNASS_SI = 0.34
HNASS_BASE = dict(C=0.076, N=0.543, P=0.051, S=0.007, Mn=19.78, Cr=17.96, Mo=0, Ni=0)
EPIT_GRID = {
    "550": [0.164, 0.136, 0.074, 0.045, -0.010, -0.039],
    "650": [0.109, 0.047, 0.004, -0.059, -0.104, -0.163],
    "750": [-0.035, -0.011, -0.096, -0.114, -0.224, -0.157],
    "850": [-0.045, -0.027, -0.128, -0.146, -0.252, -0.153],
}
TIMES = ["30min", "60min", "180min", "360min", "720min", "1440min"]

REF = {}

# 文献22
REF["NHG1"] = dict(C=0.02, Si=3.65, P=0.02, S=0.02, Cu=3.04, Mn=0.28,
                   Cr=16.44, Mo=3.42, Ni=16.22, Epit=1.102)
REF["316L"] = dict(C=0.03, Si=0.15, P=0.03, S=0.03, Cu=0, Mn=0.32,
                   Cr=18.36, Mo=3.04, Ni=12.58, Epit=0.805)
# 文献23
C034 = dict(C=0.024, Si=1.09, N=0.033, P=0.023, S=0.0039, Cu=0.031, Mn=1.11,
            Cr=20.16, Mo=0.26, Ni=9.06, Co=0.026)
for h, e in zip(["0", "500", "1000", "2000", "3000"], [0.4189, 0.3822, 0.3537, 0.2704, 0.1328]):
    REF[f"Z3CN20.09M ({h} h)"] = dict(C034, **{"Epit": e})
# 文献24（Sample A/B 已按用户裁决删除，只剩 C）
REF["Sample C (Fe-16Cr-2Co-14.50Mn-1.85Mo-0.73N)"] = dict(
    C=0.031, Si=0.44, N=0.73, P=0.01, S=0.007, Mn=14.5, Cr=16.78, Mo=1.85,
    Ni=0, Co=1.83, Epit=0.376)
# 文献25
for b, comp, e in [
    ("B1 (204Cu, 0 Ce)", (0.062, 0.38, 0.15, 8.157, 16.48, 1.99, 2.42), -0.97),
    ("B2 (204Cu, 0.007Ce)", (0.065, 0.36, 0.17, 8.162, 16.52, 1.97, 2.45), -0.94),
    ("B3 (204Cu, 0.018Ce)", (0.059, 0.37, 0.15, 8.148, 16.61, 1.98, 2.44), -0.82),
    ("B4 (204Cu, 0.025Ce)", (0.057, 0.41, 0.16, 8.152, 16.58, 1.97, 2.42), -0.89),
]:
    c, si, n, mn, cr, ni, cu = comp
    REF[b] = dict(C=c, Si=si, N=n, P=0.02, S=0.006 if b.startswith("B1") else 0.005,
                  Cu=cu, Mn=mn, Cr=cr, Ni=ni, Epit=e)
# 文献26
for temp, vals in EPIT_GRID.items():
    for t, e in zip(TIMES, vals):
        REF[f"HNASS-{temp}℃-{t}"] = dict(HNASS_BASE, **{"Si": HNASS_SI, "Epit": e})


def norm(v):
    if v is None:
        return None
    if isinstance(v, str):
        return v.strip()
    return v


def eq(a, b):
    if a is None and b is None:
        return True
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return str(a) == str(b)


wb = openpyxl.load_workbook("0.xlsx")
ws = wb["Sheet1"]

# 已删除样本不应再出现
MUST_ABSENT = {"Lewmet55", "C1", "C2", "C3", "C4", "C5", "C6", "C7"}
MUST_ABSENT_PREFIX = ("Sample A (", "Sample B (")

found = {}
# REF 只覆盖文献22–26（行 124–159）；表内其它块也有叫「316L」的样本（行 239/240），
# 全表建索引会被后者覆盖，所以取值比对限定在本区间。
found_scope = {}
for r in range(3, ws.max_row + 1):
    name = ws.cell(row=r, column=2).value
    if not name:
        continue
    found[str(name).strip()] = r
    if 124 <= r <= 159:
        found_scope[str(name).strip()] = r

problems = []
for name in MUST_ABSENT:
    if name in found:
        problems.append(f"应删未删：{name}（行 {found[name]}）")
for name in found:
    if name.startswith(MUST_ABSENT_PREFIX):
        problems.append(f"应删未删：{name}（行 {found[name]}）")

checked = 0
for name, ref in REF.items():
    if name not in found_scope:
        problems.append(f"缺行：{name}")
        continue
    r = found_scope[name]
    for field, want in ref.items():
        got = norm(ws.cell(row=r, column=COLS[field]).value)
        if not eq(got, want):
            problems.append(f"不符：{name} / {field}  表内 {got!r} ≠ 原文 {want!r}")
    checked += 1

print(f"已比对样本数：{checked} / 参考条目 {len(REF)}")
if problems:
    print("发现问题：")
    for p in problems:
        print("  -", p)
else:
    print("全部一致：文献22–26 的成分、Epit 与原文逐格相符，应删样本已全部删除。")
