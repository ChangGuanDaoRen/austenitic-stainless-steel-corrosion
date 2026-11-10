# -*- coding: utf-8 -*-
"""Fe 实算 v2：Fe = 100 - Σ(其余元素)。

关键改进：区分“原表真 0”和“原表写未提及/空白后被 0 填充”——后者若涉及主量元素
(C/Si/Mn/Cr/Mo/Ni)，Fe 会算偏，必须跳过而不是硬算。
用法: python calc_fe2.py [输出路径]
"""
import sys
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

SRC = Path(r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版.xlsx")
DST = Path(sys.argv[1]) if len(sys.argv) > 1 else SRC

EL = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni", "Ti", "Nb", "Co", "Sn"]  # 列 3..16
COL_OF = {name: 3 + i for i, name in enumerate(EL)}
MAJOR = ["C", "Si", "Mn", "Cr", "Ni"]          # 缺任一 -> Fe 不可信
COL_FE, COL_REMARK = 17, 27

RULE = {  # 文本 -> (数值, 说明)
    "<0.01": (0.01, "取上限"),
    "≤0.03": (0.03, "取上限"),
    "≤0.04": (0.04, "取上限"),
    "17.0-19.0": (18.0, "取区间中值"),
    "9.0-12.0": (10.5, "取区间中值"),
    "微量": (0.05, "文献20正文 N=0.05%"),
    "加N": (0.16, "文献20正文 0.05+0.11%"),
}
UNKNOWN_TEXT = {"未提及", "", None, "调整配比", "余量"}


def is_unknown_raw(v):
    if v is None:
        return True
    s = str(v).strip()
    return s in ("", "未提及", "调整配比")


def load_original():
    """返回 {行号: {元素: 原始文本}}，用于判断哪些 0 是填充出来的。"""
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["00_原始数据"]
    out = {}
    for r in range(3, ws.max_row + 1):
        if ws.cell(row=r, column=2).value in (None, ""):
            continue
        out[r] = {el: ws.cell(row=r, column=COL_OF[el]).value for el in EL}
    return out


def process(ws, raw_by_row, tag):
    if ws.cell(row=1, column=1).value is None:
        return 0, []
    computed, skipped, neg = 0, [], []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=2).value
        if name in (None, ""):
            continue
        raw = raw_by_row.get(r, {})
        total, unknown_major, unknown_minor, subs = 0.0, [], [], []
        for el in EL:
            cell = ws.cell(row=r, column=COL_OF[el])
            v = cell.value
            rawv = raw.get(el)
            if is_unknown_raw(rawv):
                # 原表没给这个元素，清洗时按用户要求填了 0；主量元素缺失则 Fe 不可信
                (unknown_major if el in MAJOR else unknown_minor).append(el)
                continue  # 保留已填的 0，不回退用户的 0 填充
            if isinstance(v, (int, float)):
                total += float(v)
                continue
            s = str(v).strip()
            if s in RULE:
                val, note = RULE[s]
                cell.value = val
                total += val
                subs.append(f"{el}:{s}->{val}({note})")
            else:
                try:
                    total += float(s)
                except ValueError:
                    (unknown_major if el in MAJOR else unknown_minor).append(el)
                    pass
        known_any = any(not is_unknown_raw(raw.get(el)) for el in EL)
        if unknown_major or not known_any:
            reason = "原表未给 " + "/".join(unknown_major) if unknown_major else "原表成分全为未提及"
            ws.cell(row=r, column=COL_FE).value = "余量"
            skipped.append((r, str(name), reason))
            continue
        fe = round(100.0 - total, 2)
        ws.cell(row=r, column=COL_FE).value = fe
        computed += 1
        if fe < 0:
            neg.append((r, str(name), fe))
        note = "Fe=100-Σ(已知14元素)=" + f"{fe}"
        if unknown_minor:
            note += "；原表未给 " + "/".join(unknown_minor) + " 按 0 计"
        if subs:
            note += "；" + "，".join(subs)
        old = ws.cell(row=r, column=COL_REMARK).value
        if not (isinstance(old, str) and "Fe=100-Σ" in old):
            ws.cell(row=r, column=COL_REMARK).value = ((old + "；") if old else "") + note
    print(f"[{tag}] 算出 Fe {computed} 条 / 跳过 {len(skipped)} 条")
    return computed, skipped


raw_map = load_original()
wb = openpyxl.load_workbook(SRC)
all_skip, nc = [], 0
for s in wb.sheetnames:
    if s in ("00_原始数据", "说明"):
        continue
    c, sk = process(wb[s], raw_map, s)
    nc += c
    if s == "01_清洗总表":
        all_skip = sk

print("\n跳过明细:")
for r, n, why in all_skip:
    print(f"  行{r} {n}: {why}")

ws = wb["01_清洗总表"]
fes = [ws.cell(row=r, column=COL_FE).value for r in range(3, ws.max_row + 1)]
num = [f for f in fes if isinstance(f, (int, float))]
print(f"\n01_清洗总表: 数值 Fe {len(num)} 条, 余量 {len(fes)-len(num)} 条")
print(f"Fe 范围 {min(num)} ~ {max(num)}；Fe<40 的 {sum(1 for f in num if f<40)} 条")
neg = [f for f in num if f < 0]
print("负值:", neg if neg else "无")
# 归一检查：Σ14 + Fe 应 = 100
bad = []
for r in range(3, ws.max_row + 1):
    fe = ws.cell(row=r, column=COL_FE).value
    if not isinstance(fe, (int, float)):
        continue
    s = sum(ws.cell(row=r, column=COL_OF[el]).value for el in EL
            if isinstance(ws.cell(row=r, column=COL_OF[el]).value, (int, float)))
    if abs(s + fe - 100) > 0.011:
        bad.append((r, round(s + fe, 2)))
print("Σ+Fe ≠ 100 的行:", bad if bad else "无")

doc = wb["说明"]
start = doc.max_row + 2
for i, t in enumerate([
    "五、Fe 实算（追加）",
    "Fe 由“余量”改为 Fe = 100 − Σ(其余 14 个元素)，保留 2 位小数；Σ+Fe = 100。",
    "文本折算口径：<0.01→0.01、≤0.03→0.03、≤0.04→0.04（取上限）；17.0-19.0→18.0、9.0-12.0→10.5（取区间中值）；微量→0.05、加N→0.16（文献20 正文：基础新成分 N=0.05%，含N 在此基础上多加入 0.11%）。",
    "原表写“未提及/空白”的元素已在清洗时按 0 填充（用户要求），Fe 计算时按 0 处理；但若缺的是主量元素(C/Si/Mn/Cr/Ni)则 Fe 不计算，保持“余量”，并在备注列写明缺哪些。",
    "Mo 多数文献未给，按 0 计（304 系 Mo 一般 ≤0.5 wt%，对 Fe 的影响不超过 0.5 wt%），逐行已在备注注明。",
    "文献20 的 4 条新成分试样，原文只写“调整 Cr、Ni 配比”无具体数值，Fe 无法计算；316LN 那条原表成分整行为“未提及”，同样保持“余量”。",
]):
    doc.cell(row=start + i, column=1).value = t

wb.save(DST)
print("\n已保存:", DST)


