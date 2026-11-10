# -*- coding: utf-8 -*-
"""清洗脚本：0 填充缺失成分 -> 点蚀电位统一换算到 SCE -> 按工艺拆分到不同 sheet。
不增行、不删行、不加列。原表另存为 00_原始数据。"""
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

SRC = Path(r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx")
DST = SRC.with_name(SRC.stem + "_清洗版.xlsx")
BAK = Path(r"D:\dsh_temp\cast-ss-clean\backup")
BAK.mkdir(parents=True, exist_ok=True)

# 1) 备份原件到 D 盘
shutil.copy2(SRC, BAK / SRC.name)
print(f"备份原件 -> {BAK / SRC.name}")

# 2) 生成清洗版工作副本（原件不动）
shutil.copy2(SRC, DST)
print(f"工作副本 -> {DST}")

wb = openpyxl.load_workbook(DST)
ws = wb["Sheet1"]

ELEM_COLS = {  # 列号 -> 元素名
    3: "C", 4: "Si", 5: "N", 6: "P", 7: "S", 8: "Cu", 9: "Mn", 10: "Cr",
    11: "Mo", 12: "Ni", 13: "Ti", 14: "Nb", 15: "Co", 16: "Sn", 17: "Fe",
}
COL_RE, COL_EPIT, COL_REMARK, COL_PROC = 20, 24, 27, 18
FIRST_ROW = 3


def is_num(v):
    if v is None:
        return False
    try:
        float(str(v).replace("<", "").replace("≤", "").strip())
        return True
    except ValueError:
        return False


def proc_form(text):
    t = str(text or "")
    if "离心铸造" in t:
        return "3_离心铸造"
    if "熔模铸造" in t:
        return "4_熔模铸造"
    if "砂型铸造" in t:
        return "2_砂型铸造"
    if "真空精密铸造" in t or "真空电磁搅拌" in t:
        return "5_真空精密铸造"
    if "锻造" in t:
        return "8_锻造"
    if "热锻" in t or "冷轧" in t:
        return "7_热锻冷轧"
    if "铸态" in t:
        return "1_铸造未注明造型"
    if "熔炼" in t:
        return "6_熔炼铸锭"
    return "9_未注明形态"


# ---------- 3) 0 填充成分列 ----------
FILL_SET = {"", "未提及", "未提", "-", "—", "/"}
filled, kept_text = {}, {}
for r in range(FIRST_ROW, ws.max_row + 1):
    if ws.cell(row=r, column=2).value in (None, ""):
        continue
    for c, el in ELEM_COLS.items():
        v = ws.cell(row=r, column=c).value
        if is_num(v):
            continue
        s = str(v).strip() if v is not None else ""
        if s in FILL_SET:
            ws.cell(row=r, column=c).value = 0
            filled[el] = filled.get(el, 0) + 1
        else:
            kept_text.setdefault(el, []).append((r, s))

print("\n== 0 填充（原为空白/未提及）==")
print("  ", {k: filled[k] for k in ELEM_COLS.values() if filled.get(k)})
print(" 合计", sum(filled.values()), "个单元格")
print("\n== 保留原文未填 0（需人工决策）==")
for el, items in kept_text.items():
    from collections import Counter
    cc = Counter(s for _, s in items)
    print(f"  {el}: {dict(cc)}  行号示例 {[r for r, _ in items[:6]]}")

# ---------- 4) 点蚀电位统一换算到 SCE ----------
SHIFT_3M, SHIFT_SAT = -0.031, -0.044
stat = {"SCE不变": 0, "3M→SCE": 0, "sat→SCE": 0, "未注明按3M": 0, "N/A不动": 0}
raw_vs_sce = []
converted_rows = []
for r in range(FIRST_ROW, ws.max_row + 1):
    name = ws.cell(row=r, column=2).value
    if name in (None, ""):
        continue
    re_raw = str(ws.cell(row=r, column=COL_RE).value or "").strip()
    epit = ws.cell(row=r, column=COL_EPIT).value
    if not is_num(epit):
        stat["N/A不动"] += 1
        continue
    e = float(str(epit).strip())
    if re_raw.upper() == "SCE" or "饱和甘汞" in re_raw:
        new, label = e, "SCE不变"
        note = None
    elif "3M" in re_raw or "3 M" in re_raw:
        new, label = e + SHIFT_3M, "3M→SCE"
        note = "原参比 Ag/AgCl(3M KCl)，Epit 已换算至 SCE（-0.031 V）"
    elif "sat" in re_raw.lower():
        new, label = e + SHIFT_SAT, "sat→SCE"
        note = "原参比 Ag/AgCl(饱和KCl)，Epit 已换算至 SCE（-0.044 V）"
    else:  # 仅写 Ag/AgCl，未注明浓度
        new, label = e + SHIFT_3M, "未注明按3M"
        note = "原参比仅记 Ag/AgCl（原文未注明 KCl 浓度），按 3M KCl 换算至 SCE（-0.031 V）"
    stat[label] += 1
    ws.cell(row=r, column=COL_EPIT).value = round(new, 3)
    ws.cell(row=r, column=COL_RE).value = "SCE"
    if note:
        old = str(ws.cell(row=r, column=COL_REMARK).value or "").strip()
        ws.cell(row=r, column=COL_REMARK).value = (old + "；" + note) if old else note
        converted_rows.append(r)
    raw_vs_sce.append((r, round(e, 3), round(new, 3)))

print("\n== Epit 换算统计 ==")
for k, v in stat.items():
    print(f"  {k}: {v}")
print("  换算后范围(V vs SCE):",
      min(x[2] for x in raw_vs_sce), "~", max(x[2] for x in raw_vs_sce))

# ---------- 5) 原始表改名 + 生成清洗总表 ----------
ws.title = "00_原始数据"
clean = wb.copy_worksheet(ws)
clean.title = "01_清洗总表"
# 把清洗结果写入 01（前面已直接改的是 ws，需要重做一遍：先复制原值再统一处理）
# —— 简化：重新从原件读一遍原值，写到 01
src_wb = openpyxl.load_workbook(SRC, data_only=True)
src_ws = src_wb["Sheet1"]
for r in range(1, src_ws.max_row + 1):
    for c in range(1, src_ws.max_column + 1):
        clean.cell(row=r, column=c).value = src_ws.cell(row=r, column=c).value

for r in range(FIRST_ROW, clean.max_row + 1):
    if clean.cell(row=r, column=2).value in (None, ""):
        continue
    for c, el in ELEM_COLS.items():
        v = clean.cell(row=r, column=c).value
        if is_num(v):
            continue
        s = str(v).strip() if v is not None else ""
        if s in FILL_SET:
            clean.cell(row=r, column=c).value = 0
    re_raw = str(clean.cell(row=r, column=COL_RE).value or "").strip()
    epit = clean.cell(row=r, column=COL_EPIT).value
    if not is_num(epit):
        continue
    e = float(str(epit).strip())
    if re_raw.upper() == "SCE":
        new, note = e, None
    elif "3M" in re_raw:
        new, note = e + SHIFT_3M, "原参比 Ag/AgCl(3M KCl)，Epit 已换算至 SCE（-0.031 V）"
    elif "sat" in re_raw.lower():
        new, note = e + SHIFT_SAT, "原参比 Ag/AgCl(饱和KCl)，Epit 已换算至 SCE（-0.044 V）"
    else:
        new, note = e + SHIFT_3M, "原参比仅记 Ag/AgCl（原文未注明 KCl 浓度），按 3M KCl 换算至 SCE（-0.031 V）"
    clean.cell(row=r, column=COL_EPIT).value = round(new, 3)
    clean.cell(row=r, column=COL_RE).value = "SCE"
    if note:
        old = str(clean.cell(row=r, column=COL_REMARK).value or "").strip()
        clean.cell(row=r, column=COL_REMARK).value = (old + "；" + note) if old else note

# 原始数据 sheet 还原为原样（撤销第 3、4 步对 ws 的改动）
for r in range(1, src_ws.max_row + 1):
    for c in range(1, src_ws.max_column + 1):
        ws.cell(row=r, column=c).value = src_ws.cell(row=r, column=c).value

# ---------- 6) 按工艺拆分 sheet ----------
groups = {}
for r in range(FIRST_ROW, clean.max_row + 1):
    if clean.cell(row=r, column=2).value in (None, ""):
        continue
    groups.setdefault(proc_form(clean.cell(row=r, column=COL_PROC).value), []).append(r)

print("\n== 工艺分组 ==")
for k in sorted(groups):
    print(f"  {k}: {len(groups[k])} 条")

for name in sorted(groups):
    sh = wb.copy_worksheet(clean)
    sh.title = name
    keep = set(groups[name])
    for r in range(sh.max_row, FIRST_ROW - 1, -1):
        if r not in keep:
            sh.delete_rows(r)

# ---------- 7) 说明 sheet ----------
if "说明" in wb.sheetnames:
    del wb["说明"]
doc = wb.create_sheet("说明")
lines = [
    ["本文件由原始表清洗生成，原始表原样保留在 00_原始数据，未增行、未删行、未改列结构。"],
    [""],
    ["一、成分列 0 填充"],
    ["仅将原本为空白或写“未提及”的成分单元格改为 0（对应“不含该元素写 0”）。"],
    ["原为“余量/区间(如 17.0-19.0)/调整配比/微量/加N”的单元格保留原文，未强行改 0，避免把有效信息变成错误数值。"],
    [""],
    ["二、点蚀电位统一到 SCE（25 ℃，vs SHE 标准电位）"],
    ["E(SCE)=+0.241 V；E(Ag/AgCl,3M KCl)=+0.210 V；E(Ag/AgCl,饱和KCl)=+0.197 V。"],
    ["Ag/AgCl(3M KCl) → SCE：E_新 = E_原 - 0.031 V"],
    ["Ag/AgCl(饱和 KCl) → SCE：E_新 = E_原 - 0.044 V"],
    ["仅写 Ag/AgCl 未注明浓度（文献10 共 25 条，原文亦未写）：按 3M KCl 处理，-0.031 V；若按饱和则为 -0.044 V，两者相差 0.013 V。"],
    ["换算明细已写入各自行的“备注”列；参比电极列统一改为 SCE。"],
    ["温度差异对参比电位影响小于 0.005 V，忽略。"],
    [""],
    ["三、按工艺拆分 sheet"],
    ["拆分依据为原表“加工工艺”列中的成形方式；同一 worksheet 内保留全部原始列，可直接筛选热处理状态。"],
]
for i, row in enumerate(lines, start=1):
    doc.cell(row=i, column=1).value = row[0]
doc.column_dimensions["A"].width = 120
doc.cell(row=1, column=1).font = Font(bold=True)

# 删除空白的 Sheet2 / Sheet3
for s in ("Sheet2", "Sheet3"):
    if s in wb.sheetnames:
        del wb[s]

# 调整 sheet 顺序：00 原始 -> 01 清洗总表 -> 工艺 -> 说明
order = ["00_原始数据", "01_清洗总表"] + sorted(groups) + ["说明"]
wb._sheets = [wb[n] for n in order if n in wb.sheetnames]

wb.save(DST)
print(f"\n已保存: {DST}")
print("sheet 列表:", wb.sheetnames)
