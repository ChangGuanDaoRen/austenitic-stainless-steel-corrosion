# -*- coding: utf-8 -*-
"""清洗 v2：0 填充缺失成分 -> Epit 统一到 SCE -> 按工艺拆 sheet。
原件 Sheet1 原样保留为 00_原始数据；新表取消数据区合并单元格（表头合并保留），
不增行、不删行、不加列。"""
import shutil
import sys
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

SRC = Path(r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx")
DST = SRC.with_name(SRC.stem + "_清洗版.xlsx")
BAK = Path(r"D:\dsh_temp\cast-ss-clean\backup")
BAK.mkdir(parents=True, exist_ok=True)
shutil.copy2(SRC, BAK / SRC.name)
print(f"备份原件 -> {BAK / SRC.name}")

# ---------- 读原表 ----------
src_wb = openpyxl.load_workbook(SRC, data_only=True)
src = src_wb["Sheet1"]
MAXR, MAXC = src.max_row, src.max_column
NCOL = MAXC

# 表头（1、2 行）与列宽
header = [[src.cell(row=r, column=c).value for c in range(1, NCOL + 1)] for r in (1, 2)]
widths = {}
for c in range(1, NCOL + 1):
    dim = src.column_dimensions.get(get_column_letter(c))
    widths[c] = dim.width if dim is not None and dim.width else None

# 数据区合并单元格：把锚点值铺到块内每一行（保证取消合并后信息不丢）
fill_map = {}
head_merges = []
for m in src.merged_cells.ranges:
    if m.max_row <= 2:
        head_merges.append(m)
        continue
    v = src.cell(row=m.min_row, column=m.min_col).value
    for rr in range(m.min_row, m.max_row + 1):
        for cc in range(m.min_col, m.max_col + 1):
            fill_map[(rr, cc)] = v

rows = []
for r in range(3, MAXR + 1):
    vals = {}
    for c in range(1, NCOL + 1):
        v = src.cell(row=r, column=c).value
        if v is None and (r, c) in fill_map:
            v = fill_map[(r, c)]
        vals[c] = v
    if vals.get(2) in (None, ""):
        continue
    vals["_row"] = r
    rows.append(vals)

print(f"数据行数: {len(rows)}（原件 {MAXR - 2} 行数据区）")

# ---------- 0 填充 ----------
ELEM_COLS = {3: "C", 4: "Si", 5: "N", 6: "P", 7: "S", 8: "Cu", 9: "Mn", 10: "Cr",
             11: "Mo", 12: "Ni", 13: "Ti", 14: "Nb", 15: "Co", 16: "Sn", 17: "Fe"}
COL_RE, COL_EPIT, COL_REMARK, COL_PROC = 20, 24, 27, 18
FILL_SET = {"", "未提及", "未提", "-", "—", "/"}


def is_num(v):
    if v is None:
        return False
    try:
        float(str(v).replace("<", "").replace("≤", "").strip())
        return True
    except ValueError:
        return False


filled = Counter()
kept = {}
for row in rows:
    for c, el in ELEM_COLS.items():
        v = row[c]
        if is_num(v):
            continue
        s = str(v).strip() if v is not None else ""
        if s in FILL_SET:
            row[c] = 0
            filled[el] += 1
        else:
            kept.setdefault(el, Counter())[s] += 1

print("\n== 0 填充 ==")
print("  ", dict(filled), "  合计", sum(filled.values()), "格")
print("== 保留原文（未填 0）==")
for el, cc in kept.items():
    print(f"   {el}: {dict(cc)}")

# ---------- Epit 换算到 SCE ----------
S3M, SSAT = -0.031, -0.044
stat = Counter()
for row in rows:
    re_raw = str(row[COL_RE] or "").strip()
    epit = row[COL_EPIT]
    if not is_num(epit):
        stat["N/A 未动"] += 1
        continue
    e = float(str(epit).strip())
    if re_raw.upper() == "SCE":
        new, note, tag = e, None, "SCE 不变"
    elif "3M" in re_raw:
        new, tag = e + S3M, "Ag/AgCl(3M KCl)"
        note = "原参比 Ag/AgCl(3M KCl)，Epit 已换算至 SCE（-0.031 V）"
    elif "sat" in re_raw.lower():
        new, tag = e + SSAT, "Ag/AgCl(饱和 KCl)"
        note = "原参比 Ag/AgCl(饱和 KCl)，Epit 已换算至 SCE（-0.044 V）"
    else:
        new, tag = e + S3M, "Ag/AgCl(未注明→按3M)"
        note = "原参比仅记 Ag/AgCl（原文未注明 KCl 浓度），按 3M KCl 换算至 SCE（-0.031 V）"
    stat[tag] += 1
    row[COL_EPIT] = round(new, 3)
    row[COL_RE] = "SCE"
    if note:
        old = str(row[COL_REMARK] or "").strip()
        row[COL_REMARK] = (old + "；" + note) if old else note

num_vals = [float(r[COL_EPIT]) for r in rows if is_num(r[COL_EPIT])]
print("\n== Epit 换算 ==")
for k, v in stat.items():
    print(f"   {k}: {v}")
print(f"   换算后范围: {min(num_vals)} ~ {max(num_vals)} V (vs SCE)")

# ---------- 工艺分组 ----------
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


groups = {}
for row in rows:
    groups.setdefault(proc_form(row[COL_PROC]), []).append(row)
print("\n== 工艺分组 ==")
for k in sorted(groups):
    print(f"   {k}: {len(groups[k])} 条")
assert sum(len(v) for v in groups.values()) == len(rows)

# ---------- 写新工作簿（在原文件副本上追加 sheet） ----------
'''原 Sheet1 保留为 00_原始数据'''
wb = openpyxl.load_workbook(DST)
orig = wb["Sheet1"]
orig.title = "00_原始数据"

HDR_FILL = Font(bold=True)
WRAP = Alignment(vertical="center", wrap_text=True)


def write_sheet(name, data, with_head=True):
    sh = wb.create_sheet(name)
    r0 = 1
    if with_head:
        for i, hrow in enumerate(header, start=1):
            for c, v in enumerate(hrow, start=1):
                sh.cell(row=i, column=c).value = v
                sh.cell(row=i, column=c).font = HDR_FILL
                sh.cell(row=i, column=c).alignment = WRAP
        # 仅保留表头合并
        for m in head_merges:
            sh.merge_cells(start_row=m.min_row, start_column=m.min_col,
                           end_row=m.max_row, end_column=m.max_col)
        r0 = 3
    for i, row in enumerate(data):
        for c in range(1, NCOL + 1):
            sh.cell(row=r0 + i, column=c).value = row[c]
    for c in range(1, NCOL + 1):
        if widths.get(c):
            sh.column_dimensions[get_column_letter(c)].width = widths[c]
    sh.freeze_panes = f"A{r0}"
    return sh


write_sheet("01_清洗总表", rows)
for name in sorted(groups):
    write_sheet(name, groups[name])

# 说明 sheet
doc = wb.create_sheet("说明")
lines = [
    "数据清洗说明（由原始表生成，未增行、未删行、未增加列）",
    "",
    "一、成分列 0 填充",
    "只把原本是空白或写“未提及”的成分单元格改成 0，共 617 格，对应师兄提出的“不含那种元素直接写 0”。",
    "原本写“余量 / 区间(如 17.0-19.0) / 调整配比 / 微量 / 加N”的单元格没有强行改 0，保留原文，避免把有效信息变成错误数值：",
    "   Fe“余量” 120 条；Cr 区间 6 条 + 调整配比 4 条；Ni 区间 6 条 + 调整配比 4 条；N 微量 3 条 + 加N 1 条。",
    "",
    "二、点蚀电位统一到 SCE（25 ℃）",
    "换算常数：E(SCE)=+0.241 V，E(Ag/AgCl,3M KCl)=+0.210 V，E(Ag/AgCl,饱和 KCl)=+0.197 V（均 vs SHE）。",
    "   Ag/AgCl(3M KCl) → SCE：E_新 = E_原 - 0.031 V（13 条）",
    "   Ag/AgCl(饱和 KCl) → SCE：E_新 = E_原 - 0.044 V（12 条）",
    "   仅写 Ag/AgCl、未注明浓度：按 3M KCl 处理，-0.031 V（25 条，全部为文献10；原文亦未写明浓度，两种假设相差 0.013 V）",
    "   SCE：不变（69 条）；N/A：保持 N/A（2 条）",
    "换算明细逐条写在该行“备注”列，参比电极列统一显示 SCE。温度对参比电位影响 < 0.005 V，忽略。",
    "",
    "三、按工艺拆 sheet",
    "依据“加工工艺”列里的成形方式拆分：铸造未注明造型 46、砂型铸造 13、离心铸造 5、熔模铸造 6、真空精密铸造 6、熔炼铸锭 3、热锻冷轧 35、锻造 2、未注明形态 5。",
    "每个 sheet 保留全部原始列；热处理状态（固溶/冷轧态/敏化/钝化等）仍可在“加工工艺”列筛选。",
    "",
    "四、格式说明",
    "新表的表头合并与原件一致；数据区取消了合并单元格（文献编号/文献信息/备注按样本行逐行填充），便于筛选和后续建模读取。",
    "00_原始数据 为原件原样，含全部原始合并单元格，用于对照与回溯。",
]
for i, t in enumerate(lines, start=1):
    cell = doc.cell(row=i, column=1)
    cell.value = t
    if i == 1:
        cell.font = Font(bold=True, size=12)
doc.column_dimensions["A"].width = 130

for s in ("Sheet2", "Sheet3"):
    if s in wb.sheetnames:
        del wb[s]

order = ["00_原始数据", "01_清洗总表"] + sorted(groups) + ["说明"]
wb._sheets = [wb[n] for n in order if n in wb.sheetnames]
wb.save(DST)

print("\n已保存:", DST)
print("sheet:", wb.sheetnames)
