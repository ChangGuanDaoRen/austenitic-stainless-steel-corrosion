# -*- coding: utf-8 -*-
"""
从 0.xlsx 生成 1.xlsx（清洗第一版）：
  1) 元素列 0 填充：「未提及」→ 0；带比较符按取上限；区间取中值；Fe = 100 − Σ其余元素
  2) Epit 统一到 SCE，并删掉没有 Epit 数据（N/A）的样本
  3) 非元素列统一为数值：温度→℃、pH→数值、稳定时间→分钟、对电极→铂/石墨
  4) 剔除文献20（原文无元素成分），删除「文献信息」「备注」两列，「文献编号」逐行填值（不合并，保证每条自足）
  5) 样式照 0.xlsx：以 0.xlsx 为模板复制后删列删行，字体/边框/对齐/列宽/表头合并全部沿用
  6) 统计条数，写入「统计」页

换算常数（vs SHE，25 ℃）：E(SCE)=+0.241，E(Ag/AgCl,3M KCl)=+0.210，E(Ag/AgCl,饱和 KCl)=+0.197

用法：python scripts/build_1.py [dry|write]
"""
import openpyxl
import re
import shutil
import sys
from collections import Counter
from copy import copy as style_copy

MODE = sys.argv[1] if len(sys.argv) > 1 else "dry"
SRC, DST = "0.xlsx", "1.xlsx"

COL_ID, COL_NAME = 1, 2
ELEM_COLS = list(range(3, 18))
EL_NO_FE = list(range(3, 17))
MAJOR_COLS = {3: "C", 4: "Si", 9: "Mn", 10: "Cr", 12: "Ni"}
COL_FE = 17
COL_PROC, COL_WE, COL_RE, COL_CE = 18, 19, 20, 21
COL_TEMP, COL_PH, COL_EPIT, COL_STAB = 22, 23, 24, 25
DROP_COLS = [27, 26]                     # 备注、文献信息

D3M, DSAT = -0.031, -0.044
DROP_PAPER = 20

DEF_TEMP, DEF_PH, DEF_STAB, DEF_CE = 25, 7, 30, "铂"


def num(x):
    f = float(x)
    return int(f) if f == int(f) else round(f, 2)


def fill_element(v):
    if v is None:
        return 0, "空→0"
    if isinstance(v, (int, float)):
        return v, None
    s = str(v).strip()
    if s in ("未提及", "", "—", "-"):
        return 0, "未提及→0"
    m = re.fullmatch(r"[<≤]\s*([0-9.]+)", s)
    if m:
        return float(m.group(1)), f"{s}→取上限"
    m = re.fullmatch(r"([0-9.]+)\s*[-~]\s*([0-9.]+)", s)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        return round((a + b) / 2, 4), f"{s}→取中值"
    m = re.fullmatch(r"([0-9.]+)", s)
    if m:
        return float(m.group(1)), None
    return s, "元素无法折算(保留原文)"


def unify(v, c):
    s = "" if v is None else str(v).strip()
    missing = s in ("", "未提及")

    if c == COL_TEMP:
        if missing or s == "室温":
            return DEF_TEMP, "腐蚀温度→25（缺失/室温按 25℃）"
        m = re.fullmatch(r"约?\s*([0-9.]+)\s*±\s*[0-9.]+\s*℃?", s)
        if m:
            return num(m.group(1)), None
        m = re.fullmatch(r"([0-9.]+)\s*[-~]\s*([0-9.]+)\s*℃?", s)
        if m:
            return num((float(m.group(1)) + float(m.group(2))) / 2), None
        m = re.fullmatch(r"约?\s*([0-9.]+)\s*℃", s)
        if m:
            return num(m.group(1)), None
        return v, f"温度写法未识别({s})"

    if c == COL_PH:
        if missing or s == "中性":
            return DEF_PH, "腐蚀pH→7（缺失/中性）"
        m = re.fullmatch(r"([0-9.]+)\s*[-~]\s*([0-9.]+)", s)
        if m:
            return num((float(m.group(1)) + float(m.group(2))) / 2), None
        try:
            return num(s), None
        except ValueError:
            return v, f"pH 写法未识别({s})"

    if c == COL_STAB:
        if missing:
            return DEF_STAB, "稳定时间→30min（缺失）"
        m = re.fullmatch(r"约?\s*([0-9.]+)\s*min", s)
        if m:
            return num(m.group(1)), None
        m = re.fullmatch(r"约?\s*([0-9.]+)\s*h", s)
        if m:
            return num(float(m.group(1)) * 60), None
        m = re.fullmatch(r"([0-9.]+)\s*s", s)
        if m:
            return num(float(m.group(1)) / 60), None
        return v, f"稳定时间写法未识别({s})"

    if c == COL_CE:
        if missing:
            return DEF_CE, "对电极→铂（缺失，多数文献用铂且不参与电位测量）"
        if "石墨" in s:
            return "石墨", None
        if "铂" in s:
            return "铂", (None if s == "铂" else f"对电极 {s}→铂")
        return v, f"对电极写法未识别({s})"

    return v, None


def conv_epit(v, re_raw):
    s = "" if re_raw is None else str(re_raw).strip()
    if isinstance(v, str) and v.strip().upper() == "N/A":
        return "N/A", "N/A"
    if not isinstance(v, (int, float)):
        return v, None
    su = s.upper()
    if su == "SCE":
        return round(v, 6), "SCE 不变"
    if "3M" in su or "3 M" in su:
        return round(v + D3M, 6), "Ag/AgCl(3M KCl)→SCE"
    if "SAT" in su or "饱和" in s:
        return round(v + DSAT, 6), "Ag/AgCl(饱和 KCl)→SCE"
    if su.startswith("AG/AGCL"):
        return round(v + D3M, 6), "Ag/AgCl(未注明浓度→按3M)→SCE"
    return v, f"参比未识别({s})"


# ---------- 阶段 1：读 0.xlsx 生成清洗后的行 ----------
src = openpyxl.load_workbook(SRC)
ws = src["Sheet1"]

rows, stats, conv_cnt, fill_cnt = [], Counter(), Counter(), 0
dropped20, dropped_na, unconverted, fe_unknown = [], [], [], []
delete_rows = []

cur_id = None
for r in range(3, ws.max_row + 1):
    raw_id = ws.cell(row=r, column=COL_ID).value
    if raw_id:
        cur_id = raw_id
    name = ws.cell(row=r, column=COL_NAME).value
    if not name:
        continue
    if str(cur_id) == str(DROP_PAPER):
        dropped20.append((r, f"文献{cur_id} {name}"))
        delete_rows.append(r)
        continue
    epit_raw = ws.cell(row=r, column=COL_EPIT).value
    if epit_raw is None or (isinstance(epit_raw, str) and epit_raw.strip().upper() == "N/A"):
        dropped_na.append((r, f"文献{cur_id} {name} Epit={epit_raw}"))
        delete_rows.append(r)
        continue
    new = {}
    for c in range(1, 26):
        v = ws.cell(row=r, column=c).value
        if c == COL_ID:
            new[c] = cur_id                      # 文献编号逐行填值
        elif c in ELEM_COLS and c != COL_FE:
            nv, tag = fill_element(v)
            if tag:
                fill_cnt += 1
                stats[tag] += 1
            new[c] = nv
        elif c == COL_EPIT:
            nv, tag = conv_epit(v, ws.cell(row=r, column=COL_RE).value)
            conv_cnt[tag] += 1
            if tag and tag.startswith("参比未识别"):
                unconverted.append((r, str(name), ws.cell(row=r, column=COL_RE).value))
            new[c] = nv
        elif c == COL_RE:
            new[c] = "SCE"
        else:
            nv, tag = unify(v, c)
            if tag:
                stats[tag] += 1
            new[c] = nv
    feat = ws.cell(row=r, column=COL_FE).value
    if isinstance(feat, (int, float)) and not isinstance(feat, bool):
        new[COL_FE] = feat
    else:
        miss = [MAJOR_COLS[c] for c in MAJOR_COLS
                if ws.cell(row=r, column=c).value is None
                or str(ws.cell(row=r, column=c).value).strip() in ("未提及", "", "调整配比")]
        if miss:
            new[COL_FE] = "余量"
            fe_unknown.append((r, str(name), "缺" + "/".join(miss)))
        else:
            new[COL_FE] = round(100 - sum(float(new[c]) for c in EL_NO_FE
                                          if isinstance(new[c], (int, float))), 4)
            stats["Fe 实算"] += 1
    rows.append(new)

print(f"生成行数：{len(rows)}")
print(f"剔除文献20（原文无成分）：{len(dropped20)} 条 -> {[d[1] for d in dropped20]}")
print(f"删除无 Epit 数据：{len(dropped_na)} 条 -> {[d[1] for d in dropped_na]}")
print(f"元素列处理：{fill_cnt} 格；Fe 实算 {stats['Fe 实算']} 条")
print(f"Epit 换算：{dict(conv_cnt)}")
if unconverted:
    print("参比未识别：", unconverted)
if fe_unknown:
    print("Fe 未实算：", fe_unknown)
cnt = Counter(x[COL_ID] for x in rows)
print("各文献条数：", dict(sorted(cnt.items(), key=lambda kv: int(kv[0]))))

if MODE != "write":
    print("\n[dry] 未写盘")
    sys.exit(0)

# ---------- 阶段 2：以 0.xlsx 为模板生成 1.xlsx（样式沿用） ----------
shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
o = wb["Sheet1"]

merges = [str(m) for m in o.merged_cells.ranges]
for m in merges:
    o.unmerge_cells(m)

for c in DROP_COLS:                       # 从右往左删列
    o.delete_cols(c)

for r in sorted(delete_rows, reverse=True):
    o.delete_rows(r)

# 写入清洗后的值（表头两行文字沿用模板，仅改单位标注）
for i, row in enumerate(rows, start=3):
    for c in range(1, 26):
        o.cell(row=i, column=c).value = row.get(c)
o.cell(row=1, column=COL_TEMP).value = "腐蚀温度(℃)"
o.cell(row=1, column=COL_STAB).value = "测试前稳定时间(min)"

# 重建表头（第 1–2 行）的合并区，数据区不合并（文献编号已逐行填值）
rebuilt = 0
for m in merges:
    cr = openpyxl.worksheet.cell_range.CellRange(m)
    if cr.max_row <= 2 and cr.max_col <= 25:
        o.merge_cells(str(cr))
        rebuilt += 1

for s in ("Sheet2", "Sheet3"):
    if s in wb.sheetnames:
        del wb[s]
if "统计" in wb.sheetnames:
    del wb["统计"]
st = wb.create_sheet("统计")
st.append(["项目", "数值"])
st.append(["总条数", len(rows)])
st.append(["覆盖文献数", len(cnt)])
st.append(["剔除（文献20，原文无成分）", len(dropped20)])
st.append(["删除（无 Epit 数据）", len(dropped_na)])
st.append(["元素列 0 填充格数", fill_cnt])
st.append(["Fe 实算条数", stats["Fe 实算"]])
st.append(["—", "—"])
st.append(["非元素列统一", "格数"])
for k, v in stats.items():
    if k.startswith(("腐蚀温度", "腐蚀pH", "稳定时间", "对电极")):
        st.append([k, v])
st.append(["—", "—"])
for k, v in conv_cnt.most_common():
    st.append([f"Epit：{k}", v])
st.append(["—", "—"])
st.append(["文献编号", "条数"])
for k in sorted(cnt, key=lambda v: int(v)):
    st.append([k, cnt[k]])

wb.save(DST)
print(f"\n[write] 已写盘 {DST}：{len(rows)} 条；重建表头合并 {rebuilt} 处；sheets={wb.sheetnames}")
