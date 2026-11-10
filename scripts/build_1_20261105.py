# -*- coding: utf-8 -*-
r"""从 0.xlsx（259 条，2026-11-05 删酸后基准）重建 1.xlsx（清洗版）。

口径沿用 2026-10 版 build_1.py，按新入库数据的写法补齐规则：
  1) 删 文献20（原文无成分，5 条）与 Epit=N/A（3 条） -> 251 条
  2) 元素：未提及/空 -> 0；<x/≤x -> x；a-b -> 中值；ppm -> /1e4
  3) Fe：原文实测保留；其余 = 100 − Σ14元素（保留 2 位）
  4) Epit 统一 SCE：SCE Δ=0；Ag/AgCl(未注明或3M) Δ=−0.031；
     Ag/AgCl(饱和) Δ=−0.042；文献55 两行备注注明已按 −0.042 换算 -> Δ=0
  5) 温度：室温/未提及 -> 25；22±2 -> 22；区间 -> 中值；298 K -> 25
  7) 稳定时间：统一分钟；未提及 -> 30；「预浸10 min + OCP 3600 s」-> 70
  8) 对电极：铂族 -> 铂；石墨族 -> 石墨；未提及 -> 铂
  9) 删「文献信息」「备注」两列；文献编号逐行填值

输出 sheet：Sheet1（24 列，已删 pH 列）/ 统计 / 清洗标记 / 未纳入
用法：python scripts/build_1_20261105.py [dry|write]
"""
import os
import re
import shutil
import sys
from copy import copy

import openpyxl

SRC, DST = "0.xlsx", "1.xlsx"
PREVIEW = os.path.join("temp", "audit0", "1_预览.xlsx")
BAK = os.path.join("备份", "1_旧版_150条_20261105.xlsx")

C_ID, C_NAME = 1, 2
C_ELEM = list(range(3, 17))       # C..Sn
C_FE = 17
C_PROC, C_WE, C_RE, C_CE = 18, 19, 20, 21
C_TEMP, C_EPIT, C_STAB = 22, 23, 24
DROP_PAPER = 20

# 参比 -> (ΔV 换到 SCE, 标记)
RE_MAP = {
    "SCE": (0.0, "SCE"),
    "SCE(饱和KCl甘汞)": (0.0, "SCE"),
    "SCE(饱和甘汞)": (0.0, "SCE"),
    "Ag/AgCl": (-0.031, "Ag/AgCl(未注明->按3M)"),
    "Ag/AgCl(3M KCl)": (-0.031, "Ag/AgCl(3M KCl)"),
    "Ag/AgCl(sat.KCl)": (-0.042, "Ag/AgCl(饱和KCl)"),
    "Ag/AgCl(饱和 KCl)": (-0.042, "Ag/AgCl(饱和KCl)"),
    "Ag/AgCl(饱和 KCl, 0.1989 V vs SHE)": (-0.0421, "Ag/AgCl(饱和,0.1989V vs SHE)"),
    "Ag/AgCl(饱和KCl)→SCE": (0.0, "Ag/AgCl(饱和) 原文已换算"),
}

NUM_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*$")
CAP_RE = re.compile(r"^\s*[<≤]\s*([0-9]+(?:\.[0-9]+)?)\s*$")
RANGE_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*[-~–]\s*([0-9]+(?:\.[0-9]+)?)\s*$")
PPM_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:ppm|PPM)\s*$")


def r2(x, n=2):
    v = round(float(x), n)
    return int(v) if v == int(v) else v


def parse_elem(v):
    """返回 (数值, 标签)。标签：num / 0(未提及) / cap / mid / ppm / unknown"""
    if v is None:
        return 0.0, "0(空)"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v), "num"
    s = str(v).strip()
    if s in ("", "未提及", "—", "-", "None"):
        return 0.0, "0(未提及)"
    m = CAP_RE.match(s)
    if m:
        return float(m.group(1)), "cap"
    m = RANGE_RE.match(s)
    if m:
        return round((float(m.group(1)) + float(m.group(2))) / 2, 4), "mid"
    m = PPM_RE.match(s)
    if m:
        return round(float(m.group(1)) / 1e4, 6), "ppm"
    m = NUM_RE.match(s)
    if m:
        return float(m.group(1)), "num"
    return 0.0, "unknown:" + s


def parse_temp(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "室温", "None"):
        return 25, "默认25"
    m = re.fullmatch(r"约?\s*([0-9.]+)\s*±\s*[0-9.]+\s*℃", s)
    if m:
        return r2(m.group(1)), "原文"
    m = re.fullmatch(r"([0-9.]+)\s*[-~–]\s*([0-9.]+)\s*℃", s)
    if m:
        return r2((float(m.group(1)) + float(m.group(2))) / 2), "原文"
    m = re.fullmatch(r"约?\s*([0-9.]+)\s*℃", s)
    if m:
        return r2(m.group(1)), "原文"
    m = re.fullmatch(r"([0-9.]+)\s*K", s)
    if m:
        return int(round(float(m.group(1)) - 273.15)), "原文(K)"
    return None, "未识别:" + s


def parse_ph(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return r2(v), "原文"
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "中性", "None"):
        return 7, "默认7"
    m = re.fullmatch(r"([0-9.]+)\s*[-~–]\s*([0-9.]+)", s)
    if m:
        return r2((float(m.group(1)) + float(m.group(2))) / 2), "原文(区间中值)"
    m = NUM_RE.match(s)
    if m:
        return r2(m.group(1)), "原文"
    return None, "未识别:" + s


def _mins(val, unit):
    if unit.startswith("h"):
        return val * 60
    if unit.startswith("s"):
        return val / 60
    return val


def parse_stab(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "None"):
        return 30, "默认30"
    if "预浸" in s and "OCP" in s:
        tot = sum(_mins(float(a), b) for a, b in re.findall(r"([0-9.]+)\s*(h|min|s)", s))
        return r2(tot), "原文(预浸+OCP相加)"
    if s.startswith("阴极极化"):
        ms = re.findall(r"([0-9.]+)\s*(h|min|s)", s)
        if ms:
            return r2(_mins(float(ms[-1][0]), ms[-1][1])), "原文(阴极极化段)"
    ms = re.findall(r"([0-9.]+)\s*(h|min|s)", s)
    if ms:
        a, b = ms[-1]
        return r2(_mins(float(a), b)), "原文"
    return None, "未识别:" + s


def parse_ce(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "None"):
        return "铂", "默认铂"
    if "石墨" in s:
        return "石墨", "原文"
    if "铂" in s or s.upper().startswith("PT"):
        return "铂", "原文"
    return None, "未识别:" + s


MODE = sys.argv[1] if len(sys.argv) > 1 else "dry"
assert MODE in ("dry", "write")

src = openpyxl.load_workbook(SRC)
ws = src["Sheet1"]

rows, excluded = [], []
flags = []
stat = {
    "元素0填充": 0, "元素取上限": 0, "元素取中值": 0, "元素ppm": 0,
}
epit_cnt, fe_cnt = {}, {"实测": 0, "反算": 0}
default_cnt = {}

cur = None
for r in range(3, ws.max_row + 1):
    raw_id = ws.cell(r, C_ID).value
    if raw_id not in (None, ""):
        cur = raw_id
    name = ws.cell(r, C_NAME).value
    if name in (None, ""):
        continue
    pid = int(float(cur))
    epit_raw = ws.cell(r, C_EPIT).value

    if pid == DROP_PAPER:
        excluded.append([pid, str(name), r, "文献20 原文无成分",
                         "Cr/Ni 记「调整配比」、其余元素「未提及」，无法折算"])
        continue
    if epit_raw is None or (isinstance(epit_raw, str) and epit_raw.strip().upper() == "N/A"):
        excluded.append([pid, str(name), r, "无 Epit 数值",
                         "原文未测出/未给单值，见 0.xlsx 备注"])
        continue

    rec = {C_ID: pid, C_NAME: str(name)}
    zero_cols, unk = [], []
    for c in C_ELEM:
        val, tag = parse_elem(ws.cell(r, c).value)
        rec[c] = r2(val, 6) if tag in ("cap", "mid", "ppm") else val
        if tag.startswith("0"):
            stat["元素0填充"] += 1
            zero_cols.append(ws.cell(2, c).value)
        elif tag == "cap":
            stat["元素取上限"] += 1
        elif tag == "mid":
            stat["元素取中值"] += 1
        elif tag == "ppm":
            stat["元素ppm"] += 1
        elif tag.startswith("unknown"):
            unk.append(tag)

    fe_raw = ws.cell(r, C_FE).value
    if isinstance(fe_raw, (int, float)) and not isinstance(fe_raw, bool):
        rec[C_FE] = r2(fe_raw, 2)
        fe_src = "实测"
    else:
        rec[C_FE] = r2(100 - sum(rec[c] for c in C_ELEM), 2)
        fe_src = "反算"
    fe_cnt[fe_src] += 1

    re_raw = str(ws.cell(r, C_RE).value or "").strip()
    if re_raw not in RE_MAP:
        raise SystemExit("参比未覆盖：row %d %r" % (r, re_raw))
    delta, re_tag = RE_MAP[re_raw]
    rec[C_EPIT] = r2(float(epit_raw) + delta, 4)
    rec[C_RE] = "SCE"
    epit_cnt[re_tag] = epit_cnt.get(re_tag, 0) + 1

    temp, temp_src = parse_temp(ws.cell(r, C_TEMP).value)
    stab, stab_src = parse_stab(ws.cell(r, C_STAB).value)
    ce, ce_src = parse_ce(ws.cell(r, C_CE).value)
    if None in (temp, stab, ce):
        raise SystemExit("解析失败 row %d: %r %r %r" % (r, temp_src, stab_src, ce_src))
    rec[C_TEMP], rec[C_STAB], rec[C_CE] = temp, stab, ce
    for tag in (temp_src, stab_src, ce_src):
        if tag.startswith("默认"):
            default_cnt[tag] = default_cnt.get(tag, 0) + 1
    rec[C_PROC] = ws.cell(r, C_PROC).value
    rec[C_WE] = ws.cell(r, C_WE).value

    rows.append(rec)
    tot = sum(rec[c] for c in C_ELEM) + rec[C_FE]
    notes = list(unk)
    if fe_src == "反算" and any(x in zero_cols for x in ("C", "Si", "Mn", "Cr", "Ni")):
        notes.append("Fe反算受未报主元素影响")
    if abs(tot - 100) > 1:
        notes.append("成分和=%.2f" % tot)
        stat["成分和偏离>1"] = stat.get("成分和偏离>1", 0) + 1
    flags.append([pid, str(name), r,
                  ",".join(zero_cols) if zero_cols else "—", len(zero_cols),
                  fe_src, re_tag, delta, temp_src, stab_src, ce_src,
                  "；".join(notes) if notes else ""])

print("保留 %d 条 / 未纳入 %d 条" % (len(rows), len(excluded)))
print("元素折算：", stat, " Fe：", fe_cnt)
print("Epit 换算：", epit_cnt)
print("默认值：", default_cnt)
print("文献数：", len({r[C_ID] for r in rows}))
print("未纳入：")
for e in excluded:
    print("   行%-4d 文献%-3d %-28s %s" % (e[2], e[0], e[1][:26], e[3]))
bad = [f for f in flags if "unknown:" in str(f[11])]
if bad:
    print("仍有未识别标签：", bad)

if MODE != "write":
    out = PREVIEW
else:
    if os.path.exists(DST) and not os.path.exists(BAK):
        os.makedirs(os.path.dirname(BAK), exist_ok=True)
        shutil.copy2(DST, BAK)
        print("旧版备份 ->", BAK)
    out = DST

# ---------- 生成新工作簿 ----------
wb = openpyxl.Workbook()
sh = wb.active
sh.title = "Sheet1"

WIDTHS = {"A": 9.0, "B": 20.1, "C": 7.17, "D": 7.17, "E": 7.17, "F": 7.17, "G": 7.17,
          "H": 7.17, "I": 7.17, "J": 7.17, "K": 7.17, "L": 7.17, "M": 7.17, "N": 7.17,
          "O": 7.17, "P": 7.17, "Q": 8.58, "R": 72.33, "S": 22.92, "T": 15.33,
          "U": 11.4, "V": 8.75, "W": 7.9, "X": 9.17, "Y": 14.33}

HDR1 = {1: "文献编号", 2: "样本（原文命名）", 3: "化学成分(wt.%)", 18: "加工工艺",
        19: "工作电极(WE)", 20: "参比电极(RE)", 21: "对电极(CE)",
        22: "腐蚀温度(℃)", 23: "Epit(V)", 24: "测试前稳定时间(min)"}
HDR2 = {3: "C", 4: "Si", 5: "N", 6: "P", 7: "S", 8: "Cu", 9: "Mn", 10: "Cr",
        11: "Mo", 12: "Ni", 13: "Ti", 14: "Nb", 15: "Co", 16: "Sn", 17: "Fe"}

def copy_style(dst, src):
    """按属性复制样式（直接搬 _style 索引会跨工作簿失效）。"""
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


for c in range(1, 25):
    sh.cell(1, c).value = HDR1.get(c)
    copy_style(sh.cell(1, c), ws.cell(1, c))
    sh.cell(2, c).value = HDR2.get(c)
    copy_style(sh.cell(2, c), ws.cell(2, c))

for i, rec in enumerate(rows, start=3):
    for c in range(1, 25):
        cell = sh.cell(i, c)
        cell.value = rec.get(c)
        copy_style(cell, ws.cell(3, c))

for m in ("A1:A2", "B1:B2", "C1:Q1", "R1:R2", "S1:S2", "T1:T2", "U1:U2",
          "V1:V2", "W1:W2", "X1:X2"):
    sh.merge_cells(m)
for col, w in WIDTHS.items():
    sh.column_dimensions[col].width = w
sh.sheet_view.zoomScale = 150
sh.freeze_panes = "C3"

# ---- 统计 ----
st = wb.create_sheet("统计")
st.append(["项目", "数值"])
st.append(["总条数", len(rows)])
st.append(["覆盖文献数", len({r[C_ID] for r in rows})])
st.append(["未纳入：文献20（原文无成分）", 5])
st.append(["未纳入：无 Epit 数值", len(excluded) - 5])
st.append(["—", "—"])
st.append(["元素折算", "格数"])
st.append(["未提及/空 → 0", stat["元素0填充"]])
st.append(["<x / ≤x → 取上限", stat["元素取上限"]])
st.append(["a-b 区间 → 取中值", stat["元素取中值"]])
st.append(["ppm → wt.%（÷1e4）", stat["元素ppm"]])
st.append(["—", "—"])
st.append(["Fe", "条数"])
st.append(["原文实测", fe_cnt["实测"]])
st.append(["100 − Σ14元素 反算", fe_cnt["反算"]])
st.append(["成分和偏离 100 超过 1", stat.get("成分和偏离>1", 0)])
st.append(["—", "—"])
st.append(["Epit 参比换算", "条数"])
for k, v in sorted(epit_cnt.items(), key=lambda kv: -kv[1]):
    st.append(["  " + k, v])
st.append(["  合计（统一记 SCE）", sum(epit_cnt.values())])
st.append(["—", "—"])
st.append(["缺失值默认口径", "条数"])
for k, v in sorted(default_cnt.items(), key=lambda kv: -kv[1]):
    st.append(["  " + k, v])
st.append(["—", "—"])
from collections import Counter
cnt = Counter(r[C_ID] for r in rows)
st.append(["文献编号", "条数"])
for k in sorted(cnt):
    st.append([k, cnt[k]])
st.append(["合计", sum(cnt.values())])

# ---- 清洗标记 ----
fm = wb.create_sheet("清洗标记")
fm.append(["文献编号", "样本（原文命名）", "0.xlsx行号", "0填充元素", "0填充格数",
           "Fe来源", "Epit原参比", "Epit换算Δ(V)", "温度", "稳定时间", "对电极", "异常"])
for f in flags:
    fm.append(f)
for col, w in {"A": 9, "B": 26, "C": 10, "D": 30, "E": 10, "F": 9, "G": 26,
               "H": 12, "I": 14, "J": 16, "K": 10, "L": 20}.items():
    fm.column_dimensions[col].width = w

# ---- 未纳入 ----
fx = wb.create_sheet("未纳入")
fx.append(["文献编号", "样本（原文命名）", "0.xlsx行号", "类别", "说明"])
for e in excluded:
    fx.append(e)
for col, w in {"A": 9, "B": 30, "C": 10, "D": 24, "E": 60}.items():
    fx.column_dimensions[col].width = w

wb.save(out)
print("\n[%s] -> %s  sheets=%s" % (MODE, out, wb.sheetnames))
