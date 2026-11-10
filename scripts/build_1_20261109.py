# -*- coding: utf-8 -*-
r"""从当前 0.xlsx（641 行 / 639 条 / 26 列）重建 1.xlsx（清洗版）。

用户 m03497 / m03575 / m03576 批准的口径：
  1) 数据清洗照旧 build_1_20261105.py：删文献20、删 Epit=N/A、元素归一、
     Fe 反算、Epit 统一 SCE、温度/稳定时间/对电极归一、文献编号逐行填值。
  2) 规范表述：工艺列对照师兄最初图片示例（成形方式 + 热处理制度 + 冷却方式）
     做无损记号归一并逐条审计；备注列精简为「数据来源 + 特殊情况」。
  3) 1.xlsx **保留 Y「文献信息」/ Z「备注」两列**（逐行填值，不做竖合并）。
  4) 饱和 Ag/AgCl 沿用旧脚本 −0.042 V（与口径基准 −0.044 V 差 0.002 V，报告注明）。

用法：
  D:\Python\python.exe -X utf8 scripts\build_1_20261109.py dry     # 生成校验件 + 五份审计表 + 四段报告（不写 1.xlsx）
  D:\Python\python.exe -X utf8 scripts\build_1_20261109.py commit  # 校验后写回 1.xlsx 并二次复验
"""
import csv
import hashlib
import json
import os
import re
import shutil
import sys
import unicodedata
from collections import Counter
from copy import copy

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "0.xlsx")
DST = os.path.join(ROOT, "1.xlsx")
WORK = os.path.join(ROOT, "备份", "清洗1_20261109")
STAGE = os.path.join(WORK, "1_校验.xlsx")
BASELINE = os.path.join(ROOT, "备份", "1_基准_清洗前_20261109.xlsx")
EXPECT_0_MD5 = "d4d601cd48d637b0b5b63309abb82c5e"

C_ID, C_NAME = 1, 2
C_ELEM = list(range(3, 17))          # C..Sn (14 列)
C_FE = 17                            # Fe
C_PROC, C_WE, C_RE, C_CE = 18, 19, 20, 21
C_TEMP, C_EPIT, C_STAB = 22, 23, 24
C_Y, C_Z = 25, 26
NCOL = 26
DROP_PAPER = 20

DELTA_3M = -0.031
DELTA_SAT = -0.042

# ---------------------------------------------------------------- 通用工具

def nfkc(v):
    if v is None:
        return ""
    return unicodedata.normalize("NFKC", str(v)).strip()


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def r2(x, n=2):
    v = round(float(x), n)
    return int(v) if v == int(v) else v


def unify_minus(s):
    for ch in ("−", "‑", "‒", "–", "—", "－", "﹣", "－"):
        s = s.replace(ch, "-")
    return s


def csv_write(path, header, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


# ---------------------------------------------------------------- 参比映射（30 种全覆）

RE_SCE = {
    "SCE", "SCE(饱和KCl甘汞)", "SCE(饱和甘汞)", "SSCE", "VSCE",
    "a saturated calomel electrode (SCE)", "saturated calomel electrode (SCE)",
    "saturated calomel electrode (saturated KCl)", "saturated calomel electrode",
    "a saturated mercury electrode (saturated potassium chloride)",
    "饱和甘汞电极", "飽和甘コウ電極(SCE)",
    "Ag/AgCl(饱和KCl)→SCE",           # 原文已换算到 SCE
}
RE_3M = {
    "Ag/AgCl", "Ag/AgCl(3M KCl)", "Ag/AgCl (3 M KCl)",
    "Ag/AgCl(3.3kmolm―3KCI)", "Ag/AgCl(内部液3.3 kmol·m−3 KCl)",
    "silver chloride(Ag/AgCl)", "은-염화은(Ag/AgCl)",
}
RE_SAT = {
    "Ag/AgCl(sat.KCl)", "Ag/AgCl(饱和KCl)", "Ag/AgCl(饱和 KCl)",
    "Ag/AgCl (saturated KCl)", "Ag/AgCl(saturated KCl)",
    "Ag/AgCl(饱和 KCl, 0.1989 V vs SHE)", "saturated Ag/AgCl",
    "Ag/AgCl in saturated KCl electrode", "Ag/AgCl (sat, 3M)",
}


RE_SCE_N = {nfkc(x) for x in RE_SCE}
RE_3M_N = {nfkc(x) for x in RE_3M}
RE_SAT_N = {nfkc(x) for x in RE_SAT}


def re_delta(v):
    """返回 (ΔV, 类别, 是否需排除)。未提及/无法识别 -> 排除。"""
    k = nfkc(v)
    if k in RE_SCE_N:
        return 0.0, "SCE", False
    if k in RE_3M_N:
        return DELTA_3M, "Ag/AgCl(3M/未注明)", False
    if k in RE_SAT_N:
        return DELTA_SAT, "Ag/AgCl(饱和KCl)", False
    up = k.upper()
    if "SCE" in up or "甘汞" in k or "CALOMEL" in up or "MERCURY" in up or "코우" in k:
        return 0.0, "SCE(推断)", False
    if "AG/AGCL" in up or "염화은" in k or "SILVER CHLORIDE" in up:
        low = k.lower()
        # 3.3 kmol/m³ KCl 是约 3 M，不是饱和
        if "3.3kmol" in low or "内部液" in k or "3.3 kmol" in low:
            return DELTA_3M, "Ag/AgCl(3.3kmol→3M,推断)", False
        if any(t in low for t in ("sat", "saturated", "饱和", "0.1989")):
            return DELTA_SAT, "Ag/AgCl(饱和,推断)", False
        return DELTA_3M, "Ag/AgCl(3M,推断)", False
    if k in ("", "未提及", "NONE", "-", "—"):
        return None, "参比未提及", True
    return None, "参比未识别:" + k, True


# ---------------------------------------------------------------- 元素解析

NUM_RE = re.compile(r"^\s*[-+]?\d+(?:\.\d+)?\s*$")
CAP_RE = re.compile(r"^\s*(?:<|≤|≦|Max|max|MAX|＜|＜=)\s*([0-9]+(?:\.[0-9]+)?)\s*$")
RANGE_RE = re.compile(
    r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:-|~|–|—|to|TO)\s*([0-9]+(?:\.[0-9]+)?)\s*$"
)
PPM_RE = re.compile(r"^\s*(?:<\s*)?([0-9]+(?:\.[0-9]+)?)\s*ppm\s*$", re.I)
PCT_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:%|pct|wt\.?%|wt%)\s*$", re.I)
X10_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*[×xX]\s*10\s*-\s*(\d+)\s*%?\s*$")
SCI_RE = re.compile(r"^\s*([-+]?\d+(?:\.\d+)?)[eE]([-+]?\d+)\s*$")
PM_RE = re.compile(r"^\s*([-+]?\d+(?:\.\d+)?)\s*±\s*\d+(?:\.\d+)?\s*$")

ZERO_TOKENS = ("", "未提及", "未给出", "未注明", "None", "-", "--", "—", "–", "…",
               "...", "微量", "加N", "调整配比", "–", "－", "N/A", "NA")


def parse_elem(v):
    """返回 (数值, 标签)。标签：num / 0(未提及) / cap / mid / ppm / pct / x10 / pm / 0(非数值)"""
    if v is None:
        return 0.0, "0(未提及)"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v), "num"
    s = nfkc(v)
    if s in ZERO_TOKENS or s.replace(" ", "") in ZERO_TOKENS:
        return 0.0, "0(未提及)"
    m = X10_RE.match(unify_minus(s))
    if m:
        return round(float(m.group(1)) * 10 ** (-int(m.group(2))), 6), "x10"
    m = PPM_RE.match(s)
    if m:
        return round(float(m.group(1)) / 1e4, 6), "ppm"
    m = SCI_RE.match(unify_minus(s))
    if m:
        return float(m.group(1)) * 10 ** int(m.group(2)), "num"
    m = PM_RE.match(unify_minus(s))
    if m:
        return float(m.group(1)), "pm"
    m = CAP_RE.match(s)
    if m:
        return float(m.group(1)), "cap"
    m = RANGE_RE.match(unify_minus(s))
    if m:
        return round((float(m.group(1)) + float(m.group(2))) / 2, 6), "mid"
    m = PCT_RE.match(s)
    if m:
        return float(m.group(1)), "pct"
    m = NUM_RE.match(unify_minus(s))
    if m:
        return float(m.group(1)), "num"
    return 0.0, "0(非数值:" + s[:20] + ")"


# ---------------------------------------------------------------- Epit 解析

NA_RE = re.compile(r"^(N/?A|NA|NAN|NONE|NONE|-|--|—|–|)$", re.I)


def parse_epit(v):
    """返回 (V 数值, 单位推断, 中心原文) 或 None（N/A/无法解析）。"""
    if v is None:
        return None
    s = nfkc(v)
    if s.upper().startswith("N/A") or "待读图" in s or NA_RE.match(s):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        x = float(v)
        if abs(x) > 2:
            return x / 1000.0, "mV(数值>2)", "mV"
        return x, "V", "V"
    s = unify_minus(s)
    unit_mv = bool(re.search(r"m\s*V|mVSCE|mV_SCE|mVSCE", s, re.I))
    unit_v = bool(re.search(r"(?<![A-Za-z])V(?![A-Za-z])|VSCE|V_SCE|Vvs|V vs|VSHE", s, re.I))
    # 去括号内容、单位词、约数词
    core = re.sub(r"\([^)]*\)", " ", s)
    core = re.sub(r"\b(about|approx|approximately|ca\.|around)\b", " ", core, flags=re.I)
    core = re.sub(r"m\s*V|mVSCE|mV_SCE|VSCE|V_SCE|VSHE|Vvs|V vs|\bV\b|SCE|Ag/AgCl",
                  " ", core, flags=re.I)
    core = core.replace("≈", " ").replace("~", " ").replace("∼", " ")
    core = re.sub(r"<|≤|≦|Max|max", " ", core)
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", unify_minus(core))
    if not nums:
        return None
    f = [float(x) for x in nums]
    if "±" in s or "+/-" in s:
        center = f[0]
        tag = "±中心"
    elif re.search(r"\bto\b", s, re.I) or re.search(r"\d\s*~\s*\d", s):
        center = (f[0] + f[1]) / 2 if len(f) >= 2 else f[0]
        tag = "区间中值"
    else:
        center = f[0]
        tag = "单值"
    if unit_mv:
        return center / 1000.0, "mV→V", tag
    if abs(center) > 2 and not unit_v:
        return center / 1000.0, "无单位|v|>2→mV", tag
    return center, "V", tag


# ---------------------------------------------------------------- 温度 / 稳定时间 / 对电极

WORD_NUM = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
            "six": 6, "seven": 7, "eight": 8, "ten": 10, "fifteen": 15,
            "twenty": 20, "thirty": 30, "sixty": 60, "half": 0.5}


def parse_temp(v):
    s = nfkc(v)
    if s == "" or s in ("未提及", "None", "-", "—"):
        return 25.0, "默认25"
    s = unify_minus(s)
    s = s.replace("T1s=", "").replace("T1s", "")
    s = (s.replace("° C", "℃").replace("°C", "℃").replace("◦C", "℃")
          .replace("ºC", "℃").replace("oC", "℃").replace("°Ｃ", "℃")
          .replace("°", "℃"))
    low = s.lower()
    has_k = ("K" in s.upper()) and ("℃" not in s)
    if not re.search(r"\d", s):
        if any(t in low for t in ("室温", "room", "ambient", "상온", "约", "未提及")):
            return 25.0, "默认25(室温)"
        return 25.0, "默认25(无数值)"
    nums = [float(x) for x in re.findall(r"[-+]?\d+(?:\.\d+)?", unify_minus(s))]
    if not nums:
        return 25.0, "默认25(无数值)"
    if "±" in s:
        center, tag = nums[0], "±中心"
    elif re.search(r"\bto\b", s, re.I) or re.search(r"\d\s*[-~]\s*\d", s):
        center = (nums[0] + nums[1]) / 2 if len(nums) >= 2 else nums[0]
        tag = "区间中值"
    elif len(nums) > 2 and ("," in s or " and " in low):
        center, tag = sum(nums) / len(nums), "多温度均值"
    else:
        center, tag = nums[0], "原文"
    if has_k:
        # 旧脚本口径：K → ℃ 取整数
        return float(int(round(center - 273.15))), tag + "(K→℃)"
    return r2(center, 1), tag


def parse_stab(v):
    s = nfkc(v)
    if s == "" or s in ("未提及", "None", "-", "—") or s.startswith("未提及"):
        return 30.0, "默认30"
    s = unify_minus(s)
    # 千分位空格：1 800 s -> 1800 s；不合并小数
    s = re.sub(r"(?<=\d)\s+(?=\d{3}\b)", "", s)
    # 原文明确说明 OCP 稳定时间未提及 -> 默认 30
    if "OCP稳定时间未提及" in s or "稳定时间未提及" in s:
        return 30.0, "默认30(原文说明未提及)"
    # 括注内容不参与取数（保留原句判断 OCP 关键字）
    s_num = s
    for _ in range(3):
        s2 = re.sub(r"\([^)]*\)", " ", s_num)
        if s2 == s_num:
            break
        s_num = s2
    # 英文词数字：two-hour / 15-minute
    words = []
    for m in re.finditer(r"\b([a-z]+)\s*-\s*(hour|minute|hr|min)s?\b", s_num, re.I):
        n = WORD_NUM.get(m.group(1).lower())
        if n:
            words.append(n * 60 if m.group(2).lower().startswith(("h",)) else n)
    pairs = [(float(a), b.lower()) for a, b in
             re.findall(r"([0-9]+(?:\.[0-9]+)?)\s*(hours|hour|hr|h|minutes|minute|mins|min|seconds|second|sec|s|分|小时)",
                        s_num, re.I)]
    def to_min(val, unit):
        if unit.startswith("h") or unit in ("小时",):
            return val * 60
        if unit.startswith("s"):
            return val / 60
        return val
    if "预浸" in s and "OCP" in s and len(pairs) >= 2:
        return r2(sum(to_min(a, b) for a, b in pairs), 2), "预浸+OCP相加"
    # 含 OCP/开路电位/浸泡/soaked 时取最后一段稳定段，否则取第一个时间（跳过阴极极化）
    ocp = bool(re.search(
        r"OCP|开路|open circuit|soak|immersion|浸泡|静置|稳定|monitoring|delay|"
        r"then|kept|held|record|自腐蚀|自然电位|自然电极|腐蚀电位|记录", s, re.I))
    # 英文词数字 + 数字段（如 two-hour OCP; additional 15-minute）合并求和
    if words and ocp:
        total = sum(words) + sum(to_min(a, b) for a, b in pairs)
        return r2(total, 2), "英文词数字+OCP"
    if pairs:
        a, b = (pairs[-1] if ocp else pairs[0])
        return r2(to_min(a, b), 2), "原文(%s)" % ("OCP末段" if ocp else "首段")
    if words:
        return r2(sum(words), 2), "英文词数字"
    jp = re.search(r"(\d+)\s*分", s)
    if jp:
        return r2(float(jp.group(1)), 2), "日文分→min"
    return 30.0, "默认30(无法解析)"


def parse_ce(v):
    s = nfkc(v)
    if s == "" or s in ("未提及", "None", "-", "—"):
        return "铂", "默认铂"
    if "石墨" in s or "graphite" in s.lower() or "carbon" in s.lower():
        return "石墨", "原文(石墨族)"
    if ("铂" in s or "白金" in s or "백금" in s or s.upper().startswith("PT")
            or "platinum" in s.lower()):
        return "铂", "原文(铂族)"
    return "铂", "默认铂(未识别)"


# ---------------------------------------------------------------- 工艺表述规范 + 审计

PROC_CTRL = "".join(chr(i) for i in range(0, 32))


def norm_process(v):
    s = nfkc(v)
    s = "".join(ch for ch in s if ch not in PROC_CTRL)
    s = unify_minus(s)
    s = (s.replace("° C", "℃").replace("°C", "℃").replace("◦C", "℃")
          .replace("ºC", "℃").replace("oC", "℃").replace("°Ｃ", "℃"))
    s = s.replace("，", "，").replace(",", "，")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


FORM_KEYS = [
    ("砂型铸造", "砂型铸造"), ("离心铸造", "离心铸造"), ("熔模铸造", "熔模铸造"),
    ("真空精密铸造", "真空精密铸造"), ("精密铸造", "精密铸造"),
    ("连铸", "连铸"), ("铸锭", "铸造"), ("铸态", "铸态"), ("铸造", "铸造"),
    ("真空感应", "熔炼"), ("中频炉", "熔炼"), ("电弧炉", "熔炼"), ("熔炼", "熔炼"),
    ("热锻", "锻造"), ("锻造", "锻造"), ("锻轧", "锻造+轧制"),
    ("热轧", "热轧"), ("温轧", "温轧"), ("冷轧", "冷轧"), ("深冷", "深冷轧制"),
    ("轧制", "轧制"), ("轧板", "轧制"), ("薄板", "轧制"), ("棒材", "棒材"),
]
HEAT_KEYS = ["固溶", "退火", "时效", "敏化", "稳定化", "均匀化", "去应力", "淬火",
             "solution", "anneal", "aging", "sensitiz"]
COOL_KEYS = ["水淬", "水冷", "空冷", "炉冷", "油淬", "随炉冷却", "淬火",
             "water quench", "air cool", "furnace cool"]


def canon_process(s):
    """只做无损记号归一：NFKC、控制符、减号、℃、逗号、连续空格。
    工艺原文的成形/热处理/冷却措辞不做改写，避免把 336 种复杂流程压成错误短句。
    符合师兄示例与否由 proc_audit 判定并写审计表。"""
    return norm_process(s)


def proc_audit(s):
    n = norm_process(s)
    form = ""
    for k, canon in FORM_KEYS:
        if k.lower() in n.lower():
            form = canon
            break
    heat = [k for k in HEAT_KEYS if k.lower() in n.lower()]
    temps = re.findall(r"(\d{3,4})\s*℃|(\d{3,4})\s*K", n)
    temp = ""
    for t in temps:
        temp = t[0] or (t[1] + "K")
        break
    times = re.findall(r"(\d+(?:\.\d+)?)\s*(min|h|s|小时|分钟)\b", n, re.I)
    tstr = times[0][0] + times[0][1] if times else ""
    cool = [k for k in COOL_KEYS if k.lower() in n.lower()]
    missing = ("未提及" in n or "未写" in n or "未给" in n or "未注明" in n
               or "原文未" in n or n == "")
    complex_chain = n.count("→") >= 3 or len(n) > 160
    has_core = bool(temp) and bool(tstr) and any(k in n for k in ("固溶", "退火", "solution", "anneal"))
    if missing:
        cat, ok = "缺参数(保留原文)", False
    elif complex_chain:
        cat, ok = "复杂流程(保留原文)", False
    elif has_core and form:
        cat, ok = "符合示例", True
    elif has_core:
        cat, ok = "可规范化(缺成形方式)", False
    else:
        cat, ok = "需人工核对", False
    return {
        "norm": n, "cat": cat, "ok": ok, "form": form,
        "heat": "/".join(dict.fromkeys(heat)), "temp": temp, "time": tstr,
        "cool": "/".join(dict.fromkeys(cool)),
    }


# ---------------------------------------------------------------- 37 块长备注精简（逐块人工，保留来源与特殊情况）

SHORT_NOTES = {
    12: "Fig.4读图取平均；天然海水；复核修订：SCS14-B 改记 0.96 V。",
    13: "A/D/H取Table 3，其余Fig.4读图；天然海水；复核修订：G 改记 −0.02 V。",
    28: "成分取p2 Table 1；Epit由Fig.3(a)读图（±0.02 V）；同文去重；复核修订：Mo1N2 改记 0.87 V。",
    29: "成分取p3 Table 1（无Al列）；Epit取p7正文；Nb-free的Nb原文为「-」记未提及。",
    30: "成分取p3表1、Eb取p5表2（图片放大读数）；只录3.5% NaCl；1#工艺与2#–4#不同。",
    31: "成分取p2 Table 1；Epit取p8 Table 3（100 μA·cm⁻²判据）；Steel A/B由S含量对应。",
    32: "点蚀电位为统计分布，Fig.7只给中位区间；无逐试样Epit，本行记N/A；成分表无N列。",
    33: "只录0.55 V（稳定点蚀最小电位，800℃/300s）；预浸16 h；腐蚀温度原文未提及。",
    34: "Epit取Table 5/6并去重；固溶后另有900℃×0.5h时效（两步热处理）；A29为冷轧态。",
    35: "只取750℃退火两条；Epit取p13 Table 6（mV→V）；as-received/冷轧未退火不录。",
    36: "只取3.5% NaCl SME介质CR90退火态；Epit −0.116 V（p10正文/p11 Table 3）；SBF不录。",
    37: "只取0%温预应变两条；Epit取表2/表3中心值（mV→V）；3600 s为预浸时间。",
    38: "Epit取p7 Table 4（3.56 wt% NaCl）；Table 1无Si/P列；Mo值按列对齐；3.56%口径存疑。",
    40: "9条纯铸态无热处理；成分取p4 Table 1（Al/O/Mg列未录）；Epit取p13 Table 3中心值；pH 6.4。",
    41: "成分取p2 Table 1（文字层可读）；Epit取p7 Table 4（mV→V）；As received两组工艺存疑。",
    42: "CG未击穿不录；成分取p2正文；Epit取p6 Table 1（mV→V）；退火时间未一一对应。",
    43: "成分取p2 Table 1（O列未录）；Epit读Fig.6（±0.02 V）；Nb00@1300℃按正文推断同1100℃值。",
    44: "成分取p2 Table 1（上限值照录）；Epit取25℃表值1.209 V；>40℃原文改用Ag/AgCl。",
    45: "成分取p2实验段正文；Epit取p4 Table 1；Fe原文未单列记未提及。",
    46: "成分取p3 Table 1（裁剪放大核对，空格记未提及）；Epit取p13 Table 6；TT1/TT2固溶温度见原文。",
    47: "原文件名年份标2016、正文实为2020，入库前已改名；成分取p2正文；Epit取p4 Table 1；Fe未单列。",
    48: "成分取p3表1（仅C/Si/Mn/Cr/Ni）；Epit取p5表2；1-弯管未处理Epit −0.020 V为真实值；Fe未单列。",
    49: "成分为p2正文标称值（非实测）；工艺含1200℃固溶+650℃×5000h时效；Epit读Fig.4b。",
    50: "成分取p2表1；Epit取p3表2/p5结论；对比材料Mo <0.001照录；本批唯一真铸造牌号。",
    51: "成分取p2正文（W 0.03无对应列未录）；Epit取正文320→425 mV；Ti/Sn未提及。",
    52: "成分取p2 Table 1（Ni 0.013/0.012为痕量非笔误；O列未录）；Epit读Fig.1（14条）。",
    53: "成分取p2 Table 1（两钢仅碳含量不同）；Epit取p6 Table 3；稳定含阴极极化5 min+静置1 h。",
    54: "成分取p2 Table 1（Al列未录；S值单位应为ppm）；只录pH 6组4条；Fe按余量。",
    55: "成分取p2表（裁剪放大核对）；Epit原文已按−0.042 V换算为SCE；含冷轧与950–1150℃固溶。",
    56: "成分取p2 Table 1；Epit取p8正文8–10次重复均值；AISI 304对比值未录；稳定含阴极极化+OCP。",
    57: "成分取p2成分段；Epit取p5 Table 1均值；OCP 3600 s；温度原文未提及。",
    58: "成分取p2表1；Epit取p2表2 Eb；工艺3 mm冷轧板1050℃×20 min固溶水冷；OCP未提及。",
    59: "成分取p2表1；Epit取p3表2 3.5%NaCl列Eb10均值；一回路液数据不录；同团队重复数据已去重。",
    60: "成分取p3/p4 Table 1；Epit取p6 Table 3 NaCl列321（二手304/316L数据未录）；温度未提及。",
    61: "成分取p2 Table 1；Epit取p5 Table 2；Y含量写入样本名；铸态无热处理；<−0.35为扫描下限。",
    62: "成分取p2实验段；Epit取p5表1；UFG/FG为冷轧+900℃×3min退火；650℃敏化与Epit无关。",
    63: "成分取p2表1；Epit取p2表2/摘要（mV→V）；439/443铁素体未录；OCP未提及。",
}


def trim_note(pid, z):
    if pid in SHORT_NOTES:
        return SHORT_NOTES[pid], True
    return z, False


# ---------------------------------------------------------------- 主流程

def build_records(ws):
    """返回 rows(保留), excluded(未纳入), flags(清洗标记), notes_cmp(备注对照)。"""
    rows, excluded, flags, notes_cmp = [], [], [], []
    cur_id, cur_y, cur_z = None, "", ""
    seen_note_pid = set()
    stat = Counter()
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, C_ID).value
        if a not in (None, ""):
            try:
                cur_id = int(float(a))
            except (TypeError, ValueError):
                cur_id = a
        y = ws.cell(r, C_Y).value
        z = ws.cell(r, C_Z).value
        if y not in (None, ""):
            cur_y = str(y)
        if z not in (None, ""):
            cur_z = str(z)
        name = ws.cell(r, C_NAME).value
        if name in (None, ""):
            continue
        pid = int(cur_id)
        # 块级备注对照（每块一次）
        if pid not in seen_note_pid:
            seen_note_pid.add(pid)
            new_z, changed = trim_note(pid, cur_z)
            notes_cmp.append((pid, len(cur_z), len(new_z), "是" if changed else "否",
                              cur_z, new_z))
        else:
            new_z, changed = trim_note(pid, cur_z)

        # ---- 剔除判定 ----
        epit_raw = ws.cell(r, C_EPIT).value
        if pid == DROP_PAPER:
            excluded.append([pid, str(name), r, "文献20 原文无成分",
                             "Cr/Ni 记「调整配比」、其余元素「未提及」，无法折算"])
            continue
        ep = parse_epit(epit_raw)
        if ep is None:
            excluded.append([pid, str(name), r, "无 Epit 数值",
                             "0.xlsx 记 %s" % nfkc(epit_raw)[:40]])
            stat["剔除_无Epit"] += 1
            continue
        delta, re_tag, re_drop = re_delta(ws.cell(r, C_RE).value)
        if re_drop:
            excluded.append([pid, str(name), r, "参比无法识别",
                             "RE=%s；Epit 不能统一到 SCE，未猜值" % nfkc(ws.cell(r, C_RE).value)[:40]])
            stat["剔除_参比"] += 1
            continue

        rec = {C_ID: pid, C_NAME: str(name)}
        zero_cols, unk = [], []
        for c in C_ELEM:
            val, tag = parse_elem(ws.cell(r, c).value)
            rec[c] = r2(val, 6)
            if tag.startswith("0(未提及)"):
                stat["元素0填充"] += 1
                zero_cols.append(str(ws.cell(2, c).value))
            elif tag == "cap":
                stat["元素取上限"] += 1
            elif tag == "mid":
                stat["元素取中值"] += 1
            elif tag == "ppm":
                stat["元素ppm"] += 1
            elif tag in ("pm", "pct", "x10"):
                stat["元素" + tag] += 1
            elif tag.startswith("0(非数值"):
                unk.append(tag)
                stat["元素非数值→0"] += 1

        fe_raw = ws.cell(r, C_FE).value
        if isinstance(fe_raw, (int, float)) and not isinstance(fe_raw, bool):
            rec[C_FE] = r2(fe_raw, 2)
            fe_src = "实测"
        else:
            rec[C_FE] = r2(100 - sum(rec[c] for c in C_ELEM), 2)
            fe_src = "反算"
        stat["Fe_" + fe_src] += 1

        ep_v, ep_unit, ep_tag = ep
        rec[C_EPIT] = r2(ep_v + delta, 4)
        rec[C_RE] = "SCE"
        stat["Epit_" + re_tag] += 1
        stat["Epit单位_" + ep_unit] += 1

        temp, temp_src = parse_temp(ws.cell(r, C_TEMP).value)
        stab, stab_src = parse_stab(ws.cell(r, C_STAB).value)
        ce, ce_src = parse_ce(ws.cell(r, C_CE).value)
        rec[C_TEMP], rec[C_STAB], rec[C_CE] = temp, stab, ce
        for tag in (temp_src, stab_src, ce_src):
            if tag.startswith("默认") or "无法解析" in tag or "未识别" in tag:
                stat["默认_" + tag] += 1

        proc_raw = ws.cell(r, C_PROC).value
        pa = proc_audit(proc_raw)
        rec[C_PROC] = canon_process(proc_raw)
        rec[C_WE] = nfkc(ws.cell(r, C_WE).value)
        rec[C_Y] = cur_y
        rec[C_Z] = new_z

        rows.append(rec)
        tot = sum(rec[c] for c in C_ELEM) + rec[C_FE]
        notes = list(unk)
        if fe_src == "反算" and any(x in zero_cols for x in ("C", "Si", "Mn", "Cr", "Ni")):
            notes.append("Fe反算受未报主元素影响")
        if abs(tot - 100) > 1:
            notes.append("成分和=%.2f" % tot)
            stat["成分和偏离>1"] += 1
        if not pa["ok"]:
            notes.append("工艺:%s" % pa["cat"])
        if temp_src.startswith("默认") and "室温" not in temp_src and "无数值" not in temp_src:
            notes.append("温度:" + temp_src)
        if "无法解析" in stab_src or "未识别" in stab_src:
            notes.append("稳定时间:" + stab_src)
        if "未识别" in ce_src:
            notes.append("对电极:" + ce_src)
        if "推断" in re_tag:
            notes.append("参比:" + re_tag)

        flags.append([pid, str(name), r, ",".join(zero_cols) if zero_cols else "—",
                      len(zero_cols), fe_src, nfkc(ws.cell(r, C_RE).value),
                      delta, re_tag, nfkc(epit_raw), ep_unit, rec[C_EPIT],
                      nfkc(ws.cell(r, C_TEMP).value), temp, temp_src,
                      nfkc(ws.cell(r, C_STAB).value), stab, stab_src,
                      nfkc(ws.cell(r, C_CE).value), ce, ce_src,
                      pa["cat"], "是" if pa["ok"] else "否",
                      "是" if changed else "否",
                      "；".join(notes) if notes else ""])
    return rows, excluded, flags, notes_cmp, stat


def copy_style(dst, src):
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


HDR1 = {
    1: "文献编号", 2: "样本（原文命名）", 3: "化学成分(wt.%)",
    18: "加工工艺", 19: "工作电极(WE)", 20: "参比电极(RE)", 21: "对电极(CE)",
    22: "腐蚀温度(℃)", 23: "Epit(V)", 24: "测试前稳定时间(min)",
    25: "文献信息", 26: "备注",
}
HDR2 = {3: "C", 4: "Si", 5: "N", 6: "P", 7: "S", 8: "Cu", 9: "Mn",
        10: "Cr", 11: "Mo", 12: "Ni", 13: "Ti", 14: "Nb", 15: "Co",
        16: "Sn", 17: "Fe"}


def make_workbook(ws0, rows, out):
    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = "Sheet1"
    for c in range(1, NCOL + 1):
        sh.cell(1, c).value = HDR1.get(c)
        copy_style(sh.cell(1, c), ws0.cell(1, c))
        sh.cell(2, c).value = HDR2.get(c)
        copy_style(sh.cell(2, c), ws0.cell(2, c))
    for i, rec in enumerate(rows, start=3):
        for c in range(1, NCOL + 1):
            cell = sh.cell(i, c)
            cell.value = rec.get(c)
            copy_style(cell, ws0.cell(3, c))
            if c in (C_PROC, C_Y, C_Z):
                al = copy(cell.alignment)
                al.wrap_text = True
                cell.alignment = al
    merges = ["A1:A2", "B1:B2", "C1:Q1"] + [
        "%s1:%s2" % (chr(64 + c), chr(64 + c)) for c in range(18, 27)]
    for m in merges:
        sh.merge_cells(m)
    for c in range(1, NCOL + 1):
        key = chr(64 + c)
        if key in ws0.column_dimensions:
            sh.column_dimensions[key].width = ws0.column_dimensions[key].width
    sh.column_dimensions["Y"].width = max(sh.column_dimensions["Y"].width or 14, 34)
    sh.column_dimensions["Z"].width = max(sh.column_dimensions["Z"].width or 14, 40)
    sh.sheet_view.zoomScale = 150
    sh.freeze_panes = "C3"
    wb.save(out)
    return sh


def verify_book(path, rows, problems):
    wb = openpyxl.load_workbook(path)
    if wb.sheetnames != ["Sheet1"]:
        problems.append("sheet 不是仅 Sheet1: %r" % wb.sheetnames)
    sh = wb["Sheet1"]
    if sh.max_column != NCOL:
        problems.append("列数 %d != 26" % sh.max_column)
    if sh.max_row != 2 + len(rows):
        problems.append("行数 %d != %d" % (sh.max_row, 2 + len(rows)))
    for c, want in HDR1.items():
        if sh.cell(1, c).value != want:
            problems.append("表头1 第%d列 %r != %r" % (c, sh.cell(1, c).value, want))
    for c, want in HDR2.items():
        if sh.cell(2, c).value != want:
            problems.append("表头2 第%d列 %r != %r" % (c, sh.cell(2, c).value, want))
    blank = 0
    for i in range(3, sh.max_row + 1):
        for c in (C_ID, C_Y, C_Z):
            if sh.cell(i, c).value in (None, ""):
                blank += 1
        for c in C_ELEM + [C_FE, C_TEMP, C_EPIT, C_STAB]:
            v = sh.cell(i, c).value
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                problems.append("r%d c%d 非数值 %r" % (i, c, v))
        if sh.cell(i, C_RE).value != "SCE":
            problems.append("r%d RE != SCE" % i)
        if sh.cell(i, C_CE).value not in ("铂", "石墨"):
            problems.append("r%d CE=%r" % (i, sh.cell(i, C_CE).value))
        ev = sh.cell(i, C_EPIT).value
        if isinstance(ev, (int, float)) and abs(ev) > 2:
            problems.append("r%d Epit 幅值异常 %r" % (i, ev))
        z = sh.cell(i, C_Z).value
        if z is not None and len(str(z)) > 120:
            problems.append("r%d 备注 %d 字 >120" % (i, len(str(z))))
    if blank:
        problems.append("A/Y/Z 空单元格 %d 个" % blank)
    if len(sh.merged_cells.ranges) != 12:
        problems.append("合并区 %d != 12" % len(sh.merged_cells.ranges))
    if sh.sheet_view.zoomScale != 150 or sh.freeze_panes != "C3":
        problems.append("缩放/冻结不一致")
    f = sh.cell(3, 2)
    if f.font.name != "宋体" or abs((f.font.sz or 0) - 11) > 0.01:
        problems.append("字体不是宋体11: %r %r" % (f.font.name, f.font.sz))
    return sh


# ---------------------------------------------------------------- 报告

def write_report(committed, rows, excluded, flags, notes_cmp, pa_map, stat, problems):
    kept = len(rows)
    papers = len({r[C_ID] for r in rows})
    na = sum(1 for e in excluded if e[3] == "无 Epit 数值")
    re_out = sum(1 for e in excluded if e[3] == "参比无法识别")
    p20 = sum(1 for e in excluded if e[3].startswith("文献20"))
    long_before = sum(1 for x in notes_cmp if x[1] > 120)
    trimmed = sum(1 for x in notes_cmp if x[3] == "是")
    ok_proc = sum(1 for p in pa_map.values() if p["ok"])
    L = []
    L.append("# 1.xlsx 清洗校验报告（2026-11-09）\n")
    L.append("## 一、背景\n")
    L.append("- 用户 11-09 指令：直接重建 `1.xlsx`，包含「数据清洗」与「规范表述」两项；"
             "工艺列对照师兄最初图片示例，备注列精简，`1.xlsx` 要带 Y「文献信息」/Z「备注」两列。\n")
    L.append("- 底本 `0.xlsx` 641 行 / 639 条 / 26 列 / 174 块 / 编号 1–177（缺 6/27/39），本轮不改；"
             "旧 `1.xlsx`（275 条 / 24 列）已备份为 `备份\\1_基准_清洗前_20261109.xlsx`。\n")
    L.append("- `temp\\` 已达 300 文件上限，本轮全部产物写 `备份\\清洗1_20261109\\` 与工作区根。\n")
    L.append("\n## 二、方法\n")
    L.append("- 数据清洗沿用 `scripts\\build_1_20261105.py`：元素「未提及/非数值」→0、`<x/≤x/Max x`→上限、"
             "`a-b`/`to`/`a~b`→中值、`a±b`→中心值、`ppm`→÷1e4、`%`/`pct`/`×10⁻¹%`→数值；"
             "Fe 原文实测保留、其余 `100−Σ14元素`；Epit 统一 SCE；温度→℃、稳定时间→min、对电极→铂/石墨；"
             "文献编号与 Y/Z 逐行填值。\n")
    L.append("- 扩展：Epit 支持 mV/mVSCE/`|v|>2` 隐式 mV、`±`/括号±/区间/`≈`/`about`/`<-0.35`；"
             "参比 30 种写法全覆盖；温度支持 K/`±`/区间/多温度均值/日韩表述；稳定时间支持英文词数字与复合段。\n")
    L.append("- 表述规范：工艺列做 NFKC、减号/℃ 记号统一、控制符清理，并按「成形方式 + 热处理制度"
             "（温度/时间/固溶或退火）+ 冷却方式」逐条审计；复杂流程与原文缺参数保留原文，不臆造。\n")
    L.append("- 备注精简：37 块长备注逐块人工压缩为「数据来源 + 特殊情况」，≤120 字；其余 137 块短备注原样保留。\n")
    L.append("- 饱和 Ag/AgCl 沿用旧脚本 **−0.042 V**（`口径基准_20261107.md` 记 −0.044 V，差 0.002 V）；"
             "3M/未注明 −0.031 V；`→SCE` 视为已换算 Δ=0；参比「未提及」不猜值、入未纳入清单。\n")
    L.append("- 校验件 `备份\\清洗1_20261109\\1_校验.xlsx` 逐格检查后才写回，写回后二次复验。\n")
    L.append("\n## 三、结论\n")
    L.append("| 项目 | 数值 |\n|---|---|\n")
    L.append("| 写入状态 | %s |\n" % ("**已写回 `1.xlsx`**" if committed else "仅校验件，未写回"))
    L.append("| `1.xlsx` | %d 行（含 2 行表头）/ %d 条 / 26 列 / 仅 Sheet1 |\n" % (kept + 2, kept))
    L.append("| 覆盖文献 | %d 篇 |\n" % papers)
    L.append("| 未纳入 | %d 条（文献20 %d / 无 Epit %d / 参比无法识别 %d） |\n"
             % (len(excluded), p20, na, re_out))
    L.append("| Epit | 全部数值，参比统一 SCE；无 |Epit|>2 V 异常 |\n")
    L.append("| 备注 | 原 >120 字 %d 块，精简 %d 块；现全部 ≤120 字 |\n" % (long_before, trimmed))
    L.append("| 工艺表述 | %d 种写法，符合师兄示例 %d 种；其余保留原文并列入审计表 |\n"
             % (len(pa_map), ok_proc))
    L.append("| 校验问题项 | **%d** |\n" % len(problems))
    L.append("| `0.xlsx` MD5 | %s（未改） |\n" % md5(SRC))
    L.append("| `1.xlsx` MD5 | %s |\n" % md5(DST)
             if committed else "| `1.xlsx` | 本步未写，仍为清洗前版本 |\n")
    L.append("\n关键计数：元素 0 填充 %d 格、取上限 %d、区间中值 %d、ppm %d；"
             "Fe 实测 %d / 反算 %d；成分和偏离>1 共 %d 行（多为原文未报主元素，已标记）。\n"
             % (stat["元素0填充"], stat["元素取上限"], stat["元素取中值"],
                stat["元素ppm"], stat["Fe_实测"], stat["Fe_反算"],
                stat["成分和偏离>1"]))
    if problems:
        L.append("\n问题项：\n")
        for p in problems[:40]:
            L.append("- %s\n" % p)
    L.append("\n## 四、建议\n")
    L.append("1. 先用 `1_清洗标记_20261109.csv` 复核「成分和偏离>1」「温度默认25」「稳定时间默认30」"
             "「参比推断」等标记行，确认默认值可接受后再用于建模。\n")
    L.append("2. `1_工艺表述审计_20261109.csv` 中「复杂流程/缺参数/需人工核对」的工艺原文保留未改写，"
             "若师兄要求严格按示例压成短句，建议另起一轮逐条人工拆分，不在本脚本里臆造。\n")
    L.append("3. 未纳入的 %d 条中，参比「未提及」%d 条（文献 168/174）如能回原文补出参比，"
             "可另起小批补录；读图试读值仍不回填。\n" % (len(excluded), re_out))
    L.append("4. 回滚：底本 `备份\\1_基准_清洗前_20261109.xlsx`；`0.xlsx` 全程未改；"
             "`AGENTS.md` 改前副本 `备份\\AGENTS_改前_20261109_清洗1.md`。\n")
    path = os.path.join(ROOT, "1_清洗校验_20261109.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("".join(L))
    return path


# ---------------------------------------------------------------- main

def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    assert mode in ("dry", "commit"), mode
    os.makedirs(WORK, exist_ok=True)

    if md5(SRC) != EXPECT_0_MD5:
        raise SystemExit("0.xlsx 已变化，停止：%s" % md5(SRC))
    if not os.path.exists(BASELINE):
        raise SystemExit("缺少备份 %s" % BASELINE)

    wb0 = openpyxl.load_workbook(SRC)
    ws0 = wb0["Sheet1"]
    rows, excluded, flags, notes_cmp, stat = build_records(ws0)

    # 工艺审计表（按规范后写法聚合）
    pa_map = {}
    pid_of = {}
    for rec in rows:
        pa_map.setdefault(rec[C_PROC], proc_audit(rec[C_PROC]))
        pid_of.setdefault(rec[C_PROC], set()).add(rec[C_ID])
    proc_rows = []
    for i, (txt, p) in enumerate(sorted(pa_map.items(), key=lambda kv: -len(pid_of[kv[0]])), 1):
        pids = sorted(pid_of[txt])
        proc_rows.append([i, txt, len(pids), ",".join(str(x) for x in pids),
                          p["cat"], "是" if p["ok"] else "否", p["form"], p["heat"],
                          p["temp"], p["time"], p["cool"]])

    # 校验件
    make_workbook(ws0, rows, STAGE)
    problems = []
    verify_book(STAGE, rows, problems)
    if problems:
        print("校验件问题项 %d：" % len(problems))
        for p in problems[:40]:
            print("  -", p)
        raise SystemExit(2)

    # 五份审计/对照表
    csv_write(os.path.join(ROOT, "1_清洗标记_20261109.csv"),
              ["文献编号", "样本(原文命名)", "0.xlsx行号", "0填充元素", "0填充格数", "Fe来源",
               "原参比", "ΔV", "参比类别", "Epit原文", "Epit单位推断", "Epit(V,SCE)",
               "温度原文", "温度(℃)", "温度来源", "稳定时间原文", "稳定时间(min)", "稳定时间来源",
               "对电极原文", "对电极", "对电极来源", "工艺分类", "工艺符合示例", "备注已精简", "提醒"],
              flags)
    csv_write(os.path.join(ROOT, "1_未纳入_20261109.csv"),
              ["文献编号", "样本(原文命名)", "0.xlsx行号", "类别", "说明"], excluded)
    csv_write(os.path.join(ROOT, "1_工艺表述审计_20261109.csv"),
              ["序号", "规范后工艺", "出现文献数", "文献编号", "分类", "是否符合师兄示例",
               "成形方式", "热处理", "温度", "时间", "冷却方式"], proc_rows)
    csv_write(os.path.join(ROOT, "1_备注精简对照_20261109.csv"),
              ["文献编号", "原备注字数", "精简后字数", "是否精简", "原备注", "1.xlsx备注"],
              notes_cmp)
    rp = write_report(False, rows, excluded, flags, notes_cmp, pa_map, stat, problems)

    print("[dry] 保留 %d 条 / 覆盖 %d 篇 / 未纳入 %d 条 / 工艺写法 %d 种 / 备注精简 %d 块 / 问题项 0"
          % (len(rows), len({r[C_ID] for r in rows}), len(excluded), len(pa_map),
             sum(1 for x in notes_cmp if x[3] == "是")))
    print("[dry] 校验件 ->", STAGE, os.path.getsize(STAGE), "B")
    print("[dry] 报告 ->", rp)

    if mode == "commit":
        if md5(DST) != md5(BASELINE):
            raise SystemExit("1.xlsx 已被其它进程改动，停止写回")
        shutil.copy2(STAGE, DST)
        problems2 = []
        verify_book(DST, rows, problems2)
        if problems2:
            print("[commit] 二次复验失败：", problems2[:20])
            raise SystemExit(3)
        rp2 = write_report(True, rows, excluded, flags, notes_cmp, pa_map, stat, problems2)
        print("[commit] 已写回 1.xlsx，MD5=%s；二次复验问题项 0；报告 -> %s"
              % (md5(DST), rp2))


if __name__ == "__main__":
    main()
