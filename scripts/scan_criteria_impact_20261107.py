# -*- coding: utf-8 -*-
r"""
只读分析 0.xlsx，产出「按 4 条新硬口径的逐文献块影响表」。

用法（pwsh）：
  $env:PYTHONIOENCODING='utf-8'; & 'D:\Python\python.exe' -X utf8 'scripts\scan_criteria_impact_20261107.py'

产物：
  temp\criteria_impact_20261107.csv          60 行逐块影响表
  temp\criteria_impact_说明_20261107.md      背景→方法→结论→建议 + 汇总 + 四张命中清单

严格只读：只用 openpyxl load_workbook(data_only=True) 打开 0.xlsx，不保存、不写回；
不修改、不删除任何既有文件；输出目录用 os.makedirs(exist_ok=True) 由 Python 创建。

判定口径（用户 2026-11-07 新硬口径）：
  ①必须是铸造（不是铸态）类，排除增材制造/粉末/焊接等
  ②必须是奥氏体不锈钢
  ③腐蚀介质必须是 3.5 质量分数 NaCl 溶液
  ④不要两步热处理（敏化、时效、稳定化等）
  解释甲（严格字面）：只收铸造成形件（锻轧排除）、介质严格 3.5wt%（海水/0.6mol/L/3.56% 全排除）、
                      奥氏体单相（双相排除）、无两步热处理。
  解释乙（沿用 11-05 裁定）：「铸造＝成形维度不是铸态」故锻轧可收，仅排除增材/粉末/焊接；
                      介质保留 3.5wt%≡0.6mol/L 等价与天然海水；钢种与两步热处理要求同甲。
  取值：合规 / 存疑 / 不符。
    「不符」＝至少一条硬口径被表内证据明确违反；
    「存疑」＝证据不足或处在裁定边界（Z 未注明介质、成形未注明、组织不明/奥氏体+δ、
              浓度等价或单位存疑、模拟液、pH 分组等）；
    「合规」＝四条硬口径均有正面证据且无触发条款。

可复现性：R/Z/B/Y 原文由脚本从 0.xlsx 抽取；四要素的「人工核定表」MANUAL 内嵌在下方，
逐块写明成形/组织/牌号/介质/热处理步数，便于复核与追改；规则只负责合成判定、抽风险标签与汇总。
介质证据按任务要求只取 Z 备注：Z 未写介质即记「未提及」，不拿题名或库外常识回填，
题名里可见的浓度信息只记在「其它风险」/触发条款里作线索。
热处理步数计数规则：只计「固溶 / 退火 / 再结晶 / 时效 / 敏化 / 稳定化 / 热老化 /
随炉冷却」等成品热处理名；「均匀化 / 均热 / 轧前加热 / 锻造加热 / 熔炼 / 冷轧 / 温轧 /
深冷轧 / 钝化 / 酸洗 / 阴极极化 / 双环再活化」等热机械前处理与表面/试验处理不计步数。
"""
import csv
import json
import os
import re
import unicodedata
from collections import OrderedDict

import openpyxl

ROOT = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
XLSX = os.path.join(ROOT, "0.xlsx")
OUT_CSV = os.path.join(ROOT, "temp", "criteria_impact_20261107.csv")
OUT_MD = os.path.join(ROOT, "temp", "criteria_impact_说明_20261107.md")

# ---------------------------------------------------------------- 人工核定表
# 字段：form 成形方式推断 / micro 钢种组织 / grade 牌号 / med 介质短码 / heat 热处理步数
# med 短码：3.5 / 3.5+（原文并列 mol/L 等价写法）/ 未提及 / 未提及T（题名可见）/
#           未提及R（R 可见）/ 未提及pH / 天然海水 / 3.56 / SME / wtv
# 实测：60 块 Z 里没有 0.6 mol·L⁻¹ 独立写法、人工海水、模拟海水 SME 以外的模拟液、
#       3.65%、5%、其它浓度、酸或 H₂S；3.5% 的「5%」是子串，不是 5% 溶液。
#       0.6 M / 0.61 mol/L 两处都与 3.5% 并列，已归 3.5+，不另立「0.6mol·L⁻¹」类。
MED_LABEL = OrderedDict([
    ("3.5", "3.5wt%"),
    ("3.5+", "3.5wt%（原文并列 0.61 mol/L 或 0.6 M 等价写法；0.6mol·L⁻¹ 未独立出现）"),
    ("未提及", "未提及"),
    ("未提及T", "未提及（Y 题名可见 3.5% NaCl，Z 未写）"),
    ("未提及R", "未提及（R 含 3.5% NaCl 浸泡，Z 未写）"),
    ("未提及pH", "未提及（Z 仅记 pH 3/6/8/10 分组，录入 pH 6 组）"),
    ("天然海水", "天然海水"),
    ("3.56", "3.56%"),
    ("SME", "模拟海水SME（3.5 wt% NaCl 配制）"),
    ("wtv", "3.5wt%（原文 3.5% wt./v，非质量分数）"),
    # 以下类别为任务枚举项，实测 60 块 Z 未出现，保留映射以保证脚本对后续扩展可复用：
    ("0.6mol", "0.6mol·L⁻¹"),
    ("人工海水", "人工海水"),
    ("3.65", "3.65%"),
    ("5%", "5%"),
    ("其它浓度", "其它浓度"),
    ("酸或H2S", "酸或H₂S"),
])

# fmt: id -> (form, micro, grade, med, heat)
MANUAL = {
    "1":  ("铸态", "奥氏体", "HC-ASS / LC-ASS（高碳铸造奥氏体）", "未提及", "1"),
    "2":  ("铸态", "奥氏体", "CF3M（铸造 316L 级）", "未提及", "1"),
    "3":  ("砂型/离心", "奥氏体", "G-X 10CrNiMoNb 18-10", "未提及", "2+"),
    "4":  ("铸态", "不明（Cr-Mn 系）", "CrMn-xCu", "未提及", "0"),
    "5":  ("铸态", "不明（Cr-Mn 系）", "CrMn-xNi", "未提及T", "0"),
    "7":  ("未注明", "奥氏体", "304（敏化态）", "未提及", "2+"),
    "8":  ("未注明", "奥氏体", "1Cr18Ni9Ti / 304 / 316", "未提及", "0"),
    "9":  ("铸锭→热锻→冷轧", "奥氏体", "AISI 304 系（Mn/Mo 合金化）", "未提及", "0"),
    "10": ("铸锭→热锻→冷轧", "奥氏体", "AISI 304 系（Cu/Sn 合金化）", "未提及", "0"),
    "11": ("锻轧", "奥氏体", "AISI 316L", "未提及T", "1"),
    "12": ("铸态", "奥氏体", "SCS14 铸造奥氏体不锈钢", "天然海水", "1"),
    "13": ("铸态", "奥氏体+δ", "SCS14 系铸造奥氏体（δ 铁素体变量）", "天然海水", "1"),
    "14": ("铸态", "奥氏体", "超低碳低铬镍高硅铸造不锈钢 / C15", "未提及", "1"),
    "15": ("铸态", "奥氏体", "超低碳含铜铸造不锈钢", "未提及", "1"),
    "16": ("铸态", "奥氏体", "CZHASS-1 超低碳铸造高合金不锈钢", "未提及", "2+"),
    "17": ("砂型", "奥氏体", "304 + 超细 SiC 粉体", "未提及", "0"),
    "18": ("铸态", "奥氏体", "Z3CN20-09M（核级铸造）", "未提及", "1"),
    "19": ("熔模", "奥氏体", "CF8", "未提及", "1"),
    "20": ("铸态+锻轧（混合）", "奥氏体", "316LN / 新型奥氏体不锈钢", "未提及", "2+"),
    "21": ("真空精密", "奥氏体", "精密铸造不锈钢叶轮（未给牌号）", "未提及", "0"),
    "22": ("铸造", "奥氏体", "NHG1 / 316L", "未提及", "1"),
    "23": ("铸造", "奥氏体", "Z3CN20.09M（核级铸造）", "未提及", "1"),
    "24": ("热轧", "高氮", "Fe-16Cr-14.5Mn-0.73N 高氮奥氏体", "未提及", "0"),
    "25": ("锻轧/热轧", "奥氏体", "204Cu + Ce", "未提及", "1"),
    "26": ("未注明", "高氮", "HNASS 高氮奥氏体", "未提及", "2+"),
    "28": ("锻轧/热轧", "奥氏体", "超级奥氏体 Mo-N 系（Mo2N1/Mo1N1/Mo1N2）", "3.5", "1"),
    "29": ("热轧", "奥氏体", "304（Nb-free / 0.05Nb）", "3.5", "1"),
    "30": ("热轧", "奥氏体", "309S + N/Nb", "3.5+", "1"),
    "31": ("热轧", "奥氏体", "易切削奥氏体不锈钢（S 系列）", "3.5", "1"),
    "32": ("冷轧", "奥氏体", "304L", "未提及", "2+"),
    "33": ("冷轧", "奥氏体", "316LN", "3.5", "2+"),
    "34": ("冷轧", "奥氏体", "超级奥氏体（Cr26.7/Ni32.8/Mo4.4）", "3.5+", "2+"),
    "35": ("热轧/冷轧", "奥氏体", "310s", "3.5", "1"),
    "36": ("冷轧", "奥氏体", "304L（UFG）", "SME", "2+"),
    "37": ("未注明（0% 温预应变、未轧制）", "奥氏体", "304 / 316", "3.5", "1"),
    "38": ("未注明（退火态板材）", "奥氏体", "201A/201B/201Cu/204Cu/204CuNb/304", "3.56", "1"),
    "40": ("铸态", "奥氏体", "304（Mg/S 系列）", "3.5", "0"),
    "41": ("热轧", "奥氏体", "321 / 310s", "3.5", "1"),
    "42": ("热轧/冷轧", "奥氏体", "亚稳奥氏体 Fe-17.9Cr-7.56Ni", "3.5", "2+"),
    "43": ("锻轧/热轧", "奥氏体", "Super304H + Nb", "未提及", "2+"),
    "44": ("未注明（约 5 mm 轧板取样）", "高氮", "QN2109 高氮奥氏体", "3.5", "0"),
    "45": ("温轧（锻轧类）", "奥氏体", "304", "未提及", "1"),
    "46": ("热轧", "奥氏体", "254SMO / SASSNb 超级奥氏体", "未提及", "1"),
    "47": ("温轧（锻轧类）", "奥氏体", "304", "未提及", "1"),
    "48": ("未注明（304 弯管）", "奥氏体", "304", "未提及T", "1"),
    "49": ("铸态", "奥氏体", "Super 304H", "未提及R", "2+"),
    "50": ("铸造", "不明（耐热铸钢）", "节镍经济型耐热铸钢 / ZG40Cr25Ni20", "未提及", "1"),
    "51": ("锻轧", "奥氏体", "316L", "未提及", "1"),
    "52": ("冷轧", "高氮", "0.76N / 0.92N 无镍高氮奥氏体", "未提及", "1"),
    "53": ("铸态+热轧（混合）", "奥氏体", "304 / 304L", "未提及", "2+"),
    "54": ("铸锭→热锻→冷轧", "奥氏体", "18Cr-12Ni（-0.1N / -3Mo）", "未提及pH", "1"),
    "55": ("热轧/冷轧", "奥氏体", "304 / 304Cu", "未提及T", "1"),
    "56": ("铸锭→热轧", "奥氏体", "Fe-18Cr-9Mn-5.5Ni-0.3(C+N)", "3.5", "1"),
    "57": ("锻轧", "高氮", "Fe-22Cr-16Mn-0.55N 高氮奥氏体", "3.5", "1"),
    "58": ("冷轧", "奥氏体", "1Cr18Ni9Ti / 304 / 316L", "3.5", "1"),
    "59": ("未注明（重水堆引漏管取样）", "奥氏体", "304L / 316L", "3.5", "0"),
    "60": ("热轧", "奥氏体", "ASTM A-240 TP321", "wtv", "1"),
    "61": ("铸造", "奥氏体", "304 + Y", "3.5", "0"),
    "62": ("冷轧", "奥氏体", "Fe-18Cr-8Ni（商用冷轧板）", "3.5", "1"),
    "63": ("冷轧", "奥氏体", "201 / 202 / 304", "3.5", "0"),
}

CAST_FORMS = {"铸态", "铸造", "砂型", "熔模", "离心", "真空精密", "砂型/离心"}
AM_FORMS = {"增材", "焊接", "粉末"}
MIXED_FORMS = {"铸态+锻轧（混合）", "铸态+热轧（混合）"}


def nfkc(x):
    return unicodedata.normalize("NFKC", str(x if x is not None else ""))


def s(x):
    return nfkc(x).strip()


def raw(x):
    """输出「原文」字段用：保留 0.xlsx 原字形。

    NFKC 会把 ℃→°C、₂→2、全角→半角，适合关键词匹配但不适合当原文引用。
    """
    return ("" if x is None else str(x)).strip()


def clip(x, n):
    x = x.replace("\r", " ").replace("\n", " ⏎ ")
    return x if len(x) <= n else x[:n - 1] + "…"  # 含省略号在内严格不超过 n 字


def uniq(seq):
    seen, out = set(), []
    for x in seq:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


# ---------------------------------------------------------------- 读取 0.xlsx
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Sheet1"]

# 合并区首值展开（用于 R/Z/Y 等块级字段）；块首判断绝不能用展开后的 A 值
merged = {}
for rng in ws.merged_cells.ranges:
    v = ws.cell(rng.min_row, rng.min_col).value
    for r in range(rng.min_row, rng.max_row + 1):
        for c in range(rng.min_col, rng.max_col + 1):
            merged[(r, c)] = v


def val(r, c):
    return merged.get((r, c), ws.cell(r, c).value)


# 数据从 r3 起；A 列仅块首行物理单元格有值
starts = [r for r in range(3, ws.max_row + 1)
          if ws.cell(r, 1).value not in (None, "")]
assert len(starts) == 60, "文献块数应为 60，实际 %d" % len(starts)

blocks = []
for i, st in enumerate(starts):
    en = starts[i + 1] - 1 if i + 1 < len(starts) else ws.max_row
    rows = list(range(st, en + 1))
    # 去重按 NFKC 归一值（℃ 与 °C 视为同一工艺），输出保留 0.xlsx 原字形。
    seen_R, R_unique, R_norm = set(), [], []
    for r in rows:
        rv, sv = raw(val(r, 18)), s(val(r, 18))
        if sv and sv not in seen_R:
            seen_R.add(sv)
            R_unique.append(rv)
            R_norm.append(sv)
    blocks.append({
        "id": s(ws.cell(st, 1).value),
        "n": len(rows),
        "samples": [raw(val(r, 2)) for r in rows],
        "R_unique": R_unique,
        "R_match": " ｜ ".join(R_norm),
        "Z": raw(val(st, 26)),
        "Y": raw(val(st, 25)),
        "Z_match": s(val(st, 26)),
        "Y_match": s(val(st, 25)),
        "W": [s(val(r, 23)) for r in rows],
    })
assert sum(b["n"] for b in blocks) == 283, "数据行合计应为 283"
assert set(b["id"] for b in blocks) == set(MANUAL), "核定表与文献编号不一致"


# ---------------------------------------------------------------- 风险标签
def risk_tags(b, form, med, heat):
    z, r, y, w = b["Z_match"], b["R_match"], b["Y_match"], b["W"]
    txt = " ⏎ ".join([z, r, y, " ｜ ".join(b["samples"])])
    tags = []

    def hit(pat):
        return re.search(pat, txt, re.I) is not None

    if any(x in ("N/A", "未提及") for x in w) or "未测出击穿电位" in z:
        tags.append("Epit=N/A 或占位")
    if hit(r"WebPlotDigitizer|图读|读图|Fig\.\s*\d|图\s*\d|渲染|裁剪放大|逐格读出|图上读数"):
        tags.append("图读值")
    if hit(r"成分原文未列|标称成分|无逐元素|未列具体数值|成分表为图片|无数值型点蚀电位"):
        tags.append("无逐元素成分表/标称成分")
    if hit(r"pH\s*3/6/8/10|pH\s*序列|酸调|HCl\s*调|pH\s*3\b"):
        tags.append("pH 分组/酸调待核")
    if hit(r"Fe\s*记法|Fe\s*按基准表|Fe\s*原文|Fe\s*未单列|Fe\s*记.未提及."
           r"|Fe\s*列原文仅.*Bal|统一记.余量.|Fe\s*按.*余量|Fe\s*记.余量|不要按\s*Fe|Fe\s*=\s*100"):
        tags.append("Fe 记法（余量/未提及，未反算）")
    if hit(r"Ag/AgCl|参比.*换算|换算为 SCE|饱和甘汞"):
        tags.append("参比口径/换算")
    if hit(r"上限值|右删失|<\s*-\s*0\.35|取上限"):
        tags.append("Epit 上限值（右删失）")
    if hit(r"as[- ]?received|As received|供货态|冷轧态对照|未退火"):
        tags.append("含未热处理/原始态对照行")
    if hit(r"统计分布|累积概率|中位数区间"):
        tags.append("统计分布非逐试样值")
    if hit(r"二手数据|引自\s*Pereira|分投两刊"):
        tags.append("含二手/同批数据说明")
    if hit(r"推断值而非图上读出|正文推断值"):
        tags.append("推断值非实测")
    if hit(r"SiC\s*粉体"):
        tags.append("SiC 粉体增强相（砂型铸造，非粉末冶金）")
    if hit(r"wt\./v"):
        tags.append("浓度单位 wt./v 非质量分数")
    if hit(r"模拟海水|SME|模拟重水堆|模拟体液|SBF"):
        tags.append("含模拟液/其它介质组")
    if hit(r"完全除气|通氮除氧|阴极极化"):
        tags.append("除气/阴极预极化条件")
    if hit(r"3\.56\s*wt|ASTM\s*G61"):
        tags.append("3.56% 与库内文献27 处置不一致")
    if hit(r"Z3CN20|核级铸造"):
        tags.append("Z3CN20-09M 常含 δ 铁素体，表内未直写组织")
    if "铸态+锻轧（混合）" == form or "铸态+热轧（混合）" == form:
        tags.append("块内铸造/锻轧样本混合")
    if med.startswith("未提及"):
        tags.append("Z 备注未记介质，须回原文补证")
    return "；".join(uniq(tags)[:8]) or "无"


# ---------------------------------------------------------------- 判定合成
def verdict(expl, form, micro, med, heat):
    fails, doubts = [], []

    # ① 成形
    if form in AM_FORMS:
        fails.append("①增材/焊接/粉末成形")
    elif form in MIXED_FORMS:
        if expl == "甲":
            fails.append("①块内含锻轧样本，非纯铸造成形")
    elif form in CAST_FORMS:
        pass
    elif form.startswith("未注明"):
        doubts.append("①成形方式未注明")
    else:
        if expl == "甲":
            fails.append("①锻轧/轧制件，非铸造成形")

    # ② 组织
    if micro.startswith("双相") or micro.startswith("马氏体或铁素体"):
        fails.append("②非奥氏体单相")
    elif micro.startswith("奥氏体+δ"):
        doubts.append("②奥氏体+δ，单相要求待裁定")
    elif micro.startswith("不明"):
        doubts.append("②组织未直写，是否奥氏体不锈钢待裁定")
    # 奥氏体 / 高氮 视为满足（高氮属奥氏体不锈钢）

    # ③ 介质
    if med in ("3.5", "3.5+"):
        pass
    elif med == "天然海水":
        if expl == "甲":
            fails.append("③天然海水非严格 3.5wt% NaCl 溶液")
    elif med == "0.6mol":
        if expl == "甲":
            fails.append("③0.6mol·L⁻¹ 写法非严格 3.5 质量分数（甲明确排除）")
        # 乙：师兄 :50/:63/:105/:205 直述 3.5 wt.% 与 0.6 mol/L 等价，视为合规
    elif med == "人工海水":
        if expl == "甲":
            fails.append("③人工海水非严格 3.5wt% NaCl 溶液")
        else:
            doubts.append("③人工海水（乙只直述天然海水，按 11-07 补充裁定待师兄确认）")
    elif med == "3.56":
        if expl == "甲":
            fails.append("③3.56% NaCl 非严格 3.5wt%（甲明确排除，库内文献27 曾因此整篇删）")
        else:
            doubts.append("③3.56 wt% NaCl 略偏离 3.5，库内文献27 曾因此删除，待裁定")
    elif med == "SME":
        if expl == "甲":
            fails.append("③模拟海水SME 非 3.5wt% NaCl 纯溶液")
        else:
            doubts.append("③模拟液 SME，是否视同 3.5wt% 待裁定")
    elif med == "wtv":
        if expl == "甲":
            fails.append("③原文为 3.5% wt./v（质量体积浓度），非 3.5 质量分数")
        else:
            doubts.append("③原文为 3.5% wt./v（质量体积浓度），非 3.5 质量分数，待裁定")
    elif med in ("3.65", "5%", "其它浓度", "酸或H2S"):
        fails.append("③介质浓度/体系不符：%s" % MED_LABEL[med])
    elif med.startswith("未提及"):
        doubts.append("③Z 未注明介质，无法在表内核验硬口径③")

    # ④ 热处理
    if heat == "2+":
        fails.append("④两步热处理（敏化/时效/稳定化或两次热处理）")

    if fails:
        v = "不符"
    elif doubts:
        v = "存疑"
    else:
        v = "合规"
    trig = "；".join(fails + doubts) if (fails or doubts) else "四条硬口径均有正面证据，无触发条款"
    return v, trig


# ---------------------------------------------------------------- 生成 CSV 行
HEADERS = [
    "编号", "样本数", "样本名示例", "加工工艺原文", "备注原文",
    "成形方式推断", "钢种组织", "牌号", "介质", "热处理步数", "热处理原文佐证",
    "甲_判定", "甲_触发条款", "乙_判定", "乙_触发条款", "其它风险",
]

out_rows = []
for b in blocks:
    bid = b["id"]
    form, micro, grade, med, heat = MANUAL[bid]
    va, ta = verdict("甲", form, micro, med, heat)
    vb, tb = verdict("乙", form, micro, med, heat)

    Rs = b["R_unique"]
    # 佐证必须取「真正体现步数」的那条工艺；块内常有 as-received / 未固溶 对照行，
    # 它们含「未时效」等否定词，不能拿来当 2+ 的原文佐证。
    neg_pat = r"未固溶|未时效|未退火|未热处理|无热处理|未经任何热处理|as[\s-]?received|供货态"
    Rm = [s(x) for x in Rs]  # NFKC 归一，兼容 ℃ / °C

    def hits(pat, positive_only=True):
        return [Rs[i] for i, x in enumerate(Rm)
                if re.search(pat, x, re.I)
                and not (positive_only and re.search(neg_pat, x, re.I))]

    if heat == "2+":
        # 按证据强度分三级：①明确两次热处理名 → ②两段温度并列 → ③多段箭头工艺。
        # 53 的单步固溶条也带 →K，故 ③ 只作最后兜底。
        pats = [
            r"敏化|时效|稳定化|热老化|退火.*退火|固溶.*(?:退火|时效|敏化|稳定化|再结晶|固溶)|再结晶.*退火",
            r"\+\s*\d+\s*(?:℃|°C|K|min|h)",
            r"→.*(?:℃|°C|K)",
        ]
        cs = []
        for p in pats:
            cs = hits(p)
            if cs:
                break
        quote = cs[0] if cs else (Rs[0] if Rs else "")
    elif heat == "1":
        cs = hits(r"固溶|退火|再结晶|热老化|随炉冷却|保温")
        quote = cs[0] if cs else (Rs[0] if Rs else "")
    else:
        cs = [Rs[i] for i, x in enumerate(Rm) if re.search(neg_pat, x, re.I)]
        quote = cs[0] if cs else (Rs[0] if Rs else "")

    out_rows.append({
        "编号": bid,
        "样本数": b["n"],
        "样本名示例": "；".join(b["samples"][:2]),
        "加工工艺原文": clip(" ｜ ".join(Rs[:3]), 400),
        "备注原文": clip(b["Z"], 200),
        "成形方式推断": form,
        "钢种组织": micro,
        "牌号": grade,
        "介质": MED_LABEL[med],
        "热处理步数": heat,
        "热处理原文佐证": clip(quote, 200),
        "甲_判定": va,
        "甲_触发条款": ta,
        "乙_判定": vb,
        "乙_触发条款": tb,
        "其它风险": risk_tags(b, form, med, heat),
    })

os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=HEADERS)
    w.writeheader()
    w.writerows(out_rows)


# ---------------------------------------------------------------- 汇总
def summary(key):
    d = OrderedDict((k, [0, 0]) for k in ("合规", "存疑", "不符"))
    for b, row in zip(blocks, out_rows):
        k = row[key]
        d[k][0] += 1
        d[k][1] += b["n"]
    return d


sa, sb = summary("甲_判定"), summary("乙_判定")
total_blocks, total_rows = len(blocks), sum(b["n"] for b in blocks)

heat_hit = [r["编号"] for r in out_rows if r["热处理步数"] == "2+"]
med_hit = [r["编号"] for r in out_rows
           if r["介质"] in ("天然海水", "3.56%", "模拟海水SME（3.5 wt% NaCl 配制）",
                            "0.6mol·L⁻¹", "人工海水", "3.65%", "5%", "其它浓度", "酸或H₂S")
           or r["介质"].startswith("3.5wt%（原文 3.5% wt./v")]
micro_hit = [r["编号"] for r in out_rows if r["钢种组织"].startswith(("双相", "奥氏体+δ"))]
am_hit = [r["编号"] for r in out_rows if r["成形方式推断"] in AM_FORMS]
med_missing = [r["编号"] for r in out_rows if r["介质"].startswith("未提及")]


def ids(xs):
    return "、".join(xs) if xs else "无"


def sum_line(d):
    return " / ".join("%s %d 块、%d 条" % (k, v[0], v[1]) for k, v in d.items())


a_ok = sa["合规"][1]
a_ok_blocks = sa["合规"][0]
a_upper_rows = sa["合规"][1] + sa["存疑"][1]
a_upper_blocks = sa["合规"][0] + sa["存疑"][0]
b_upper_rows = sb["合规"][1] + sb["存疑"][1]
b_upper_blocks = sb["合规"][0] + sb["存疑"][0]

md = f"""# 4 条新硬口径对 0.xlsx 的逐文献块影响表（说明）

## 一、背景

用户 2026-11-07 提出 4 条新硬口径：①必须是铸造（不是铸态）类，排除增材制造/粉末/焊接等；②必须是奥氏体不锈钢；③腐蚀介质必须是 3.5 质量分数 NaCl 溶液；④不要两步热处理（敏化、时效、稳定化等）。为评估这 4 条口径对现有主库的影响，本次只读分析 `0.xlsx` 的 `Sheet1`：285 行 / 26 列 / **283 条数据 / 60 个文献块**（缺号 6、27、39），不打开 `文献\\*.pdf`，不参照 `1.xlsx` / `2.xlsx`。

两套解释并列评估：**解释甲（严格字面）**只收铸造成形件（锻轧排除）、介质严格 3.5wt%（海水 / 0.6 mol·L⁻¹ / 3.56% 全排除）、奥氏体单相、无两步热处理；**解释乙（沿用 11-05 裁定）**「铸造＝成形维度不是铸态」，锻轧可收、仅排除增材/粉末/焊接，介质保留 3.5wt%≡0.6 mol·L⁻¹ 等价与天然海水，钢种与两步热处理要求同甲。判定取值 `合规 / 存疑 / 不符`。

## 二、方法

1. **只读抽取**：`openpyxl.load_workbook(data_only=True)` 打开 `0.xlsx`，不保存、不写回；`temp\\`、`scripts\\` 目录由 Python `os.makedirs(exist_ok=True)` 创建。既有的 `0.xlsx`、`1.xlsx`、`2.xlsx` 与任何其它文件均未修改、未删除。
2. **块级还原**：表头占两行（r1 分类名、r2 元素符号），数据从 r3 起；A 列文献编号只在每块首行的物理单元格有值，故块首用 `ws.cell(r,1).value` 判断，绝不能用合并区展开后的 A 值（会把块内每行都填成编号）。Y / Z 为按块竖合并，取块首行原值；R 列按块内去重、取前 3 条；备注原文截断 200 字。
3. **四要素判定**：依据 R（成形、热处理）、Z（介质、备注）、B（样本名/牌号）关键词并 NFKC 全角归一，逐块人工核定成形方式、钢种组织与牌号、介质、热处理步数；核定表内嵌于 `scripts\\scan_criteria_impact_20261107.py` 的 `MANUAL`，可逐块追改。
4. **介质证据口径**：按任务要求**只取 Z 备注**，NFKC 全角归一后按用户枚举的类别判定：3.5wt% / 0.6mol·L⁻¹ / 天然海水 / 人工海水 / 模拟海水 SME / 3.56% / 3.65% / 5% / 其它浓度 / 未提及 / 酸或 H₂S。Z 未写介质一律记「未提及」，不拿 Y 题名、牌号常识或库外资料回填；题名可见的浓度信息只写进「其它风险」和触发条款作线索。实测 60 块 Z 里：30 写「3.5%(0.61 mol/L) NaCl」、34 写「3.5%(0.6 M) NaCl」，属 3.5wt% 并列等价写法，不另立 0.6mol·L⁻¹ 类；**未见**独立的 0.6mol·L⁻¹、人工海水、3.65%、5%、其它浓度、酸或 H₂S（「3.5%」里的「5%」是子串，不是 5% 溶液）；28 的 Z 提到 HCl 循环极化是另一组试验、不属录入数据介质。这是本次最大影响面之一：早期文献的 Z 备注普遍只记数据来源与图号。
5. **热处理步数计数**：只计「固溶 / 退火 / 再结晶 / 时效 / 敏化 / 稳定化 / 热老化 / 随炉冷却」等成品热处理名；「均匀化 / 均热 / 轧前加热 / 锻造加热 / 熔炼 / 冷轧 / 温轧 / 深冷轧 / 钝化 / 酸洗 / 阴极极化 / 双环再活化」等前处理与试验处理不计步数。块内多种工艺并存时按最高步数记，并在「热处理原文佐证」附原文。
6. **三档规则**：「不符」＝至少一条硬口径被表内证据明确违反；「存疑」＝证据不足或处于裁定边界（Z 未注明介质、成形未注明、组织不明 / 奥氏体+δ、3.56% 与 wt./v 单位、模拟液、pH 分组等）；「合规」＝四条均有正面证据且无触发条款。汇总按样本数加权得到条数。

## 三、结论

**结论先行：按解释甲严格收紧，库内只剩 {a_ok_blocks} 块 / {a_ok} 条确定合规；即便把全部「存疑」块按最宽口径放行，也只有 {a_upper_blocks} 块 / {a_upper_rows} 条，仍达不到师兄「150 条以上」目标（缺口至少 {150 - a_upper_rows} 条）。**

### 3.1 两套解释的汇总（块数 / 按样本数加权条数）

| 解释 | 合规 | 存疑 | 不符 | 合计 |
|---|---|---|---|---|
| 甲（严格字面） | {sa['合规'][0]} 块 / {sa['合规'][1]} 条 | {sa['存疑'][0]} 块 / {sa['存疑'][1]} 条 | {sa['不符'][0]} 块 / {sa['不符'][1]} 条 | {total_blocks} 块 / {total_rows} 条 |
| 乙（11-05 裁定） | {sb['合规'][0]} 块 / {sb['合规'][1]} 条 | {sb['存疑'][0]} 块 / {sb['存疑'][1]} 条 | {sb['不符'][0]} 块 / {sb['不符'][1]} 条 | {total_blocks} 块 / {total_rows} 条 |

- 解释甲：{sum_line(sa)}；严格合规 {a_ok_blocks} 块 / {a_ok} 条，含存疑上限 {a_upper_blocks} 块 / {a_upper_rows} 条。
- 解释乙：{sum_line(sb)}；严格合规 {sb['合规'][0]} 块 / {sb['合规'][1]} 条，含存疑上限 {b_upper_blocks} 块 / {b_upper_rows} 条。
- 甲下只剩文献 **40（9 条，纯铸态 304 + 3.5 wt% NaCl + 无热处理）** 与 **61（5 条，铸造 304+Y + 3.5% NaCl + 未经热处理）** 两块确定合规；两块的其它风险分别是「Fe 记法（余量/未提及，未反算）；参比口径/换算」与「Epit 上限值（右删失）」（对应原文 `<−0.35`），均不属四条硬口径。
- **影响面分解**：甲下不符的主因是①锻轧/轧制件（库内按 11-05 裁定收录了大量热轧+固溶、冷轧+退火数据），叠加③Z 未记介质与④两步热处理；乙下锻轧不再排除，剩量主要被③Z 未记介质（{len(med_missing)} 块）与④两步热处理（{len(heat_hit)} 块）压住。
- 「存疑」不是「介质不符」：{len(med_missing)} 块的 Z 备注根本没有介质字段，新硬口径③在表内无法核验，须回原文补证或由用户裁定是否按库内既有收录口径视为合规。若把这部分一律判「不符」，两套解释的合规数还会更低。

### 3.2 四张命中清单

| 清单 | 命中块（编号） | 块数 |
|---|---|---|
| ① 两步热处理 | {ids(heat_hit)} | {len(heat_hit)} |
| ② 非 3.5wt% 介质 | {ids(med_hit)} | {len(med_hit)} |
| ③ 双相 / 奥氏体+δ | {ids(micro_hit)} | {len(micro_hit)} |
| ④ 增材 / 焊接 / 粉末 | {ids(am_hit)} | {len(am_hit)} |

- ① 两步热处理 {len(heat_hit)} 块：含「固溶+稳定化」（3）、「固溶+敏化」（7、16、53）、「固溶+时效」（26、34、49）、「固溶+退火 / 两次退火 / 两次固溶」（20、32、33、36、42、43）等，两套解释下均判「不符」。
- ② 非 3.5wt% 介质 {len(med_hit)} 块：12、13 为**天然海水**（甲不符、乙合规/存疑），36 为**模拟海水 SME**，38 为 **3.56 wt% NaCl（ASTM G61 标准液）**，60 原文写 **3.5% wt./v**（质量体积浓度，非 3.5 质量分数）。30、34 虽并列 0.61 mol/L、0.6 M，但 Z 明确写 3.5% NaCl，按乙等价、按甲仍属 3.5wt%，未列入；60 块 Z 未出现独立的 0.6mol·L⁻¹、人工海水、3.65%、5%、其它浓度、酸或 H₂S。另有 {len(med_missing)} 块 Z **未提及**介质（编号：{ids(med_missing)}），属证据缺失而非浓度不符，单列备查。
- ③ 双相 / 奥氏体+δ：仅 **13**（Y 日文题名含「δフェライト」、Z 复核段出现「δ≈1.3」，判「奥氏体+δ」）。库内**无真双相块**；2 的 Y 题名含 "Cast Lean Duplex"，但样本只取 CF3M 奥氏体；18、23 为核级铸造 Z3CN20-09M，表内未直写 δ，已写入其它风险建议回原文核组织。
- ④ 增材 / 焊接 / 粉末：**无**。库内 60 块未见增材制造、焊接/堆焊、粉末冶金成形件（旧口径已排除）。17 为「砂型铸造 + 超细 SiC 粉体强化」，粉体是外加增强相而非粉末冶金成形，已在其它风险标注，避免被「粉末」字面误伤。

## 四、建议

1. **先裁定口径，再动数据**：甲会砍掉 {total_blocks - a_upper_blocks} 块以上且即便放行存疑仍不足 150 条；若目标是保留「150 条以上」可建模数据量，甲不可单独执行，须同时放宽①（回到乙的锻轧可收）并补齐③的介质证据。
2. **优先补证而非删数**：{len(med_missing)} 块 Z 未记介质，建议回 `文献\\01.pdf`–`63.pdf` 方法段补录「介质」字段（不改动 0.xlsx 现有列，可在清洗版加列或另建证据台账），把「存疑」转成有证据的判定；这一步对甲、乙都增益最大。
3. **两步热处理 13 块单独建子集**：它们本身是「成分 × 热处理」维度的数据，若师兄 10-29「不要两步热处理」只针对建模子集而非收录门槛，可在分组表里单列、不进主建模集，避免整块作废。
4. **补真铸造文献**：库内真铸造牌号极稀缺（AGENTS.md 记「真铸造只有 2 篇」），甲下只剩 2 块合规即由此而来；若要同时满足「铸造硬口径」与「150 条以上」，须新增约 {150 - a_upper_rows} 条以上铸造奥氏体不锈钢数据，靠现有 `准\\` 191 篇待入库件优先筛铸造件。
5. **边界项提请裁定**：38 的 3.56 wt%（库内文献27 曾因同浓度整篇删除）、60 的 wt./v 写法、36 的模拟海水 SME、13 的奥氏体+δ、50 的耐热铸钢是否算奥氏体不锈钢，均按「存疑」留痕，不擅自判合规或不符。
6. 本次为只读影响评估，**未改 `0.xlsx` / `1.xlsx` / `2.xlsx`，未删除任何文件**；逐块依据见 `temp\\criteria_impact_20261107.csv`（60 行）。

---
数据来源：`0.xlsx` Sheet1（283 条 / 60 块）；判定脚本 `scripts\\scan_criteria_impact_20261107.py`；用户 2026-11-07 四条新硬口径及甲/乙解释。
"""

with open(OUT_MD, "w", encoding="utf-8", newline="\n") as f:
    f.write(md)

print("CSV:", OUT_CSV, "rows=", len(out_rows))
print("MD :", OUT_MD)
print("甲:", json.dumps(sa, ensure_ascii=False))
print("乙:", json.dumps(sb, ensure_ascii=False))
print("两步热处理:", ids(heat_hit))
print("非3.5wt%介质:", ids(med_hit))
print("双相/奥氏体+δ:", ids(micro_hit))
print("增材/焊接/粉末:", ids(am_hit))
print("Z未提及介质:", ids(med_missing))
