# -*- coding: utf-8 -*-
"""只读核查 0.xlsx（2026-11-08）。

目的：
  1) 按 口径基准_20261107.md 定稿 N1-N4 + A1-A5 出逐块「合规/存疑/不符」判级；
  2) 检查 0.xlsx 内重复样本（多层指纹，只报告不改动）。

只读：0.xlsx、temp/criteria_impact_20261107.csv。
写出：0xlsx_口径核查_底稿_20261107.csv、0xlsx_重复样本检查_底稿_20261108.csv、
      temp/audit_0xlsx_20261108_stats.json。
不改 0.xlsx / 1.xlsx / 2.xlsx / 文献 / 准 / 待 / AGENTS.md / 口径基准_20261107.md。
"""
import csv
import json
import os
import re
import unicodedata
from collections import defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
X0 = os.path.join(ROOT, "0.xlsx")
CSV_IMPACT = os.path.join(ROOT, "temp", "criteria_impact_20261107.csv")
OUT_JUDGE = os.path.join(ROOT, "0xlsx_口径核查_底稿_20261107.csv")
OUT_DUP = os.path.join(ROOT, "0xlsx_重复样本检查_底稿_20261108.csv")
OUT_STAT = os.path.join(ROOT, "temp", "audit_0xlsx_20261108_stats.json")


def nfkc(s):
    if s is None:
        return ""
    return unicodedata.normalize("NFKC", str(s)).strip()


def norm(s):
    """数值型按 6 位小数归一，文本型按 NFKC 归一；区分 N/T 前缀避免 0.34 撞 '0.34'。"""
    t = nfkc(s)
    if t == "":
        return ""
    try:
        return "N" + repr(round(float(t), 6))
    except ValueError:
        return "T" + t


# ---------------- 1. 读 0.xlsx（26 列完整版，Y/Z 合并区按块首行取值） ----------------
wb = load_workbook(X0, data_only=True, read_only=True)
ws = wb["Sheet1"]
grid = [list(row) for row in ws.values]
wb.close()
maxr = len(grid)
maxc = max(len(r) for r in grid)

block_rows = []
for r in range(3, maxr + 1):
    v = nfkc(grid[r - 1][0]) if len(grid[r - 1]) > 0 else ""
    if v != "":
        block_rows.append((r, v))
row2block = {}
for i, (r, bid) in enumerate(block_rows):
    end = block_rows[i + 1][0] - 1 if i + 1 < len(block_rows) else maxr
    for rr in range(r, end + 1):
        row2block[rr] = (bid, r)


def cell(r, c):
    """1-based 列号；Y(25)/Z(26) 为空时回退到所在块首行（合并单元格）。"""
    row = grid[r - 1]
    v = row[c - 1] if len(row) >= c else None
    if c in (25, 26) and nfkc(v) == "":
        _, br = row2block.get(r, ("", r))
        brow = grid[br - 1]
        v = brow[c - 1] if len(brow) >= c else None
    return nfkc(v)


rows = []
for r in range(3, maxr + 1):
    if r not in row2block:
        continue
    bid, br = row2block[r]
    rows.append({
        "row": r, "block": bid, "block_row": br,
        "sample": cell(r, 2),
        "elems": [cell(r, c) for c in range(3, 18)],   # C-Q 15 元素
        "process": cell(r, 18), "we": cell(r, 19), "ref": cell(r, 20),
        "ce": cell(r, 21), "temp": cell(r, 22), "epit": cell(r, 23),
        "stab": cell(r, 24), "info": cell(r, 25), "note": cell(r, 26),
    })

block_rep = {x["block"]: x for x in rows if x["row"] == x["block_row"]}
block_count = defaultdict(int)
for x in rows:
    block_count[x["block"]] += 1

# ---------------- 2. 读影响面底稿（定稿前甲/乙解释，仅作结构化线索） ----------------
impact = {}
with open(CSV_IMPACT, "r", encoding="utf-8-sig", newline="") as f:
    for d in csv.DictReader(f):
        impact[nfkc(d["编号"])] = d

# ---------------- 3. 定稿判级（口径基准_20261107.md 3.1 / 3.2） ----------------
TWO = ("敏化", "时效", "稳定化")          # A3 定义：固溶/退火之后叠加的第二次热处理
ADD_EXCL = ("SLM", "LPBF", "DED", "WAAM", "增材制造", "粉末冶金",
            "焊接", "堆焊", "渗氮", "渗碳", "镀层", "涂层")
NEG_CTX = ("排除", "不收", "不包括", "不是", "非", "只")

NOTES = {
    "2": "题名含 Cast Lean Duplex；本块只录 CF3M 一行，须回 文献/02.pdf 确认所录试样为奥氏体基体（A5 真双相硬排除）。",
    "7": "样本名「304不锈钢(敏化态)」+ 1060℃固溶后 650℃×2h 敏化，按 A3 判 N4 两步。",
    "13": "奥氏体+δ（δ≈1.3%），按 A5 视为奥氏体，须标注 δ 含量。",
    "17": "SiC 超细粉体为外加增强相，工艺为砂型铸造，不是粉末冶金，按 A1 合法（须标注）。",
    "20": "960℃/2h + 1020℃/4h 属锻造加热/两次固溶，按 A3 不属敏化/时效/稳定化，不判两步；另块内铸造+锻轧样本混合、无逐元素成分表。",
    "28": "Z 记「原文另有 FeCl₃ 浸泡，本表只录 3.5% NaCl 极化」，按 3.4 不是介质不符。",
    "30": "介质 3.5%(0.61 mol/L) NaCl 等价写法，按 A2 合规；FeCl₃ 同上不是介质不符。",
    "36": "3.5 wt% NaCl 模拟海水 SME 属人工配方，按 A2 标存疑待补证；固溶→冷轧→退火按 A3 不算两步。",
    "38": "3.56 wt% NaCl (ASTM G61)，按 A2 标存疑待补证；文献27 曾因 3.56% 整篇剔除，维持不回改。",
    "60": "原文 3.5% wt./v（质量体积浓度）非质量分数，按 A2 标存疑待补证。",
    "61": "Z 中「只排除增材/焊接/粉末冶金」为否定性表述，工艺原文 Si-Mo 电阻炉氩气保护熔炼（铸造），按 N1 合法。",
    "62": "900℃×3min 逆相变退火按 A3 不算两步（3.4 关键词线索命中，定稿 A3 下不判两步）。",
    "63": "3.4 关键词线索曾命中「铁素体」，本次按 0.xlsx Y/Z 全文复核；若全文无该字样则不列存疑。",
}


def form_excl(text):
    hits = []
    for w in ADD_EXCL:
        for m in re.finditer(re.escape(w), text, re.I):
            ctx = text[max(0, m.start() - 10):m.start()]
            if any(n in ctx for n in NEG_CTX):
                continue
            hits.append(w)
            break
    return hits


def judge(bid, d, rec):
    bad, sus, mark = [], [], []
    heat_blob = " ".join([
        d.get("热处理原文佐证", ""), d.get("加工工艺原文", ""),
        d.get("样本名示例", ""), d.get("备注原文", ""), rec["process"],
    ])
    wide = "④两步热处理" in (d.get("甲_触发条款", "") + d.get("乙_触发条款", ""))
    hits = [w for w in TWO if w in heat_blob]
    two_step = "是" if (wide and hits) else "否"

    # N4 两步（A3 窄定义）
    if two_step == "是":
        bad.append("N4 两步热处理（%s）" % "/".join(hits))
    elif wide:
        mark.append("底稿宽口径 2+ 步，按 A3 不判两步")
    if "热老化" in heat_blob and two_step == "否":
        sus.append("N4 热老化是否属 A3「时效」未裁定，维持存疑")

    # N3 介质（A2）
    med = d.get("介质", "")
    mtxt = med + " " + rec["note"] + " " + rec["info"] + " " + rec["process"]
    if "3.56" in med:
        sus.append("N3 3.56 wt.% 按 A2 存疑待补证")
    elif "模拟海水" in med or "人工海水" in med or "ASTM D1141" in mtxt or "JIS G 0577" in mtxt:
        sus.append("N3 人工/模拟海水按 A2 存疑待补证")
    elif "wt./v" in mtxt or "wt/v" in mtxt:
        sus.append("N3 3.5% wt./v 非质量分数，按 A2 存疑待补证")
    elif "天然海水" in med:
        mark.append("N3 天然海水按 A2 合规，须标注")
    elif "3.5wt%" in med:
        pass                                    # 含 0.6mol/L 等价写法，A2 合规
    elif "未提及" in med or med == "":
        sus.append("N3 Z/Y 未记介质，须回 文献/%s.pdf 方法段取证" % bid.zfill(2))
    else:
        sus.append("N3 介质写法待核：%s" % med)

    # N2 组织（A5）
    org = d.get("钢种组织", "")
    yz = rec["info"] + " " + rec["note"] + " " + rec["process"]
    if "奥氏体+δ" in org:
        mark.append("N2 奥氏体+δ 按 A5 视为奥氏体，须标注 δ 含量")
    elif "不明" in org:
        sus.append("N2 原文未直写奥氏体组织，须回原文确认")
    elif "双相" in org or "Duplex" in org:
        bad.append("N2 真双相按 A5 硬排除")
    if re.search(r"duplex|双相", yz, re.I):
        sus.append("N2 题名/备注含 Duplex/双相字样，须回原文确认所录试样基体")
    if "铁素体" in yz and "奥氏体+δ" not in org:
        sus.append("N2 备注含铁素体字样，须回原文确认非铁素体为主/真双相")

    # N1 成形（A1 默认熔铸成立，只在原文明确排除项时排除）
    fh = form_excl(yz)
    if fh:
        bad.append("N1 原文明确 %s，按 A1 排除" % "/".join(fh))

    verdict = "不符" if bad else ("存疑" if sus else "合规")
    return verdict, bad, sus, mark, two_step, ("是" if wide else "否")


judge_rows = []
for _, bid in block_rows:
    d = impact.get(bid, {})
    rec = block_rep.get(bid, {})
    if not rec:
        continue
    verdict, bad, sus, mark, two, wide = judge(bid, d, rec)
    judge_rows.append({
        "编号": bid, "条数": block_count[bid],
        "样本名示例": d.get("样本名示例", rec["sample"]),
        "加工工艺原文": d.get("加工工艺原文", rec["process"]),
        "Z备注全文": rec["note"].replace("\n", " ⏎ ")[:400],
        "Y文献信息": rec["info"].replace("\n", " ⏎ ")[:200],
        "介质_底稿": d.get("介质", ""),
        "组织_底稿": d.get("钢种组织", ""),
        "定稿判定": verdict,
        "不符项": "；".join(bad),
        "存疑项": "；".join(sus),
        "标注项": "；".join(mark),
        "定稿两步_N4": two,
        "底稿宽口径2plus": wide,
        "甲_判定": d.get("甲_判定", ""), "乙_判定": d.get("乙_判定", ""),
        "数据质量提示_非口径": d.get("其它风险", ""),
        "人工注记": NOTES.get(bid, ""),
    })

with open(OUT_JUDGE, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(judge_rows[0].keys()))
    w.writeheader()
    w.writerows(judge_rows)

# ---------------- 4. 重复样本检查（多层指纹，只报告不改动） ----------------
E = lambda x: tuple(map(norm, x["elems"]))
FINGERPRINTS = {
    "F0_业务字段全同": lambda x: (x["sample"], E(x), x["process"], norm(x["epit"]),
                                  x["ref"], x["temp"], x["stab"]),
    "F1_样本名+成分+工艺+Epit": lambda x: (x["sample"], E(x), x["process"], norm(x["epit"])),
    "F2_成分+Epit同": lambda x: (E(x), norm(x["epit"])),
    "F3_同块同名": lambda x: (x["block"], x["sample"]),
    "F4_同名+成分（跨块）": lambda x: (x["sample"], E(x)),
    "F5_样本名+Epit同": lambda x: (x["sample"], norm(x["epit"])),
}

dup_rows = []
dup_stat = {}
for name, fn in FINGERPRINTS.items():
    groups = defaultdict(list)
    for x in rows:
        groups[fn(x)].append(x)
    dups = {k: v for k, v in groups.items() if len(v) > 1}
    cross_block = sum(1 for v in dups.values() if len({y["block"] for y in v}) > 1)
    dup_stat[name] = {"组数": len(dups), "涉及行数": sum(len(v) for v in dups.values()),
                      "跨块组数": cross_block}
    for k, v in sorted(dups.items(), key=lambda kv: -len(kv[1])):
        dup_rows.append({
            "指纹": name,
            "组内行数": len(v),
            "涉及块": "/".join(sorted({y["block"] for y in v}, key=lambda s: int(s))),
            "跨块": "是" if len({y["block"] for y in v}) > 1 else "否",
            "行号": ",".join(str(y["row"]) for y in v),
            "样本名": " | ".join(y["sample"] for y in v[:8]),
            "Epit": " | ".join(y["epit"] for y in v[:8]),
            "工艺": " | ".join(y["process"][:60] for y in v[:4]),
            "说明": "同成分不同 Epit/工艺可能是设计网格（如温度×时长），须人工确认，不自动判为录入错误",
        })

with open(OUT_DUP, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(dup_rows[0].keys()) if dup_rows else
                       ["指纹", "组内行数", "涉及块", "跨块", "行号", "样本名", "Epit", "工艺", "说明"])
    w.writeheader()
    w.writerows(dup_rows)

# ---------------- 5. 汇总 ----------------
def dist(key):
    out = defaultdict(lambda: [0, 0])
    for j in judge_rows:
        out[j["定稿判定"]][0] += 1
        out[j["定稿判定"]][1] += j["条数"]
    return dict(out)

stats = {
    "0xlsx": {"行数": maxr, "列数": maxc, "数据条数": len(rows),
              "文献块数": len(block_rows), "缺号": [str(i) for i in range(1, 64)
              if str(i) not in {b for _, b in block_rows}]},
    "判定分布_块数_条数": dist("定稿判定"),
    "不符块": [{"编号": j["编号"], "条数": j["条数"], "不符项": j["不符项"]}
               for j in judge_rows if j["定稿判定"] == "不符"],
    "存疑块": [{"编号": j["编号"], "条数": j["条数"], "存疑项": j["存疑项"]}
               for j in judge_rows if j["定稿判定"] == "存疑"],
    "合规块": [{"编号": j["编号"], "条数": j["条数"], "标注项": j["标注项"]}
               for j in judge_rows if j["定稿判定"] == "合规"],
    "N4定稿两步块": [{"编号": j["编号"], "条数": j["条数"], "不符项": j["不符项"]}
                     for j in judge_rows if j["定稿两步_N4"] == "是"],
    "底稿宽口径2plus_但A3不判两步": [
        {"编号": j["编号"], "条数": j["条数"]}
        for j in judge_rows if j["底稿宽口径2plus"] == "是" and j["定稿两步_N4"] == "否"],
    "标注项块": [{"编号": j["编号"], "标注项": j["标注项"]}
                 for j in judge_rows if j["标注项"]],
    "重复检查": dup_stat,
    "重复组内行数Top10": [{"指纹": r["指纹"], "行数": r["组内行数"],
                          "涉及块": r["涉及块"], "跨块": r["跨块"], "行号": r["行号"]}
                         for r in dup_rows[:10]],
    "闭合校验": {"块数": len(judge_rows), "条数合计": sum(j["条数"] for j in judge_rows),
                 "判定块数合计": sum(v[0] for v in dist("定稿判定").values()),
                 "判定条数合计": sum(v[1] for v in dist("定稿判定").values())},
}
with open(OUT_STAT, "w", encoding="utf-8") as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)

print("=== 0.xlsx 结构 ===", stats["0xlsx"])
print("=== 判定分布（块数/条数）===", stats["判定分布_块数_条数"])
print("=== 闭合校验 ===", stats["闭合校验"])
print("=== 不符块 ===")
for x in stats["不符块"]:
    print("  ", x["编号"], x["条数"], x["不符项"][:110])
print("=== N4 定稿两步块 ===", [x["编号"] for x in stats["N4定稿两步块"]])
print("=== 宽口径2+ 但 A3 不判两步 ===", [x["编号"] for x in stats["底稿宽口径2plus_但A3不判两步"]])
print("=== 合规块 ===", [x["编号"] for x in stats["合规块"]])
print("=== 存疑块数 ===", len(stats["存疑块"]))
print("=== 重复检查 ===")
for k, v in dup_stat.items():
    print("  ", k, v)
print("=== 输出 ===")
print("  ", OUT_JUDGE)
print("  ", OUT_DUP)
print("  ", OUT_STAT)
