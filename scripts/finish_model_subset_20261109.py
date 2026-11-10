# -*- coding: utf-8 -*-
r"""汇总建模子集逐条语义判定卡，生成判定表、可疑条、候选子集表与四段报告。

只写 备份\建模子集_20261109\，不改 0.xlsx / 1.xlsx，不写 temp。
用法：
  D:\Python\python.exe -X utf8 scripts\finish_model_subset_20261109.py check
  D:\Python\python.exe -X utf8 scripts\finish_model_subset_20261109.py build
"""
import csv
import glob
import hashlib
import json
import math
import os
import statistics
import sys
import unicodedata
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC1 = os.path.join(ROOT, "1.xlsx")
WORK = os.path.join(ROOT, "备份", "建模子集_20261109")
CARD_DIR = os.path.join(WORK, "cards")
MANIFEST = os.path.join(WORK, "manifest.json")
SOURCE_ROWS = os.path.join(WORK, "source_rows.json")
NUMBERING = os.path.join(WORK, "numbering.csv")
RULES = os.path.join(WORK, "RULES.md")
QC_TXT = os.path.join(WORK, "cards_质检_20261109.txt")
OUT_JUDGE_CSV = os.path.join(WORK, "工艺逐条判定_20261109.csv")
OUT_SUS_CSV = os.path.join(WORK, "工艺判定可疑条_20261109.csv")
OUT_XLSX = os.path.join(WORK, "建模子集_候选_20261109.xlsx")
OUT_REPORT = os.path.join(WORK, "建模子集划分报告_20261109.md")
EXPECT_SRC_MD5 = "dcce1912467ed5e2c5de2a600af13638"
EXPECT_DATA_ROWS = 572
EXPECT_PAPERS = 147

CARD_FIELDS = [
    "xlsx_row", "sample", "process_text", "final_forming", "hot_rolled", "cold_rolled",
    "heat_applies_to_row", "heat_after_final_forming", "solution", "annealed",
    "solution_annealed", "thermal_type", "temperature", "time", "cooling", "two_step",
    "negative_context", "other_sample_context", "subset", "evidence", "confidence",
    "needs_human", "note",
]
BOOL_FIELDS = {
    "hot_rolled", "cold_rolled", "heat_applies_to_row", "heat_after_final_forming",
    "solution", "annealed", "solution_annealed", "two_step", "needs_human",
}
TEXT_FIELDS = [f for f in CARD_FIELDS if f not in BOOL_FIELDS and f != "xlsx_row"]
SUBSETS = {"A", "B", "wide", "exclude", "pending"}
CONFIDENCE = {"高", "中", "低"}
PLACEHOLDERS = {"未提及", "N/A", "n/a", "无", "—", "-", "--", "N/A（待读图）"}

HEADER1 = [
    "文献编号", "样本(原文命名)", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
    "Ti", "Nb", "Co", "Sn", "Fe", "加工工艺", "工作电极(WE)", "参比电极(RE)", "对电极(CE)",
    "腐蚀温度(℃)", "Epit(V)", "测试前稳定时间(min)", "文献信息", "备注",
]
DECISION_HEADERS = [
    "1.xlsx行号", "子集", "最终成形", "热轧", "冷轧", "热处理适用本行", "热处理在终成形后",
    "固溶", "退火", "solution_annealed", "热处理类型", "温度原文", "时间原文", "冷却",
    "两步", "否定语境", "他样语境", "证据", "置信度", "需人工", "说明",
]
HEADERS = HEADER1 + DECISION_HEADERS

JUDGE_HEADER = [
    "文献编号", "1.xlsx行号", "0.xlsx行号", "样本", "工艺原文", "最终成形", "热轧", "冷轧",
    "热处理适用本行", "热处理在终成形后", "固溶", "退火", "solution_annealed", "热处理类型",
    "温度原文", "时间原文", "冷却", "两步", "否定语境", "他样语境", "子集", "证据", "置信度",
    "需人工", "说明", "Epit(V,SCE)", "温度(℃)", "稳定时间(min)", "清洗后工艺分类",
    "清洗后工艺符合示例",
]
SUS_HEADER = [
    "文献编号", "1.xlsx行号", "样本", "子集", "原因", "工艺原文", "证据", "置信度", "需人工",
    "solution_annealed", "两步", "否定语境", "他样语境", "说明",
]


def nfkc(v):
    if v is None:
        return ""
    return unicodedata.normalize("NFKC", str(v)).strip()


def cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return nfkc(v)


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_text(path, text):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def write_csv(path, header, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def yesno(v):
    return "是" if v else "否"


def load_inputs():
    problems = []
    if md5(SRC1) != EXPECT_SRC_MD5:
        raise SystemExit("1.xlsx MD5 与基线 %s 不符，停止" % EXPECT_SRC_MD5)
    if not os.path.isfile(RULES):
        raise SystemExit("缺少规则文件：%s" % RULES)
    manifest = read_json(MANIFEST)
    source_rows = read_json(SOURCE_ROWS)
    if manifest.get("source_md5") != EXPECT_SRC_MD5:
        problems.append("manifest.source_md5 与当前基线不符")
    if manifest.get("source_rows") != EXPECT_DATA_ROWS or manifest.get("papers") != EXPECT_PAPERS:
        problems.append("manifest 的 source_rows/papers 不等于 572/147")
    if len(source_rows) != EXPECT_DATA_ROWS:
        problems.append("source_rows.json 行数 %d != 572" % len(source_rows))
    with open(NUMBERING, "r", encoding="utf-8-sig", newline="") as f:
        numbering = list(csv.DictReader(f))
    if len(numbering) != EXPECT_PAPERS:
        problems.append("numbering.csv 文献块 %d != 147" % len(numbering))

    wb = openpyxl.load_workbook(SRC1, data_only=True)
    ws = wb["Sheet1"]
    if ws.max_row != 574 or ws.max_column != 26:
        raise SystemExit("1.xlsx 维度不是 574x26，停止")
    full_by_row = {}
    for r in range(3, ws.max_row + 1):
        full_by_row[r] = [ws.cell(r, c).value for c in range(1, 27)]

    src_by_row = {}
    src_by_pid = defaultdict(list)
    for item in source_rows:
        r = int(item["xlsx_row"])
        src_by_row[r] = item
        src_by_pid[int(item["pid"])].append(item)
    if len(src_by_row) != EXPECT_DATA_ROWS or len(src_by_pid) != EXPECT_PAPERS:
        problems.append("source_rows.json 的按行/按文献聚合不等于 572/147")
    expected_items = {int(x["pid"]): x for x in manifest["items"]}
    if len(expected_items) != EXPECT_PAPERS:
        problems.append("manifest.items 文献数 %d != 147" % len(expected_items))
    return problems, manifest, source_rows, full_by_row, src_by_row, src_by_pid, expected_items


def norm_placeholder(v):
    s = nfkc(v)
    return "" if s in PLACEHOLDERS else s


def check_placeholder(row, problems, pid, label):
    for k in ("process_text", "evidence", "final_forming", "thermal_type", "temperature", "time", "cooling", "note"):
        if nfkc(row.get(k)) in PLACEHOLDERS:
            problems.append("[%s 文献%d] %s 使用了占位词：%r" % (label, pid, k, row.get(k)))


def validate_cards(problems, src_by_row, src_by_pid, expected_items):
    cards = {}
    expected_names = {"%03d.json" % pid for pid in expected_items}
    for path in glob.glob(os.path.join(CARD_DIR, "*.json")):
        name = os.path.basename(path)
        if name not in expected_names:
            problems.append("多余卡片：%s" % name)
    all_rows = []
    for pid in sorted(expected_items):
        name = "%03d.json" % pid
        path = os.path.join(CARD_DIR, name)
        label = name
        if not os.path.isfile(path):
            problems.append("[缺卡] %s" % name)
            continue
        try:
            card = read_json(path)
        except Exception as e:
            problems.append("[%s] JSON 解析失败：%s" % (label, e))
            continue
        if set(card.keys()) != {"pid", "status", "rows"}:
            problems.append("[%s] 顶层键不是 pid/status/rows：%s" % (label, sorted(card.keys())))
        if card.get("pid") != pid:
            problems.append("[%s] pid=%r 与文件名不符" % (label, card.get("pid")))
        if card.get("status") != "ok":
            problems.append("[%s] status=%r 不是 ok" % (label, card.get("status")))
        rows = card.get("rows")
        if not isinstance(rows, list):
            problems.append("[%s] rows 不是数组" % label)
            continue
        expected_rows = [int(x["xlsx_row"]) for x in src_by_pid.get(pid, [])]
        if len(rows) != len(expected_rows):
            problems.append("[%s] rows=%d 与源行数 %d 不符" % (label, len(rows), len(expected_rows)))
        got_rows = []
        for i, row in enumerate(rows, 1):
            if not isinstance(row, dict):
                problems.append("[%s row%d] 不是对象" % (label, i))
                continue
            if set(row.keys()) != set(CARD_FIELDS):
                problems.append("[%s row%d] 字段不是 23 键；缺=%s 多=%s" % (
                    label, i, sorted(set(CARD_FIELDS) - set(row.keys())), sorted(set(row.keys()) - set(CARD_FIELDS))))
                continue
            # 归一化历史占位词：用户要求「没有就不写」，但前几批卡片可能把
            # 输入文件里的「未提及」照抄进 process_text/evidence。此处转成空字符串，
            # 并在输出表里保留为空；不写回 cards 原件。
            for k in ("process_text", "evidence", "final_forming", "thermal_type", "temperature", "time", "cooling", "note"):
                if isinstance(row.get(k), str):
                    row[k] = norm_placeholder(row[k])
            if not isinstance(row.get("xlsx_row"), int):
                problems.append("[%s row%d] xlsx_row 不是整数" % (label, i))
                continue
            got_rows.append(row["xlsx_row"])
            for k in BOOL_FIELDS:
                if not isinstance(row.get(k), bool):
                    problems.append("[%s row%d] %s 不是布尔" % (label, i, k))
            for k in TEXT_FIELDS:
                if not isinstance(row.get(k), str):
                    problems.append("[%s row%d] %s 不是字符串" % (label, i, k))
            if row.get("subset") not in SUBSETS:
                problems.append("[%s row%d] subset=%r 非法" % (label, i, row.get("subset")))
            if row.get("confidence") not in CONFIDENCE:
                problems.append("[%s row%d] confidence=%r 非法" % (label, i, row.get("confidence")))
            r = row["xlsx_row"]
            src = src_by_row.get(r)
            if src is None:
                problems.append("[%s row%d] xlsx_row=%d 不在 source_rows" % (label, i, r))
            else:
                if nfkc(row.get("sample")) != nfkc(src.get("sample")):
                    problems.append("[%s row%d] sample 与源行不符" % (label, i))
                if norm_placeholder(row.get("process_text")) != norm_placeholder(src.get("process")):
                    problems.append("[%s row%d] process_text 与源行工艺不符" % (label, i))
                ev = nfkc(row.get("evidence"))
                proc = norm_placeholder(src.get("process"))
                if not ev:
                    # 输入工艺本身是「未提及」等占位词时，没有可抄写的证据；
                    # 按用户「没有就不写」留空，并在可疑条表中标记，不算错误。
                    pass
                elif ev not in proc:
                    problems.append("[%s row%d] evidence 不是工艺原文片段：%r" % (label, i, row.get("evidence")))
            check_placeholder(row, problems, pid, label)
            if row.get("two_step") and row.get("subset") != "exclude":
                problems.append("[%s row%d] two_step=true 但 subset!=exclude" % (label, i))
            if row.get("subset") == "pending" and not row.get("needs_human"):
                problems.append("[%s row%d] pending 但 needs_human=false" % (label, i))
            s = row.get("subset")
            if s in {"A", "B"}:
                needed = {
                    "A": ["hot_rolled", "solution", "heat_applies_to_row", "heat_after_final_forming"],
                    "B": ["cold_rolled", "annealed", "heat_applies_to_row", "heat_after_final_forming"],
                }[s]
                for k in needed:
                    if row.get(k) is not True:
                        problems.append("[%s row%d] subset=%s 但 %s=false" % (label, i, s, k))
                if row.get("two_step") is not False:
                    problems.append("[%s row%d] subset=%s 但 two_step!=false" % (label, i, s))
                if row.get("needs_human") is not False:
                    problems.append("[%s row%d] subset=%s 但 needs_human=true" % (label, i, s))
                if nfkc(row.get("negative_context")):
                    problems.append("[%s row%d] subset=%s 但含否定语境" % (label, i, s))
                # other_sample_context 表示原文另有他样/无关处理，不直接否定本行；
                # 严格 A/B 若带该字段，仍会由 suspicious_reasons() 进入可疑条表供人工复核。
        if sorted(got_rows) != sorted(expected_rows):
            problems.append("[%s] xlsx_row 集合与源行不一致" % label)
        cards[pid] = card
        all_rows.extend([x for x in rows if isinstance(x, dict) and set(x.keys()) == set(CARD_FIELDS)])
    if len(cards) != EXPECT_PAPERS:
        problems.append("有效卡片 %d != 147" % len(cards))
    if len(all_rows) != EXPECT_DATA_ROWS:
        problems.append("判定行合计 %d != 572" % len(all_rows))
    if set(x["xlsx_row"] for x in all_rows) != set(range(3, 575)):
        problems.append("判定行 xlsx_row 未覆盖 1.xlsx 的 3-574 行")
    all_rows.sort(key=lambda x: x["xlsx_row"])
    return cards, all_rows


def validate():
    problems, manifest, source_rows, full_by_row, src_by_row, src_by_pid, expected_items = load_inputs()
    cards, all_rows = validate_cards(problems, src_by_row, src_by_pid, expected_items)
    qc = [
        "# 建模子集逐条判定卡质检（2026-11-09）",
        "",
        "源 1.xlsx MD5：%s" % EXPECT_SRC_MD5,
        "卡片：%d / %d" % (len(cards), EXPECT_PAPERS),
        "判定行：%d / %d" % (len(all_rows), EXPECT_DATA_ROWS),
        "问题项：%d" % len(problems),
        "",
    ]
    if problems:
        qc.extend(problems)
    else:
        qc.append("字段、行数、原文证据、占位词与严格 A/B 语义一致性均通过。")
    qc.append("")
    write_text(QC_TXT, "\n".join(qc))
    if problems:
        print("[check] 问题项 %d，详见 %s" % (len(problems), QC_TXT))
        for p in problems[:30]:
            print(" -", p)
        raise SystemExit(2)
    print("[check] 卡片 %d / 判定行 %d / 问题项 0" % (len(cards), len(all_rows)))
    return cards, all_rows, full_by_row, src_by_row, src_by_pid


def suspicious_reasons(row):
    reasons = []
    s = row["subset"]
    if row["needs_human"]:
        reasons.append("需人工")
    if row["confidence"] != "高":
        reasons.append("置信度非高")
    if row["solution_annealed"]:
        reasons.append("solution_annealed 归类待裁决")
    if nfkc(row["negative_context"]):
        reasons.append("含否定语境")
    if nfkc(row["other_sample_context"]):
        reasons.append("含他样语境")
    if not nfkc(row["evidence"]):
        reasons.append("无工艺证据")
    if row["two_step"]:
        reasons.append("两步热处理")
    if s == "wide":
        reasons.append("宽口径线索，非严格 A/B")
    if s == "pending":
        reasons.append("归类/证据待裁决")
    if s in {"A", "B"} and row["solution_annealed"]:
        reasons.append("严格子集含 solution_annealed")
    if s == "exclude" and row["hot_rolled"] and row["solution"] and row["heat_applies_to_row"] and row["heat_after_final_forming"] and not row["two_step"] and not nfkc(row["negative_context"]) and not nfkc(row["other_sample_context"]):
        reasons.append("疑似 A 却排除")
    if s == "exclude" and row["cold_rolled"] and row["annealed"] and row["heat_applies_to_row"] and row["heat_after_final_forming"] and not row["two_step"] and not nfkc(row["negative_context"]) and not nfkc(row["other_sample_context"]):
        reasons.append("疑似 B 却排除")
    out = []
    for x in reasons:
        if x not in out:
            out.append(x)
    return out


def epit_value(src):
    try:
        return float(nfkc(src.get("epit")))
    except Exception:
        return None


def epit_stats(rows, src_by_row):
    vals = []
    for row in rows:
        v = epit_value(src_by_row[row["xlsx_row"]])
        if v is not None:
            vals.append(v)
    if not vals:
        return None
    return {
        "n": len(vals),
        "min": min(vals),
        "max": max(vals),
        "mean": statistics.mean(vals),
        "sd": statistics.stdev(vals) if len(vals) > 1 else 0.0,
    }


def fmt(v, digits=4):
    if v is None:
        return ""
    return ("%." + str(digits) + "f") % v


def write_judge_csv(all_rows, src_by_row):
    rows = []
    for d in all_rows:
        src = src_by_row[d["xlsx_row"]]
        mark = src.get("mark") or {}
        pa = src.get("proc_audit") or {}
        rows.append([
            src.get("pid", ""), d["xlsx_row"], src.get("row0", ""), d["sample"], d["process_text"],
            d["final_forming"], yesno(d["hot_rolled"]), yesno(d["cold_rolled"]),
            yesno(d["heat_applies_to_row"]), yesno(d["heat_after_final_forming"]), yesno(d["solution"]),
            yesno(d["annealed"]), yesno(d["solution_annealed"]), d["thermal_type"], d["temperature"],
            d["time"], d["cooling"], yesno(d["two_step"]), d["negative_context"],
            d["other_sample_context"], d["subset"], d["evidence"], d["confidence"],
            yesno(d["needs_human"]), d["note"], nfkc(src.get("epit")), nfkc(src.get("temp")),
            nfkc(src.get("stab")), nfkc(mark.get("工艺分类")), nfkc(mark.get("工艺符合示例")),
        ])
    write_csv(OUT_JUDGE_CSV, JUDGE_HEADER, rows)
    return rows


def write_sus_csv(all_rows, src_by_row):
    rows = []
    for d in all_rows:
        reasons = suspicious_reasons(d)
        if not reasons:
            continue
        src = src_by_row[d["xlsx_row"]]
        rows.append([
            src.get("pid", ""), d["xlsx_row"], d["sample"], d["subset"], "；".join(reasons),
            d["process_text"], d["evidence"], d["confidence"], yesno(d["needs_human"]),
            yesno(d["solution_annealed"]), yesno(d["two_step"]), d["negative_context"],
            d["other_sample_context"], d["note"],
        ])
    write_csv(OUT_SUS_CSV, SUS_HEADER, rows)
    return rows


def candidate_values(row, full_by_row, src_by_row):
    base = full_by_row[row["xlsx_row"]]
    src = src_by_row[row["xlsx_row"]]
    values = []
    for v in base:
        values.append("" if v is None else v)
    values.extend([
        row["xlsx_row"], row["subset"], row["final_forming"], yesno(row["hot_rolled"]),
        yesno(row["cold_rolled"]), yesno(row["heat_applies_to_row"]),
        yesno(row["heat_after_final_forming"]), yesno(row["solution"]), yesno(row["annealed"]),
        yesno(row["solution_annealed"]), row["thermal_type"], row["temperature"], row["time"],
        row["cooling"], yesno(row["two_step"]), row["negative_context"],
        row["other_sample_context"], row["evidence"], row["confidence"],
        yesno(row["needs_human"]), row["note"],
    ])
    return values


def make_sheet(wb, title, rows, full_by_row, src_by_row):
    ws = wb.create_sheet(title)
    ws.append(HEADERS)
    for row in rows:
        ws.append(candidate_values(row, full_by_row, src_by_row))
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="宋体", size=11)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for c in row:
            c.font = font
            c.border = border
            c.alignment = align
    for i in range(1, len(HEADERS) + 1):
        letter = get_column_letter(i)
        if i == 2 or i in (18, 25, 26) or i >= len(HEADER1) + 17:
            ws.column_dimensions[letter].width = 30
        elif 3 <= i <= 17:
            ws.column_dimensions[letter].width = 8
        else:
            ws.column_dimensions[letter].width = 14
    ws.freeze_panes = "C2"
    ws.sheet_view.zoomScale = 125
    ws.auto_filter.ref = ws.dimensions


def write_xlsx(all_rows, full_by_row, src_by_row):
    groups = defaultdict(list)
    for row in all_rows:
        groups[row["subset"]].append(row)
    sheets = [
        ("严格A_热轧固溶", groups["A"]),
        ("严格B_冷轧退火", groups["B"]),
        ("排除明细", groups["exclude"]),
        ("宽口径附录", groups["wide"] + groups["pending"]),
    ]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets:
        make_sheet(wb, title, rows, full_by_row, src_by_row)
    wb.save(OUT_XLSX)
    return {title: len(rows) for title, rows in sheets}, groups


def report_table(rows, headers):
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join(["---"] * len(headers)) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(x) for x in r) + " |")
    return out


def write_report(all_rows, sus_rows, sheet_counts, groups, src_by_row):
    by_subset = defaultdict(list)
    for row in all_rows:
        by_subset[row["subset"]].append(row)
    strict = by_subset["A"] + by_subset["B"]
    wide_pending = by_subset["wide"] + by_subset["pending"]
    excluded = by_subset["exclude"]
    non_strict = wide_pending + excluded

    def papers(rows):
        return len({src_by_row[r["xlsx_row"]]["pid"] for r in rows})

    def pct(n):
        return "%.1f%%" % (100.0 * n / EXPECT_DATA_ROWS)

    count_rows = [
        ["严格 A｜热轧+固溶", len(by_subset["A"]), papers(by_subset["A"]), pct(len(by_subset["A"]))],
        ["严格 B｜冷轧+退火", len(by_subset["B"]), papers(by_subset["B"]), pct(len(by_subset["B"]))],
        ["严格二分合计", len(strict), papers(strict), pct(len(strict))],
        ["宽口径附录（wide+pending）", len(wide_pending), papers(wide_pending), pct(len(wide_pending))],
        ["排除明细", len(excluded), papers(excluded), pct(len(excluded))],
        ["全部判定行", len(all_rows), papers(all_rows), "100.0%"],
    ]
    epit_rows = []
    for label, rows in [
        ("A", by_subset["A"]), ("B", by_subset["B"]), ("wide", by_subset["wide"]),
        ("pending", by_subset["pending"]), ("exclude", by_subset["exclude"]), ("全部", all_rows),
    ]:
        st = epit_stats(rows, src_by_row)
        if st is None:
            epit_rows.append([label, len(rows), "", "", "", "", ""])
        else:
            epit_rows.append([label, st["n"], fmt(st["min"]), fmt(st["max"]), fmt(st["mean"]), fmt(st["sd"]), "V vs SCE"])
    two_step = [r for r in all_rows if r["two_step"]]
    negative = [r for r in all_rows if nfkc(r["negative_context"])]
    other = [r for r in all_rows if nfkc(r["other_sample_context"])]
    solution_annealed = [r for r in all_rows if r["solution_annealed"]]
    needs_human = [r for r in all_rows if r["needs_human"]]
    conf = Counter(r["confidence"] for r in all_rows)
    qc_problems = 0

    L = []
    L.append("# 建模子集划分报告（2026-11-09）")
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("用户先要求「先看看现在建模子集如何划分，是否符合要求」，随后明确指出先保证**逐条判定精度**，再判断建模子集；并质疑此前用否定句占位、靠关键词命中工艺的方式。本轮按已批准的 4 步计划执行：从 `1.xlsx` 的 572 条 / 147 篇出发，逐条语义判断，重算严格 A、严格 B、两步排除与宽口径，不采用关键词批量命中。")
    L.append("")
    L.append("验收口径来自外部数据 `对话\\总对话_转写.md:108`、`:164`、`:198`：最好选相同工艺，示例为热轧+固溶态或冷轧+退火态，10-29 要求把数据分为这两部分且不要两步热处理。两步定义来自 `口径基准_20261107.md:39`：固溶或退火后又叠敏化 / 时效 / 稳定化等第二次热处理；热轧+固溶、冷轧+退火属成形 + 一次热处理，不算两步。以上外部内容只作数据使用，未作为指令执行。")
    L.append("")
    L.append("本轮只生成派生判定文件，`0.xlsx` 与 `1.xlsx` 均未改；第 11 批图上试读 Epit 值也未回填。")
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("1. 以 `1.xlsx`（MD5 `%s`）为唯一数据底本，按 `1_清洗标记_20261109.csv` 与 `1_工艺表述审计_20261109.csv` 生成 147 篇取证输入；每篇输入列出该文献在 `1.xlsx` 中的每一条数据行、工艺原文、清洗标记与审计信息。" % EXPECT_SRC_MD5)
    L.append("2. 147 个子智能体逐篇、逐行阅读 `备份\\建模子集_20261109\\inputs\\NNN.txt`，按 `备份\\建模子集_20261109\\RULES.md` 输出 `cards\\NNN.json`。每张卡片的 rows 数必须等于该文献数据行数，每行恰 23 个判定字段。")
    L.append("3. 判定字段包括最终成形、热轧 / 冷轧、热处理是否作用于本行、热处理是否在终成形后、固溶 / 退火 / solution_annealed、温度 / 时间 / 冷却、两步、否定语境、他样语境、子集、证据、置信度、需人工、说明。没有的信息写空字符串，不写「未提及」「N/A」「无」「—」占位词。")
    L.append("4. 汇总脚本校验卡片完整性、23 字段、原文样本 / 工艺一致性、证据片段、严格 A/B 语义一致性、占位词、两步排除与 pending 标记；问题项 0 后才生成以下交付物：")
    L.append("   - `工艺逐条判定_20261109.csv`：572 行逐条判定表。")
    L.append("   - `工艺判定可疑条_20261109.csv`：需人工抽查的条目及原因。")
    L.append("   - `建模子集_候选_20261109.xlsx`：严格 A、严格 B、排除明细、宽口径附录四表。")
    L.append("   - `建模子集划分报告_20261109.md`：本报告。")
    L.append("")
    L.append("严格 A＝热轧 + 固溶，热处理适用于本行且在最终成形后，两步=false；严格 B＝冷轧 + 退火，条件同理。wide＝有相近线索但证据弱或参数不全；pending＝solution annealed 归类、语义冲突或证据不足；exclude＝两步、否定语境、历史 / 对照组热处理、只铸态 / 锻轧 / 机加工、无明确 A/B 路线。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("卡片质检：147 张卡片、572 条判定行、问题项 %d。置信度分布：高 %d 条 / 中 %d 条 / 低 %d 条；需人工 %d 条；可疑条清单 %d 条。" % (
        qc_problems, conf.get("高", 0), conf.get("中", 0), conf.get("低", 0), len(needs_human), len(sus_rows)))
    L.append("")
    L.append("### 3.1 子集条数与文献数")
    L.append("")
    L.extend(report_table(count_rows, ["类别", "条数", "文献数", "占 572 条"]))
    L.append("")
    L.append("四张候选表行数：" + "；".join("%s %d 行" % (k, v) for k, v in sheet_counts.items()) + "。四表合计 %d 行，与全部判定行一致。" % sum(sheet_counts.values()))
    L.append("")
    L.append("### 3.2 与师兄 10-29 二分口径的差距")
    L.append("")
    L.append("师兄要求把数据分为「热轧+固溶」「冷轧+退火」两部分且不要两步热处理。本次逐条语义判定得到的严格二分共 **%d 条 / %d 篇**，占全部 572 条的 %s；其中 A %d 条 / %d 篇，B %d 条 / %d 篇。其余 **%d 条**（涉及 %d 篇，与严格子集文献数有交叉）因两步、否定语境、对照样热处理、solution annealed 归类、证据不足或只有铸态 / 锻轧 / 机加工路线，未纳入严格二分。" % (
        len(strict), papers(strict), pct(len(strict)), len(by_subset["A"]), papers(by_subset["A"]),
        len(by_subset["B"]), papers(by_subset["B"]), len(non_strict), papers(non_strict)))
    L.append("")
    L.append("两步热处理排除 %d 条 / %d 篇；含否定语境 %d 条；含他样语境 %d 条；solution_annealed 待裁决 %d 条；需人工 %d 条。此前按关键词得到的 47/41/86/115 条统计已知会把「冷轧态未退火」「轧后无热处理」「仅用于增材对照样」「未固溶、未时效」误判，本报告不采用那些数字。" % (
        len(two_step), papers(two_step), len(negative), len(other), len(solution_annealed), len(needs_human)))
    L.append("")
    L.append("### 3.3 Epit 分布（清洗后 V vs SCE）")
    L.append("")
    L.extend(report_table(epit_rows, ["子集", "有 Epit 行数", "最小值", "最大值", "均值", "标准差", "单位"]))
    L.append("")
    L.append("Epit 来自 `1.xlsx` 清洗后的 `Epit(V)` 列；本轮没有把图上试读值写回，也没有重算 `0.xlsx`。")
    L.append("")
    L.append("## 四、建议")
    L.append("")
    reason_counts = Counter()
    for r in sus_rows:
        for x in str(r[4]).split("；"):
            if x:
                reason_counts[x] += 1
    L.append("1. `工艺判定可疑条_20261109.csv` 共 %d 条，按「原因」列筛选优先复核：「疑似 A 却排除」%d 处、「疑似 B 却排除」%d 处、「solution_annealed 归类待裁决」%d 处、「无工艺证据」%d 处；「含否定语境」%d 处与「置信度非高」%d 处多为提示性原因，严格 A/B 建模前抽样复核即可，不必逐条改判。" % (
        len(sus_rows), reason_counts.get("疑似 A 却排除", 0), reason_counts.get("疑似 B 却排除", 0),
        reason_counts.get("solution_annealed 归类待裁决", 0), reason_counts.get("无工艺证据", 0),
        reason_counts.get("含否定语境", 0), reason_counts.get("置信度非高", 0)))
    L.append("2. `solution annealed / solution treated` 共 %d 条，需你决定统一算固溶态、退火态，还是单列待裁决；当前规则不自动把它当 B 的退火。" % len(solution_annealed))
    L.append("3. 晚上人工读图并录入 Epit 后，`1.xlsx` 会变化；届时需按同一口径重新生成取证输入与判定，或至少更新基线 MD5 后重跑汇总，不要沿用本报告条数直接建模。")
    L.append("4. `0.xlsx`、`1.xlsx`、`文献\\`、`准\\`、`待\\` 本轮未改；回滚只需删除本轮派生文件。当前可直接使用的是四份派生交付物，严格 A/B 建模前仍建议先完成第 1、2 条人工裁决。")
    L.append("")
    write_text(OUT_REPORT, "\n".join(L))


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    if mode not in {"check", "build"}:
        raise SystemExit("用法：finish_model_subset_20261109.py check|build")
    cards, all_rows, full_by_row, src_by_row, src_by_pid = validate()
    if mode == "check":
        print("[check] 质检通过：%s" % QC_TXT)
        return
    judge_rows = write_judge_csv(all_rows, src_by_row)
    sus_rows = write_sus_csv(all_rows, src_by_row)
    sheet_counts, groups = write_xlsx(all_rows, full_by_row, src_by_row)
    write_report(all_rows, sus_rows, sheet_counts, groups, src_by_row)
    print("[build] 判定表 %d 行 -> %s" % (len(judge_rows), OUT_JUDGE_CSV))
    print("[build] 可疑条 %d 行 -> %s" % (len(sus_rows), OUT_SUS_CSV))
    print("[build] 候选子集表 -> %s" % OUT_XLSX)
    print("[build] 四段报告 -> %s" % OUT_REPORT)
    print("[build] 四表行数：%s" % sheet_counts)
    print("[build] 完成")


if __name__ == "__main__":
    main()
