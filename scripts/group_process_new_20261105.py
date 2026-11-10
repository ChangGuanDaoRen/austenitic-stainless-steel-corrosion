# -*- coding: utf-8 -*-
"""按师兄 2026-10-29 20:29 口径重做工艺分组（1.xlsx 现只有 Sheet1）。

口径原文（`对话\总对话_转写.md` 第 198 行）：
  「把里面的数据分为两部分，一部分是热轧+固溶的，另一部分是冷轧+退火的，
    不要两步热处理的，因为敏化是研究焊接的，时效一般来说不用在奥氏体不锈钢，
    稳定化应该也不需要」

分类规则（逐条落表、可回溯）：
  A  热轧+固溶       最后一成形步=热轧，唯一热处理=固溶
  B  冷轧+退火       最后一成形步=冷轧，唯一热处理=退火
  X1 排除            ≥2 步且含 时效/稳定化/敏化/热老化/未注明保温（师兄点名不要）
  X2 排除待判        ≥2 步但同为固溶+固溶 / 退火+退火 / 固溶+退火（同族，中间夹热加工）
  O  其他            热处理 ≤1 步但路线不属 A/B（铸态+固溶、冷轧+固溶、热轧+退火、
                     无热处理、成形未注明…）
  「固溶退火」按单步固溶计（solution annealing）；均热/均匀化视为前置加热，不计步。

用法：python scripts/group_process_new_20261105.py [diag|write]
"""
import re
import sys
import collections

import openpyxl

SRC = "1.xlsx"
DST = "1_工艺分组_新口径_20261105.xlsx"

FORM_CAST = ("铸态", "铸造", "熔模", "离心", "砂型", "真空精密铸造", "浇注", "连铸", "铸锭", "铸坯")
HT_MAIN = ("固溶", "退火", "时效", "稳定化", "敏化", "老化")
PRE_HEAT = ("均热", "均匀化", "保温", "预热", "加热", "升温")

CLS_NAME = {
    "A": "A 热轧+固溶",
    "B": "B 冷轧+退火",
    "X1": "X1 排除（时效/稳定化/敏化/热老化）",
    "X2": "X2 排除待判（同族两步）",
    "O": "O 其他（路线不符/无记录）",
}
ORDER = ["A", "B", "X1", "X2", "O"]


def split_chain(text):
    t = str(text or "")
    t = re.sub(r"[（(][^）)]*[）)]", "", t)          # 括号内是注解，不参与判定
    t = t.replace("；", "+").replace(";", "+")
    t = t.replace("，", "+").replace(",", "+")
    parts = re.split(r"[+→]|—>", t)
    return [p.strip() for p in parts if p.strip()]


def tag_segment(seg):
    """返回 (kind, label)。kind ∈ FORM / HT / PRE / OTHER"""
    s = seg
    if "固溶退火" in s or ("固溶" in s and "退火" in s):
        return ("HT", "固溶")
    for k in HT_MAIN:
        if k in s:
            return ("HT", "热老化" if k == "老化" else k)
    if "保温" in s and "℃" in s:
        return ("HT", "保温/未注明")
    if re.search(r"冷轧|温轧|深冷轧|冷变形|冷作|预变形|预应变|冷拔", s):
        return ("FORM", "冷轧")
    if "热轧" in s or "轧至" in s:
        return ("FORM", "热轧")
    if "锻" in s:
        return ("FORM", "锻造")
    if any(k in s for k in FORM_CAST):
        return ("FORM", "铸造")
    if "轧制" in s:
        return ("FORM", "轧制(未注明冷热)")
    for k in PRE_HEAT:
        if k in s and "℃" in s:
            return ("PRE", "均热/加热")
    if "℃" in s and re.search(r"\d\s*(h|min|小时|分钟)", s):
        return ("HT", "未注明热处理")
    return ("OTHER", "")


def classify(text):
    segs = split_chain(text)
    kinds = [tag_segment(s) for s in segs]
    last_form = max((i for i, (k, _) in enumerate(kinds) if k == "FORM"), default=-1)
    form, hts, ht_idx = None, [], []
    for i, (kind, label) in enumerate(kinds):
        if kind == "FORM":
            form = label
        elif kind == "HT":
            if label == "未注明热处理" and i < last_form:
                continue
            hts.append(label)
            ht_idx.append(i)
    n = len(hts)
    # 末道组合：最后一次成形之后的那次热处理
    last_form_idx = max((i for i, (k, _) in enumerate(kinds) if k == "FORM"), default=-1)
    tail_ht = None
    if ht_idx and ht_idx[-1] > last_form_idx:
        tail_ht = hts[-1]
    match = None
    if tail_ht == "固溶" and form == "热轧":
        match = "A"
    elif tail_ht == "退火" and form == "冷轧":
        match = "B"

    if n >= 2:
        if {"时效", "稳定化", "敏化", "热老化", "保温/未注明"} & set(hts):
            return "X1", form, hts, "含后处理：" + "+".join(hts), None
        tag = ("（若按末道口径 → %s）" % match) if match else ""
        return "X2", form, hts, "同族两步：" + "+".join(hts) + tag, match
    if match:
        return match, form, hts, "", None
    if n == 1:
        return "O", form, hts, "单步%s，但成形=%s" % (hts[0], form or "未注明"), None
    if form == "铸造":
        return "O", form, hts, "铸态直接用，无热处理", None
    if form:
        return "O", form, hts, "%s 后无热处理记录" % form, None
    return "O", form, hts, "无成形记录", None


def load_rows():
    ws = openpyxl.load_workbook(SRC, data_only=True)["Sheet1"]
    out, cur = [], None
    for r in range(3, ws.max_row + 1):
        p = ws.cell(row=r, column=1).value
        if p not in (None, ""):
            cur = p
        out.append({
            "row": r, "paper": cur,
            "sample": ws.cell(row=r, column=2).value,
            "proc": ws.cell(row=r, column=18).value,
            "epit": ws.cell(row=r, column=23).value,
        })
    return out


def stats(rs):
    ep = [x["epit"] for x in rs if isinstance(x["epit"], (int, float))]
    papers = sorted({int(x["paper"]) for x in rs})
    mean, sd = "—", ""
    if ep:
        m = sum(ep) / len(ep)
        s = (sum((v - m) ** 2 for v in ep) / len(ep)) ** 0.5 if len(ep) > 1 else 0
        mean, sd = "%.3f" % m, "%.3f" % s
    rng = ("%.3f~%.3f" % (min(ep), max(ep))) if ep else "—"
    return len(rs), len(papers), papers, mean, sd, rng


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "diag"
    rows = load_rows()
    for x in rows:
        x["cls"], x["form"], x["hts"], x["why"], x["match"] = classify(x["proc"])

    groups = collections.OrderedDict((k, [x for x in rows if x["cls"] == k]) for k in ORDER)

    print("总 %d 条 / %d 篇\n" % (len(rows), len({int(x['paper']) for x in rows})))
    for k in ORDER:
        n, np_, papers, mean, sd, rng = stats(groups[k])
        print("%-34s %3d 条 / %2d 篇 | Epit %s±%s | %s" % (CLS_NAME[k], n, np_, mean, sd, rng))
        print("    文献: " + ",".join(map(str, papers)))

    loose_a = groups["A"] + [x for x in groups["X2"] if x["match"] == "A"]
    loose_b = groups["B"] + [x for x in groups["X2"] if x["match"] == "B"]
    print("\n【两种口径对比】")
    print("  严格（全程只 1 步热处理）：A %d 条 / B %d 条 → 合计 %d 条" %
          (len(groups["A"]), len(groups["B"]), len(groups["A"]) + len(groups["B"])))
    print("  按末道（末道成形+末道热处理对上即可）：A %d 条 / B %d 条 → 合计 %d 条" %
          (len(loose_a), len(loose_b), len(loose_a) + len(loose_b)))

    print("\n--- O 组按「成形+热处理」再分 ---")
    sub = collections.Counter()
    for x in groups["O"]:
        key = "%s + %s" % (x["form"] or "成形未注明",
                           "+".join(x["hts"]) if x["hts"] else "无热处理")
        sub[key] += 1
    for key, c in sub.most_common():
        print("   %-28s %3d 条" % (key, c))

    print("\n--- X1 / X2 明细 ---")
    for k in ("X1", "X2"):
        for (p, why), c in sorted(collections.Counter(
                (x["paper"], x["why"]) for x in groups[k]).items(), key=lambda z: -z[1]):
            print("   [%s] 文献%-2s %-30s %d 条" % (k, p, why, c))

    if mode != "write":
        print("\n[diag] 未写文件")
        return

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "分组统计"
    ws.append(["类别", "条数", "文献数", "文献编号（条数）", "Epit 均值(V)", "sd", "Epit 范围(V)", "说明"])
    notes = {
        "A": "师兄 10-29：末道热轧 + 固溶，全程仅 1 步热处理",
        "B": "师兄 10-29：末道冷轧 + 退火，全程仅 1 步热处理",
        "X1": "师兄明确不要：含敏化/时效/稳定化/热老化",
        "X2": "同族两步（如固溶→冷轧→退火）；说明列标了按末道口径应归哪类",
        "O": "单步但成形不符，或热处理缺失",
    }
    for k in ORDER:
        n, np_, papers, mean, sd, rng = stats(groups[k])
        ws.append([CLS_NAME[k], n, np_,
                   "; ".join("%d×%d" % (p, groups[k] and
                                        sum(1 for x in groups[k] if int(x["paper"]) == p))
                             for p in papers),
                   mean, sd, rng, notes[k]])
    ws.append([])
    ws.append(["严格口径 A+B（全程 1 步热处理）",
               len(groups["A"]) + len(groups["B"]),
               len({int(x["paper"]) for x in groups["A"] + groups["B"]}),
               "", "", "", "", "可直接建模"])
    ws.append(["末道口径 A+B（末道对上即可，含 X2 可归入者）",
               len(loose_a) + len(loose_b),
               len({int(x["paper"]) for x in loose_a + loose_b}),
               "", "", "", "", "A %d 条 + B %d 条" % (len(loose_a), len(loose_b))])
    ws.append(["总条数", len(rows), len({int(x["paper"]) for x in rows}), "", "", "", "", ""])

    ws2 = wb.create_sheet("O组细分")
    ws2.append(["成形方式 + 热处理", "条数", "文献编号", "样本清单（前 10）"])
    buckets = collections.defaultdict(list)
    for x in groups["O"]:
        key = "%s + %s" % (x["form"] or "成形未注明",
                           "+".join(x["hts"]) if x["hts"] else "无热处理")
        buckets[key].append(x)
    for key, rs in sorted(buckets.items(), key=lambda z: -len(z[1])):
        ws2.append([key, len(rs),
                    ",".join(map(str, sorted({int(x["paper"]) for x in rs}))),
                    "; ".join("%s:%s" % (x["paper"], x["sample"]) for x in rs[:10])
                    + (" …" if len(rs) > 10 else "")])

    ws3 = wb.create_sheet("明细")
    ws3.append(["1.xlsx行号", "文献编号", "样本", "加工工艺（原文）", "成形方式",
                "热处理序列", "步数", "类别", "说明"])
    for x in sorted(rows, key=lambda z: z["row"]):
        ws3.append([x["row"], x["paper"], x["sample"], x["proc"], x["form"] or "未注明",
                    "+".join(x["hts"]) if x["hts"] else "无", len(x["hts"]),
                    CLS_NAME[x["cls"]], x["why"]])

    for w in (ws, ws2, ws3):
        w.freeze_panes = "A2"
        for col, wd in {"A": 10, "B": 9, "C": 30, "D": 62, "E": 12,
                        "F": 18, "G": 8, "H": 26, "I": 26}.items():
            w.column_dimensions[col].width = wd
    ws.column_dimensions["D"].width = 46
    ws2.column_dimensions["C"].width = 24
    ws2.column_dimensions["D"].width = 80

    wb.save(DST)
    print("\n已写出:", DST)


if __name__ == "__main__":
    main()
