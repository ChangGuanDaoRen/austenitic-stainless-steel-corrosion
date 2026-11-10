# -*- coding: utf-8 -*-
"""把「工艺相同」的 4 个可选档位做成可以直接挑的选项包。

只读：铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版_v2.xlsx
只写：铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_建模档位.xlsx

与 build_subset_ladder.py 的区别：那份是 L1→L6 六级梯度的分析报告；
这份只留 4 个候选档位，每一档都导出成「可直接用的建模表」，
并给出各档差异（放宽一步具体多了哪些样本），供师兄看数据挑档位。

版式：完全按用户自己那张总表（铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）.xlsx）的习惯
      —— 宋体 11、细边框、水平垂直居中、无底色无加粗、两行表头、
         文献编号/文献信息/备注按文献块竖合并、不冻结窗格、缩放 125%。

用法：python scripts/build_option_pack.py [dry|write]
"""
import sys
import collections
import statistics as st

sys.path.insert(0, "scripts")

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import build_subset_ladder as L

SRC = L.SRC
DST = "铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_建模档位.xlsx"
# lite 版：只留 选项对照 + 四个档位的建模表，不要 档位差异 与 说明
DST_LITE = "铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_档位.xlsx"

# 只保留 4 个候选档位（对应阶梯里的 L3 / L4 / L5 / L6）
OPTIONS = [
    ("A", "L3", "成形方式 + 热处理类型 + 固溶温度 ±25℃ + 冷却方式 全相同"),
    ("B", "L4", "成形方式 + 热处理类型 相同（温度、时间、冷却改作特征）"),
    ("C", "L5", "只固定热处理类型（固溶），五种铸造方法合并、铸造方法作特征"),
    ("D", "L6", "不加工艺条件（全库，含锻轧材）"),
]

# 元素列顺序、列名、取值下标 —— 与用户自己那张总表完全一致
ELEMENTS = [("C", 2), ("Si", 3), ("N", 4), ("P", 5), ("S", 6), ("Cu", 7), ("Mn", 8),
            ("Cr", 9), ("Mo", 10), ("Ni", 11), ("Ti", 12), ("Nb", 13), ("Co", 14),
            ("Sn", 15), ("Fe", 16)]

# 版式沿用用户自己表格的习惯：宋体 11、细边框、全居中、无底色无加粗、不冻结窗格
FONT = Font(name="宋体", size=11, bold=False)
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_full():
    """返回 (recs, full) —— recs 走阶梯脚本的解析，full 保留 01 表原始 27 列"""
    recs = L.load()
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["01_清洗总表"]
    full = {}
    for a in ws.iter_rows(min_row=3, values_only=True):
        if a and a[0] is not None:
            full[(str(a[0]), str(a[1]))] = a
    return recs, full


def fmt(ws, last_row, last_col, start_row=1, border=True):
    """整片区域套用户格式：宋体 11 / 居中 / 细边框"""
    for r in range(start_row, last_row + 1):
        for c in range(1, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT
            cell.alignment = CENTER
            if border:
                cell.border = BORDER


def widths(ws, vals):
    for i, w in enumerate(vals, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def merge_up(ws, col, first_row, keys):
    """按连续相同的 key 竖向合并（用户表里文献编号、文献信息、备注就是这样合并的）"""
    n = len(keys)
    start = 0
    for i in range(1, n + 1):
        if i == n or keys[i] != keys[start]:
            if i - start > 1:
                ws.merge_cells(start_row=first_row + start, start_column=col,
                               end_row=first_row + i - 1, end_column=col)
            start = i


def span(vals):
    vals = [v for v in vals if v is not None]
    if not vals:
        return "—"
    return f"{min(vals):.2f}–{max(vals):.2f}"


def judge_size(n):
    if n < 15:
        return "<15：不能单独建模"
    if n < 30:
        return "15–29：只能做留出"
    if n < 50:
        return "30–49：够留出，不够 5 折"
    return "≥50：够 5 折"


def judge_paper(np_):
    if np_ >= 5:
        return "≥5 篇：可按文献分折"
    if np_ >= 3:
        return "3–4 篇：只能按文献 3 折"
    return "<3 篇：随机划分必泄漏"


def judge_div(sd):
    if sd is None:
        return "成分数据不足"
    return "Cr sd ≥1：成分有梯度" if sd >= 1.0 else f"Cr sd {sd:.2f}<1：成分近似常量"


def option_stats(recs, levels_by_code, code):
    lv = levels_by_code[code]
    s = L.level_stats(recs, lv[3], lv[4])
    big = s["big"]
    e = [r["epit"] for r in big if r["epit"] is not None]
    return s, big, e


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    lite = ("--lite" in sys.argv) or ("lite" in args)
    mode = next((a for a in args if a in ("dry", "write")), "dry")
    recs, full = load_full()
    levels = L.build_levels(recs)
    by_code = {lv[0]: lv for lv in levels}

    packed = []
    for tag, code, desc in OPTIONS:
        s, big, e = option_stats(recs, by_code, code)
        packed.append({"tag": tag, "code": code, "desc": desc, "s": s, "rows": big, "epit": e})
        forms = collections.Counter(r["form"] for r in big)
        temps = collections.Counter("—" if r["temp"] is None else f"{r['temp']:.0f}" for r in big)
        print(f"[{tag}] {code} 条数={len(big)} 篇={s['n_papers']} 组数={s['n_groups']}")
        print(f"     文献: {dict(sorted(collections.Counter(r['paper'] for r in big).items(), key=lambda x: int(x[0])))}")
        print(f"     成形: {dict(forms)}")
        print(f"     温度: {dict(sorted(temps.items()))}")
        print(f"     Cr {span([r['Cr'] for r in big])}  Ni {span([r['Ni'] for r in big])}  "
              f"Epit {span(e)}  r(PREN)={'—' if s['r_pren'] is None else round(s['r_pren'], 3)}")
        print(f"     {judge_size(len(big))} | {judge_paper(s['n_papers'])} | {judge_div(s['cr'][2])}")

    overlaps = []
    for i in range(1, len(packed)):
        prev = {(r["paper"], r["sample"]) for r in packed[i - 1]["rows"]}
        cur = packed[i]["rows"]
        added = [r for r in cur if (r["paper"], r["sample"]) not in prev]
        lost = [r for r in packed[i - 1]["rows"] if (r["paper"], r["sample"]) not in
                {(x["paper"], x["sample"]) for x in cur}]
        overlaps.append((packed[i - 1]["tag"], packed[i]["tag"], added, lost))
        print(f"差异 {packed[i-1]['tag']}→{packed[i]['tag']}: 新增 {len(added)} 条，掉出 {len(lost)} 条")

    if mode != "write":
        print("\n[dry] 未写文件")
        return

    out = write_book(packed, overlaps, full, lite=lite)
    print(f"\n已写出: {out}")


def write_book(packed, overlaps, full, lite=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 选项对照
    ws = wb.create_sheet("选项对照")
    head = ["选项", "工艺相同的口径", "条数", "文献数", "文献构成（篇:条）",
            "铸造方法构成（条）", "固溶温度分布（℃:条）", "同工艺组数", "最大组条数",
            "Epit 条数", "Epit 范围(V)", "Epit 均值±sd", "r(PREN,Epit)",
            "Cr 范围", "Ni 范围", "Mo 范围",
            "条数判读", "文献判读", "成分判读",
            "与「热轧+固溶、冷轧+退火」的关系"]
    ws.append(head)
    for p in packed:
        s, rows, e = p["s"], p["rows"], p["epit"]
        forms = collections.Counter(r["form"] for r in rows)
        temps = collections.Counter("—" if r["temp"] is None else f"{r['temp']:.0f}" for r in rows)
        papers = collections.Counter(r["paper"] for r in rows)
        has_cast = any(r["is_cast"] for r in rows)
        has_roll = any("轧" in str(r["form"]) for r in rows)
        if p["tag"] == "D":
            rel = "含 35 条冷轧（原文未写退火）、0 条热轧；铸锻混合"
        elif has_cast and not has_roll:
            rel = "全部为铸件，无热轧、无冷轧+退火样本"
        else:
            rel = "—"
        ws.append([
            p["tag"], p["desc"], len(rows), s["n_papers"],
            "; ".join(f"{k}:{v}" for k, v in sorted(papers.items(), key=lambda x: int(x[0]))),
            "; ".join(f"{k}:{v}" for k, v in forms.most_common()),
            "; ".join(f"{k}:{v}" for k, v in sorted(temps.items(), key=lambda x: x[0])),
            s["n_groups"], len(rows),
            f"{len(e)}/{len(rows)}",
            span(e) if e else "—",
            "—" if not e else f"{st.mean(e):.3f}±{(st.pstdev(e) if len(e) > 1 else 0):.3f}",
            "—" if s["r_pren"] is None else f"{s['r_pren']:.3f}",
            span([r["Cr"] for r in rows]), span([r["Ni"] for r in rows]), span([r["Mo"] for r in rows]),
            judge_size(len(rows)), judge_paper(s["n_papers"]), judge_div(s["cr"][2]),
            rel,
        ])
    fmt(ws, ws.max_row, ws.max_column)
    widths(ws, [6, 34, 7, 7, 30, 30, 24, 8, 8, 8, 14, 15, 11, 14, 14, 12, 17, 19, 20, 32])
    ws.sheet_view.zoomScale = 125

    # ---------------- 各档建模表（版式 = 用户总表 27 列 + 右侧工艺拆解/建模列）
    for p in packed:
        ws = wb.create_sheet(f"{p['tag']}_{len(p['rows'])}条")
        n = len(p["rows"])
        last_col = 27 + 10
        # 第一行：大类；第二行：细项
        ws.cell(row=1, column=1, value="文献编号")
        ws.cell(row=1, column=2, value="样本（原文命名）")
        ws.cell(row=1, column=3, value="化学成分(wt.%)")
        for j, (name, _) in enumerate(ELEMENTS):
            ws.cell(row=2, column=3 + j, value=name)
        ws.cell(row=1, column=18, value="加工工艺")
        ws.cell(row=1, column=19, value="工作电极(WE)")
        ws.cell(row=1, column=20, value="参比电极(RE)")
        ws.cell(row=1, column=21, value="对电极(CE)")
        ws.cell(row=1, column=22, value="腐蚀温度")
        ws.cell(row=1, column=23, value="腐蚀pH")
        ws.cell(row=1, column=24, value="Epit(V)")
        ws.cell(row=1, column=25, value="测试前稳定时间")
        ws.cell(row=1, column=26, value="文献信息")
        ws.cell(row=1, column=27, value="备注")
        ws.cell(row=1, column=28, value="工艺拆解")
        for j, name in enumerate(["成形方式", "热处理类型", "固溶温度(℃)", "保温时间",
                                  "冷却方式", "附加处理", "表面处理"]):
            ws.cell(row=2, column=28 + j, value=name)
        ws.cell(row=1, column=35, value="工艺路线ID")
        ws.cell(row=1, column=36, value="铸造方法类")
        ws.cell(row=1, column=37, value="PREN")

        rows_sorted = sorted(p["rows"], key=lambda x: (int(x["paper"]), x["sample"]))
        for i, r in enumerate(rows_sorted):
            a = full.get((r["paper"], r["sample"]))
            if a is None:
                raise SystemExit(f"01 表里找不到 {r['paper']}/{r['sample']}")
            rr = 3 + i
            vals = [a[0], a[1]] + [a[idx] for _, idx in ELEMENTS] + [
                a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24], a[25], a[26],
                r["form"], r["ht"], (r["temp"] if r["temp"] is not None else "未提及"),
                r["time_raw"], r["cool_raw"], r["extra"], r["surface"],
                r["route"], ("锻轧" if not r["is_cast"] else "铸件"), r["pren"],
            ]
            for j, v in enumerate(vals, start=1):
                ws.cell(row=rr, column=j, value=v)

        fmt(ws, 2 + n, last_col)
        # 表头合并
        for col in [1, 2, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 35, 36, 37]:
            ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=17)
        ws.merge_cells(start_row=1, start_column=28, end_row=1, end_column=34)
        # 同文献块的 文献编号 / 文献信息 / 备注 竖合并；工艺路线ID 同路线合并
        merge_up(ws, 1, 3, [r["paper"] for r in rows_sorted])
        merge_up(ws, 26, 3, [full[(r["paper"], r["sample"])][25] for r in rows_sorted])
        merge_up(ws, 27, 3, [full[(r["paper"], r["sample"])][26] for r in rows_sorted])
        merge_up(ws, 35, 3, [r["route"] for r in rows_sorted])
        widths(ws, [9.0, 20.1, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17,
                    7.17, 7.17, 7.17, 7.17, 7.17, 7.17,
                    55.3, 22.9, 15.3, 11.4, 8.75, 7.9, 9.17, 14.3, 70.4, 55.7,
                    22, 26, 11, 9, 11, 16, 20, 50, 12, 9])
        ws.sheet_view.zoomScale = 125

    if lite:
        wb.save(DST_LITE)
        return DST_LITE

    # ---------------- 档位差异
    ws = wb.create_sheet("档位差异")
    ws.append(["从 → 到", "放宽了什么", "新增条数", "新增样本（文献:样本）", "掉出条数", "掉出样本"])
    relax = {"A→B": "温度、时间、冷却 由固定改为特征列",
             "B→C": "成形方式（五种铸造方法）由固定改为特征列",
             "C→D": "取消全部工艺条件，并入 35 条冷轧与商用锻材"}
    for prev, cur, added, lost in overlaps:
        key = f"{prev}→{cur}"
        ws.append([key, relax.get(key, ""), len(added),
                   "; ".join(f"{r['paper']}:{r['sample']}" for r in added[:40]) or "—",
                   len(lost),
                   "; ".join(f"{r['paper']}:{r['sample']}" for r in lost[:40]) or "—"])
    fmt(ws, ws.max_row, ws.max_column)
    widths(ws, [8, 36, 9, 90, 9, 40])
    ws.sheet_view.zoomScale = 125

    # ---------------- 说明（不带表格线，纯文字）
    ws = wb.create_sheet("说明")
    lines = [
        ["建模档位选项包", ""],
        ["", ""],
        ["用途", "四个候选档位，每档都导出成可直接建模的表；看数据挑档位，不做推荐。"],
        ["", ""],
        ["A", "成形方式 + 热处理类型 + 固溶温度 ±25℃ + 冷却方式 全相同。"],
        ["B", "成形方式 + 热处理类型 相同；固溶温度、保温时间、冷却方式作为特征列。"],
        ["C", "只固定热处理类型（主要为固溶）；砂型/熔模/离心/真空精密/铸态 合并，铸造方法作为特征列。"],
        ["D", "不加工艺条件；只保证 3.5% NaCl（或 0.6 mol/L）+ Epit 统一 SCE 的口径。"],
        ["", ""],
        ["口径提醒 1", "「未提及热处理（原文未写）」是原文没写，不等于未处理；本包未做区分，与阶梯脚本口径一致。"],
        ["口径提醒 2", "文献12、13 用天然海水而非 3.5% NaCl，共 12 条，已按师兄 10-16 的答复保留在库内。"],
        ["口径提醒 3", "文献20 的 5 条原文无元素成分，未纳入本包（原始行留在 00_原始数据 页）。"],
        ["口径提醒 4", "r(PREN,Epit) 越接近 0 说明该档内成分几乎不变、模型学不到成分效应；越低越不利。"],
        ["口径提醒 5", "D 档含 35 条冷轧（原文未写退火）与 1 条商用锻材，与「只要铸造」的口径不符，仅作对照。"],
        ["", ""],
        ["数据来源", "_清洗版_v2.xlsx 的 01_清洗总表 与 02_工艺拆解"],
        ["生成脚本", "scripts/build_option_pack.py"],
        ["生成日期", "2026-10-26"],
    ]
    for a, b in lines:
        ws.append([a, b])
    fmt(ws, ws.max_row, 2, border=False)
    for row in ws.iter_rows():
        row[0].alignment = Alignment(horizontal="left", vertical="center")
        row[1].alignment = Alignment(horizontal="left", vertical="center")
    widths(ws, [16, 110])
    ws.sheet_view.zoomScale = 125

    wb.save(DST)
    return DST


if __name__ == "__main__":
    main()
