# -*- coding: utf-8 -*-
"""构建「工艺相同严格程度」子集阶梯（L1 最严 → L6 最松）。

只读：铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版_v2.xlsx
只写：子集阶梯_20261026.xlsx

用法：python scripts/build_subset_ladder.py [dry|write]
"""
import sys
import re
import collections
import statistics as st

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SRC = "铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版_v2.xlsx"
DST = "子集阶梯_20261026.xlsx"

CAST_FORMS = {"铸态（造型未注明）", "砂型铸造", "熔模铸造", "真空精密铸造", "离心铸造"}
NO_HT = {"未提及热处理（原文未写）", "无热处理", "未固溶（铸态直接用）", "未提及"}
SURFACE_HINT = ("钝化", "酸洗", "抛光", "打磨", "钝化膜")

# ---------------------------------------------------------------- 解析工具

def num(v):
    return v if isinstance(v, (int, float)) else None


def parse_temp(s):
    """'1065-1120' -> 1092.5 ; '1100' -> 1100.0 ; None -> None"""
    if s is None:
        return None
    s = str(s).strip().replace("℃", "")
    m = re.match(r"^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)$", s)
    if m:
        return (float(m.group(1)) + float(m.group(2))) / 2
    try:
        return float(s)
    except ValueError:
        return None


def parse_time(s):
    """'2h'->2.0 ; '30min'->0.5 ; '0.5-5h'->2.75 ; None->None"""
    if s is None:
        return None
    s = str(s).strip()
    m = re.match(r"^(\d*\.?\d+)\s*-\s*(\d*\.?\d+)\s*h$", s)
    if m:
        return (float(m.group(1)) + float(m.group(2))) / 2
    for suf, div in (("min", 60.0), ("h", 1.0)):
        if s.endswith(suf):
            try:
                return float(s[: -len(suf)]) / div
            except ValueError:
                return None
    return None


def norm_cool(s):
    if s is None or str(s).strip() == "":
        return None
    s = str(s)
    if "水" in s:
        return "水冷类"
    for k in ("油冷", "空冷", "炉冷"):
        if k in s:
            return k
    return s


def norm_ht(s):
    s = "" if s is None else str(s)
    if s in NO_HT:
        return "无热处理类"
    return s


def connected_buckets(values, tol):
    """把温度值按 |差|<=tol 做链式连通分桶，返回 {值: 桶代表值}"""
    vals = sorted({v for v in values if v is not None})
    parent = {v: v for v in vals}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i, a in enumerate(vals):
        for b in vals[i + 1:]:
            if b - a <= tol:
                ra, rb = find(a), find(b)
                if ra != rb:
                    parent[rb] = ra
    comps = collections.defaultdict(list)
    for v in vals:
        comps[find(v)].append(v)
    out = {}
    for root, members in comps.items():
        rep = round(sum(members) / len(members), 1)
        for m in members:
            out[m] = rep
    return out


def pearson(xs, ys):
    pts = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
    if len(pts) < 3:
        return None
    xs2 = [p[0] for p in pts]
    ys2 = [p[1] for p in pts]
    mx, my = st.mean(xs2), st.mean(ys2)
    sxy = sum((a - mx) * (b - my) for a, b in pts)
    sxx = sum((a - mx) ** 2 for a in xs2)
    syy = sum((b - my) ** 2 for b in ys2)
    if sxx <= 0 or syy <= 0:
        return None
    return sxy / (sxx * syy) ** 0.5


def pooled_sd(groups, key="epit"):
    """组内合并标准差（按自由度加权）。groups: list[list[row]]"""
    num_, den = 0.0, 0
    for g in groups:
        v = [r[key] for r in g if r.get(key) is not None]
        if len(v) < 2:
            continue
        num_ += (len(v) - 1) * st.variance(v)
        den += len(v) - 1
    return (num_ / den) ** 0.5 if den else None


def rng(rows, key):
    v = [r[key] for r in rows if r.get(key) is not None]
    if not v:
        return None, None, None
    return min(v), max(v), (st.pstdev(v) if len(v) > 1 else None)


# ---------------------------------------------------------------- 读数据

def load():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    w1 = wb["01_清洗总表"]
    rows1 = [r for r in w1.iter_rows(min_row=3, values_only=True) if r and r[0] is not None]
    w2 = wb["02_工艺拆解"]
    rows2 = [r for r in w2.iter_rows(min_row=2, values_only=True) if r and r[0] is not None]
    if len(rows1) != len(rows2):
        raise SystemExit(f"行数不一致: 总表 {len(rows1)} vs 拆解 {len(rows2)}")

    recs = []
    for a, b in zip(rows1, rows2):
        if str(a[0]) != str(b[0]) or str(a[1]) != str(b[1]):
            raise SystemExit(f"逐行对齐失败: {a[0]}/{a[1]} vs {b[0]}/{b[1]}")
        cr, mo, nn = num(a[9]), num(a[10]), num(a[4])
        pren = None if None in (cr, mo, nn) else cr + 3.3 * mo + 16 * nn
        recs.append({
            "id": str(a[0]),
            "sample": str(a[1]),
            "paper": str(a[0]),
            "proc_raw": b[2],
            "form": b[3],
            "ht": b[4],
            "temp_raw": b[5],
            "time_raw": b[6],
            "cool_raw": b[7],
            "extra": b[8],
            "surface": b[9],
            "route": b[10],
            "epit": num(b[11]),
            "test_temp": b[12],
            "pH": num(a[22]),
            "Cr": cr, "Ni": num(a[11]), "Mo": mo, "N": nn,
            "Mn": num(a[8]), "Cu": num(a[7]),
            "pren": pren,
            "is_cast": b[3] in CAST_FORMS,
            "temp": parse_temp(b[5]),
            "time_h": parse_time(b[6]),
            "cool": norm_cool(b[7]),
            "ht_n": norm_ht(b[4]),
        })
    return recs


# ---------------------------------------------------------------- 级别定义

def build_levels(recs):
    temps = [r["temp"] for r in recs if r["temp"] is not None]
    b25 = connected_buckets(temps, 25)
    b50 = connected_buckets(temps, 50)

    def k_L1(r):
        return (r["route"],)

    def k_L2(r):
        return (r["form"], r["ht_n"],
                None if r["temp"] is None else round(r["temp"], 1),
                None if r["time_h"] is None else round(r["time_h"], 2),
                r["cool"])

    def k_L3(r):
        return (r["form"], r["ht_n"],
                None if r["temp"] is None else b25.get(r["temp"]),
                r["cool"])

    def k_L3b(r):
        return (r["form"], r["ht_n"],
                None if r["temp"] is None else b50.get(r["temp"]),
                None)

    def k_L4(r):
        return (r["form"], r["ht_n"])

    def k_L5(r):
        return ("铸件类" if r["is_cast"] else ("锻轧类" if ("轧" in str(r["form"]) or "锻" in str(r["form"])) else "未注明类"), r["ht_n"])

    def k_L6(r):
        return ("全部",)

    return [
        ("L1", "最严：7 字段全同（成形+热处理类型+温度+时间+冷却+附加+表面）", "仅铸件类", k_L1, True),
        ("L2", "写法归一：温度区间取中点、水淬≡水冷、时间统一小时、无热处理 4 种说法归一（忽略附加/表面）", "仅铸件类", k_L2, True),
        ("L3", "成形方式固定 + 热处理族固定 + 固溶温度 ±25℃ 链式 + 冷却方式固定", "仅铸件类", k_L3, True),
        ("L3b", "成形方式固定 + 热处理族固定 + 固溶温度 ±50℃ 链式（冷却转为特征）", "仅铸件类", k_L3b, True),
        ("L4", "成形方式固定 + 热处理族固定（温度/时间/冷却转为特征）", "仅铸件类", k_L4, True),
        ("L5", "铸造方法可合并，只固定热处理族（= 铸件类 + 固溶）", "仅铸件类", k_L5, True),
        ("L6", "不加工艺条件（含锻轧材，只用 3.5% NaCl + SCE 口径）", "全库", k_L6, False),
    ]


# ---------------------------------------------------------------- 统计

def level_stats(recs, keyfn, cast_only):
    pool = [r for r in recs if r["is_cast"]] if cast_only else recs
    groups = collections.defaultdict(list)
    for r in pool:
        groups[keyfn(r)].append(r)
    order = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    big_key, big = order[0]
    papers = collections.Counter(r["paper"] for r in big)
    e = [r["epit"] for r in big if r["epit"] is not None]
    cr_lo, cr_hi, cr_sd = rng(big, "Cr")
    ni_lo, ni_hi, ni_sd = rng(big, "Ni")
    rr = pearson([r["pren"] for r in big], [r["epit"] for r in big])
    return {
        "pool": pool,
        "groups": groups,
        "order": order,
        "n_groups": len(groups),
        "big_key": big_key,
        "big": big,
        "n": len(big),
        "papers": papers,
        "n_papers": len(papers),
        "epit_n": len(e),
        "epit_mean": st.mean(e) if e else None,
        "epit_sd": st.pstdev(e) if len(e) > 1 else None,
        "pooled_sd": pooled_sd([g for g in groups.values()]),
        "r_pren": rr,
        "cr": (cr_lo, cr_hi, cr_sd),
        "ni": (ni_lo, ni_hi, ni_sd),
        "mo": rng(big, "Mo"),
        "n10": sum(1 for _, g in order if len(g) >= 10),
        "n20": sum(1 for _, g in order if len(g) >= 20),
        "n30": sum(1 for _, g in order if len(g) >= 30),
    }


def judge(s):
    """给师兄看的四格判读"""
    if s["n"] < 15:
        size = "太小（<15）不能建模"
    elif s["n"] < 30:
        size = "偏小（15–29）只能留出"
    elif s["n"] < 50:
        size = "够留出（≥30）不够 5 折"
    else:
        size = "够 5 折（≥50）"
    if s["n_papers"] >= 5:
        leak = "可按文献 5 折"
    elif s["n_papers"] >= 3:
        leak = "只能按文献 3 折"
    else:
        leak = "文献太少，随机划分必泄漏"
    sd = s["cr"][2]
    div = "成分有变异（Cr sd %.2f）" % sd if sd and sd >= 1.0 else "成分近似常量（Cr sd %.2f）" % (sd or 0)
    return size, leak, div


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    recs = load()
    levels = build_levels(recs)

    results = []
    for code, desc, scope, keyfn, cast_only in levels:
        s = level_stats(recs, keyfn, cast_only)
        s["code"] = code
        s["desc"] = desc
        s["scope"] = scope
        results.append(s)
        size, leak, div = judge(s)
        print(f"{code} 组数={s['n_groups']:>2} 最大={s['n']:>3} 条 / {s['n_papers']} 篇 "
              f"Epit {s['epit_n']} 条 sd={s['epit_sd']:.3f} r(PREN)={s['r_pren']:.3f} "
              f"| {size} | {leak} | {div}")
        print(f"     键: {s['big_key']}")
        print(f"     文献: {dict(s['papers'])}")

    # 备选：锻轧 35 条
    roll = [r for r in recs if "轧" in str(r["form"])]
    roll_papers = collections.Counter(r["paper"] for r in roll)
    roll_epit = [r["epit"] for r in roll if r["epit"] is not None]
    print(f"\n备选(锻轧): {len(roll)} 条 / {len(roll_papers)} 篇 {dict(roll_papers)} "
          f"Epit {len(roll_epit)} 条 sd={st.pstdev(roll_epit):.3f} "
          f"r(PREN)={pearson([r['pren'] for r in roll], [r['epit'] for r in roll]):.3f}")

    if mode != "write":
        print("\n[dry] 未写文件")
        return

    write_book(recs, results, roll)
    print(f"\n已写出: {DST}")


# ---------------------------------------------------------------- 写工作簿

HFILL = PatternFill("solid", fgColor="1F4E79")
HFONT = Font(color="FFFFFF", bold=True, size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row=1, ncol=None):
    ncol = ncol or ws.max_column
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HFILL
        cell.font = HFONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 42
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_book(recs, results, roll):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---- 00_阶梯总览
    ws = wb.create_sheet("00_阶梯总览")
    head = ["级别", "判据（工艺相同的严格程度）", "范围", "组数", "最大组条数",
            "最大组工艺路线", "独立文献数", "文献构成（篇:条）", "Epit 覆盖",
            "Epit 均值±sd", "组内合并 sd(全级)", "r(PREN,Epit)",
            "Cr 范围(sd)", "Ni 范围(sd)", "Mo 范围",
            "≥10 条组数", "规模判读", "防泄漏判读", "成分多样性", "一句话结论"]
    ws.append(head)
    for s in results:
        size, leak, div = judge(s)
        concl = {
            "L1": "最贴合师兄原话，但只有 2 篇、无法验证",
            "L2": "归一写法后仍只有 2 篇",
            "L3": "条数与文献数第一个可用的平衡点",
            "L3b": "±50℃ 把 1050–1250℃ 串成一组，工艺已不严格",
            "L4": "成形一致 + 9 篇，能做一次留出",
            "L5": "唯一同时过 5 折 + 文献分折 + 成分多样",
            "L6": "混工艺，作为对照集",
        }.get(s["code"], "")
        ws.append([
            s["code"], s["desc"], s["scope"], s["n_groups"], s["n"],
            str(s["big_key"][0]) if len(s["big_key"]) == 1 else " | ".join("—" if k is None else str(k) for k in s["big_key"]),
            s["n_papers"], "; ".join(f"{k}:{v}" for k, v in s["papers"].most_common()),
            f"{s['epit_n']}/{s['n']}",
            "—" if s["epit_mean"] is None else f"{s['epit_mean']:.3f}±{s['epit_sd']:.3f}",
            "—" if s["pooled_sd"] is None else f"{s['pooled_sd']:.3f}",
            "—" if s["r_pren"] is None else f"{s['r_pren']:.3f}",
            f"{s['cr'][0]:.2f}–{s['cr'][1]:.2f} ({s['cr'][2]:.2f})" if s["cr"][0] is not None else "—",
            f"{s['ni'][0]:.2f}–{s['ni'][1]:.2f} ({s['ni'][2]:.2f})" if s["ni"][0] is not None else "—",
            f"{s['mo'][0]:.2f}–{s['mo'][1]:.2f}" if s["mo"][0] is not None else "—",
            s["n10"], size, leak, div, concl,
        ])
    style_header(ws)
    autosize(ws, [6, 46, 9, 6, 8, 40, 10, 26, 9, 15, 12, 10, 20, 20, 12, 9, 16, 18, 22, 34])

    # ---- 01..06 各级明细（最大组）
    titles = {"L1": "01_L1_最严", "L2": "02_L2_写法归一", "L3": "03_L3_温度±25",
              "L3b": "04_L3b_温度±50", "L4": "05_L4_成形+热处理", "L5": "06_L5_铸件+固溶",
              "L6": "07_L6_全库"}
    cols = ["级别", "文献编号", "样本", "成形方式", "热处理类型", "温度", "时间", "冷却",
            "附加处理", "表面处理", "固溶温度(数值)", "保温(h)", "冷却(归一)",
            "Cr", "Ni", "Mo", "N", "Mn", "Cu", "PREN", "Epit(V)", "pH", "测试温度", "工艺原文"]
    for s in results:
        ws = wb.create_sheet(titles[s["code"]])
        ws.append(cols)
        rows = sorted(s["big"], key=lambda r: (r["paper"], r["sample"]))
        for r in rows:
            ws.append([s["code"], r["paper"], r["sample"], r["form"], r["ht"], r["temp_raw"],
                       r["time_raw"], r["cool_raw"], r["extra"], r["surface"],
                       r["temp"], r["time_h"], r["cool"],
                       r["Cr"], r["Ni"], r["Mo"], r["N"], r["Mn"], r["Cu"], r["pren"],
                       r["epit"], r["pH"], r["test_temp"], r["proc_raw"]])
        style_header(ws)
        autosize(ws, [6, 8, 22, 20, 18, 12, 8, 10, 16, 22, 12, 9, 10,
                      8, 8, 8, 8, 8, 8, 9, 10, 7, 13, 52])

    # ---- 07_各级分组明细
    ws = wb.create_sheet("08_各级分组明细")
    ws.append(["级别", "组键", "条数", "文献数", "文献构成", "Epit 覆盖", "Epit 均值±sd", "r(PREN,Epit)", "成分"])
    for s in results:
        for key, g in s["order"]:
            papers = collections.Counter(r["paper"] for r in g)
            e = [r["epit"] for r in g if r["epit"] is not None]
            rr = pearson([r["pren"] for r in g], [r["epit"] for r in g])
            cr = [r["Cr"] for r in g if r["Cr"] is not None]
            ni = [r["Ni"] for r in g if r["Ni"] is not None]
            ws.append([
                s["code"],
                " | ".join("—" if k is None else str(k) for k in key),
                len(g), len(papers),
                "; ".join(f"{k}:{v}" for k, v in papers.most_common()),
                f"{len(e)}/{len(g)}",
                "—" if not e else f"{st.mean(e):.3f}±{(st.pstdev(e) if len(e) > 1 else 0):.3f}",
                "—" if rr is None else f"{rr:.3f}",
                (f"Cr {min(cr):.1f}–{max(cr):.1f} / Ni {min(ni):.1f}–{max(ni):.1f}"
                 if cr and ni else "—"),
            ])
    style_header(ws)
    autosize(ws, [6, 62, 7, 8, 26, 10, 15, 12, 28])

    # ---- 08_备选_锻轧35
    ws = wb.create_sheet("09_备选_锻轧35")
    ws.append(["警告", "文献编号", "样本", "成形方式", "热处理类型", "温度原文", "Cr", "Ni", "Mo", "N", "PREN", "Epit(V)", "备注"])
    for r in sorted(roll, key=lambda r: (r["paper"], r["sample"])):
        ws.append(["违反「只要铸造」口径；仅文献9、10 两篇", r["paper"], r["sample"], r["form"], r["ht"],
                   r["temp_raw"], r["Cr"], r["Ni"], r["Mo"], r["N"], r["pren"], r["epit"],
                   "原文：热锻 4 mm → 冷轧 2.5 mm，未提退火"])
    style_header(ws)
    autosize(ws, [40, 8, 22, 22, 22, 12, 8, 8, 8, 8, 9, 10, 40])

    # ---- 09_说明
    ws = wb.create_sheet("10_说明")
    lines = [
        ["子集阶梯说明", ""],
        ["", ""],
        ["目的", "把 116 条数据按「工艺相同」的严格程度切成 6 级，供选择建模子集。"],
        ["", ""],
        ["L1 最严", "7 字段全同（成形方式+热处理类型+温度+时间+冷却+附加处理+表面处理），且限铸件类。"],
        ["L2 写法归一", "温度区间取中点、水淬/水冷/CNT 统一为水冷类、时间统一到小时、4 种「无热处理」说法归一，忽略附加/表面。"],
        ["L3 ±25℃", "成形方式固定、热处理族固定、固溶温度按 ±25℃ 链式合并、冷却方式固定。"],
        ["L3b ±50℃", "同上但温度容差放宽到 ±50℃，冷却不再固定（转为特征）。注意：±50℃ 链式会把 1050→1250℃ 全串成一组，工艺已不严格。"],
        ["L4 成形+热处理", "成形方式固定、热处理族固定；固溶温度、时间、冷却转为特征列。"],
        ["L5 铸件+固溶", "5 种铸造方法合并为一类，只固定热处理族（固溶/无热处理/稳定化/敏化）。"],
        ["L6 全库", "不加任何工艺条件，只用 3.5% NaCl + SCE + Epit 口径过滤；含锻轧材。"],
        ["", ""],
        ["口径警告 1", "「未提及热处理（原文未写）」与「无热处理（原文明确未处理）」在 L2 起被并入「无热处理类」，"
                       "但两者含义不同；如需严格区分请说明。"],
        ["口径警告 2", "文献12、13 使用天然海水而非 3.5% NaCl（共 12 条，占库 10%），本阶梯未剔除，按师兄裁定处理。"],
        ["口径警告 3", "「组内合并 sd」= 该级别全部组按自由度加权的 Epit 标准差，用来判断分组是否真的提高了同质性；"
                       "若不低于全库 sd（约 0.378），说明按工艺分组并没有让 Epit 更集中。"],
        ["口径警告 4", "「r(PREN,Epit)」= 该组内 PREN 与 Epit 的相关系数，反映该组是否还剩足够的成分梯度可供学习；"
                       "接近 0 表示成分几乎不变，模型学不到成分效应。"],
        ["", ""],
        ["数据来源", "_清洗版_v2.xlsx 的 01_清洗总表（成分/Epit/pH）与 02_工艺拆解（工艺七字段）。"],
        ["生成脚本", "scripts/build_subset_ladder.py"],
        ["生成日期", "2026-10-26"],
    ]
    for a, b in lines:
        ws.append([a, b])
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 120
    for row in ws.iter_rows():
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
        row[0].font = Font(bold=True)
    ws["A1"].font = Font(bold=True, size=13)

    wb.save(DST)


if __name__ == "__main__":
    main()
