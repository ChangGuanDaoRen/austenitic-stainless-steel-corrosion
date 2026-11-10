# -*- coding: utf-8 -*-
"""在「工艺分组」基础上给出候选建模子集阶梯（师兄 10-24 口径：尽量选相同工艺）。

粒度由严到宽：
  S1 铸造族 + 固溶 + 固溶温度 1100±25℃
  S2 铸造族 + 固溶（不限温度）
  S3 铸造族 + 铸态直接使用（无热处理 / 未固溶 / 原文未写热处理）
  S4 铸造族（全部）
  S5 铸锭→热锻→冷轧（文献9,10，原文未写最终热处理）
  S6 全部 150 条（不筛）

用法：python scripts/subset_process_1.py [dry|write]
  write 在 1.xlsx 的「工艺分组」sheet 末尾追加 E 段
"""
import re
import statistics
import sys

import openpyxl

from group_process_1 import (parse_form, parse_heat, SRC, COL_ID, COL_NAME,
                             COL_PROC, COL_EPIT, COL_CR, COL_MO, COL_NI)

COL_N = 5
CAST = {"铸态（造型未注明）", "砂型铸造", "离心铸造", "熔模铸造",
        "真空精密铸造", "铸造（造型未注明）"}


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def temp_mid(t):
    """把 '1065-1120' 之类取中值，返回 float 或 None。"""
    if not t:
        return None
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(t))]
    return sum(nums) / len(nums) if nums else None


def pren(rec):
    cr, mo, n = num(rec["cr"]), num(rec["mo"]), num(rec["n"])
    if cr is None or mo is None or n is None:
        return None
    return cr + 3.3 * mo + 16 * n


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    wb = openpyxl.load_workbook(SRC)
    ws = wb["Sheet1"]
    recs = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=COL_NAME).value
        if not name:
            continue
        proc = str(ws.cell(row=r, column=COL_PROC).value or "")
        recs.append({"id": ws.cell(row=r, column=COL_ID).value, "name": name,
                     "proc": proc,
                     "form": parse_form(proc), "heat": parse_heat(proc),
                     "epit": ws.cell(row=r, column=COL_EPIT).value,
                     "cr": ws.cell(row=r, column=COL_CR).value,
                     "mo": ws.cell(row=r, column=COL_MO).value,
                     "ni": ws.cell(row=r, column=COL_NI).value,
                     "n": ws.cell(row=r, column=COL_N).value})
    for x in recs:
        x["tm"] = temp_mid(x["heat"]["温度"])
        x["pr"] = pren(x)

    def sub(pred):
        return [x for x in recs if pred(x)]

    is_cast = lambda x: x["form"] in CAST
    is_sol = lambda x: x["heat"]["类型"] == "固溶"
    no_ht = lambda x: x["heat"]["类型"] in ("无热处理", "未固溶（铸态直接用）",
                                            "未提及热处理（原文未写）", "未提及")

    cands = [
        ("S1 铸造族 + 固溶 + 1100±25℃", sub(lambda x: is_cast(x) and is_sol(x)
                                        and x["tm"] is not None and abs(x["tm"] - 1100) <= 25)),
        ("S2 铸造族 + 固溶（不限温度）", sub(lambda x: is_cast(x) and is_sol(x))),
        ("S3 铸造族 + 铸态直接使用", sub(lambda x: is_cast(x) and no_ht(x))),
        ("S4 铸造族（全部）", sub(is_cast)),
        ("S5 铸锭→热锻→冷轧", sub(lambda x: x["form"] == "铸锭→热锻→冷轧")),
        ("S6 全部（不筛工艺）", recs),
    ]

    print("【候选建模子集阶梯】")
    print(f"{'方案':<28}{'条数':>5}{'文献':>6}{'Epit范围(V)':>20}{'Epit sd':>10}{'PREN范围':>18}")
    rows = []
    for label, items in cands:
        eps = [num(x["epit"]) for x in items]
        eps = [e for e in eps if e is not None]
        prs = [x["pr"] for x in items if x["pr"] is not None]
        papers = sorted({str(x["id"]) for x in items}, key=lambda v: int(v))
        sd = statistics.pstdev(eps) if len(eps) > 1 else 0.0
        print(f"{label:<28}{len(items):>5}{len(papers):>6}"
              f"{f'{min(eps):.3f}~{max(eps):.3f}':>20}{sd:>10.3f}"
              f"{f'{min(prs):.1f}~{max(prs):.1f}' if prs else '-':>18}")
        rows.append([label, len(items), ",".join(papers),
                     f"{min(eps):.3f}~{max(eps):.3f}" if eps else "-",
                     round(sd, 3),
                     f"{min(prs):.1f}~{max(prs):.1f}" if prs else "-"])

    for label, items in cands:
        if not items:
            continue
        print(f"\n-- {label}（{len(items)} 条）明细")
        by_route = {}
        for x in items:
            by_route.setdefault((x["form"], x["heat"]["类型"],
                                 x["heat"]["温度"], x["heat"]["时间"]), []).append(x)
        for k, v in sorted(by_route.items(), key=lambda kv: -len(kv[1])):
            eps = [num(x["epit"]) for x in v if num(x["epit"]) is not None]
            print(f"   {len(v):>3} 条  {k[0]} | {k[1]}"
                  f"{(' | ' + str(k[2]) + '℃') if k[2] else ''}"
                  f"{(' | ' + str(k[3])) if k[3] else ''}"
                  f"   Epit {min(eps):.3f}~{max(eps):.3f}  文献"
                  f"{','.join(sorted({str(x['id']) for x in v}, key=lambda z: int(z)))}")

    if mode != "write":
        print("\n[dry] 未写盘")
        return

    g = wb["工艺分组"]
    start = g.max_row + 2
    g.cell(row=start, column=1, value="E. 候选建模子集阶梯（师兄 10-24 口径：尽量选相同工艺）")
    g.cell(row=start + 1, column=1, value="方案")
    g.cell(row=start + 1, column=2, value="条数")
    g.cell(row=start + 1, column=3, value="文献编号")
    g.cell(row=start + 1, column=4, value="Epit范围(V)")
    g.cell(row=start + 1, column=5, value="Epit sd")
    g.cell(row=start + 1, column=6, value="PREN范围")
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            g.cell(row=start + 2 + i, column=1 + j, value=v)
    r = start + 2 + len(rows) + 1
    g.cell(row=r, column=1, value="说明：PREN = Cr + 3.3Mo + 16N；"
                                   "师兄举例的「热轧+固溶」与「冷轧+退火」在本库中均为 0 条，"
                                   "最接近的替代是铸造族+固溶。")
    wb.save(SRC)
    print(f"\n[write] 已追加到 1.xlsx「工艺分组」sheet 第 {start} 行起")


if __name__ == "__main__":
    main()
