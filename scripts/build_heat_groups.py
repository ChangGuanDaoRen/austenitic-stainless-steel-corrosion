# -*- coding: utf-8 -*-
"""按「最后一步热处理的温度 + 时间」分组并导出工作簿（师兄 2026-10-26 口徑）。

输出：铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_热处理分组.xlsx
  - 分组统计：每个（温度×时间）组合的条数、文献、Epit 均值与分散、样本清单
  - 数据_116条：用户原表 27 列 + 最后一步热处理温度/时间/分组 + PREN
版式沿用用户自己那张总表（宋体 11、细框、全居中、两行表头、按文献块竖合并、缩放 125%）。
用法：python scripts/build_heat_groups.py [dry|write]
"""
import sys
import collections
import statistics as st

import openpyxl

import build_option_pack as B
import group_last_heat as G

DST = "铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_热处理分组.xlsx"


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    recs, full = B.load_full()

    groups = collections.defaultdict(list)
    info = {}
    for r in recs:
        key = (r["paper"], r["sample"])
        a = full[key]
        steps = G.steps(str(a[17]))
        temp, tm = (steps[-1] if steps else ("—", "—"))
        tag = f"{temp}℃×{tm}" if steps else "无热处理记录"
        info[key] = (temp, tm, tag)
        groups[tag].append(r)

    order = sorted(groups.items(), key=lambda x: (-len(x[1]), x[0]))
    print(f"总 {len(recs)} 条，{len(order)} 组")
    for tag, rows in order:
        e = [r["epit"] for r in rows if r.get("epit") is not None]
        print(f"  {tag:<22}{len(rows):>4} 条  Epit {'—' if not e else f'{st.mean(e):.3f}±{st.pstdev(e):.3f}'}")

    if mode != "write":
        print("\n[dry] 未写文件")
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 分组统计
    ws = wb.create_sheet("分组统计")
    ws.append(["最后一步热处理", "条数", "文献数", "文献（条数）", "Epit 均值±sd", "Cr 范围",
               "样本清单（最多 12 个）"])
    for tag, rows in order:
        e = [r["epit"] for r in rows if r.get("epit") is not None]
        papers = collections.Counter(str(r["paper"]) for r in rows)
        samples = [f"{r['paper']}:{r['sample']}" for r in
                   sorted(rows, key=lambda x: (int(x["paper"]), x["sample"]))]
        ws.append([
            tag, len(rows), len(papers),
            "; ".join(f"{k}号×{v}" for k, v in sorted(papers.items(), key=lambda x: int(x[0]))),
            "—" if not e else f"{st.mean(e):.3f}±{(st.pstdev(e) if len(e) > 1 else 0):.3f}",
            B.span([r["Cr"] for r in rows]),
            "; ".join(samples[:12]) + (f" …等 {len(samples)} 个" if len(samples) > 12 else ""),
        ])
    B.fmt(ws, ws.max_row, ws.max_column)
    B.widths(ws, [22, 7, 7, 30, 15, 13, 100])
    ws.sheet_view.zoomScale = 125

    # ---------------- 数据_116条
    ws = wb.create_sheet("数据_116条")
    last_col = 31
    ws.cell(row=1, column=1, value="文献编号")
    ws.cell(row=1, column=2, value="样本（原文命名）")
    ws.cell(row=1, column=3, value="化学成分(wt.%)")
    for j, (name, _) in enumerate(B.ELEMENTS):
        ws.cell(row=2, column=3 + j, value=name)
    for col, name in [(18, "加工工艺"), (19, "工作电极(WE)"), (20, "参比电极(RE)"),
                      (21, "对电极(CE)"), (22, "腐蚀温度"), (23, "腐蚀pH"), (24, "Epit(V)"),
                      (25, "测试前稳定时间"), (26, "文献信息"), (27, "备注")]:
        ws.cell(row=1, column=col, value=name)
    ws.cell(row=1, column=28, value="最后一步热处理")
    ws.cell(row=2, column=28, value="温度(℃)")
    ws.cell(row=2, column=29, value="保温时间")
    ws.cell(row=1, column=30, value="分组（温度×时间）")
    ws.cell(row=1, column=31, value="PREN")

    rows_sorted = sorted(recs, key=lambda x: (int(x["paper"]), x["sample"]))
    for i, r in enumerate(rows_sorted):
        a = full[(r["paper"], r["sample"])]
        temp, tm, tag = info[(r["paper"], r["sample"])]
        rr = 3 + i
        vals = ([a[0], a[1]] + [a[idx] for _, idx in B.ELEMENTS] +
                [a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24], a[25], a[26],
                 temp, tm, tag, r["pren"]])
        for j, v in enumerate(vals, start=1):
            ws.cell(row=rr, column=j, value=v)

    B.fmt(ws, 2 + len(rows_sorted), last_col)
    for col in [1, 2, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 30, 31]:
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=17)
    ws.merge_cells(start_row=1, start_column=28, end_row=1, end_column=29)
    B.merge_up(ws, 1, 3, [r["paper"] for r in rows_sorted])
    B.merge_up(ws, 26, 3, [full[(r["paper"], r["sample"])][25] for r in rows_sorted])
    B.merge_up(ws, 27, 3, [full[(r["paper"], r["sample"])][26] for r in rows_sorted])
    B.widths(ws, [9.0, 20.1, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17, 7.17,
                  7.17, 7.17, 7.17, 7.17, 7.17, 7.17,
                  55.3, 22.9, 15.3, 11.4, 8.75, 7.9, 9.17, 14.3, 70.4, 55.7,
                  10, 11, 18, 9])
    ws.sheet_view.zoomScale = 125

    wb.save(DST)
    print(f"\n已写出: {DST}")


if __name__ == "__main__":
    main()
