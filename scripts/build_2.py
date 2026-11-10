# -*- coding: utf-8 -*-
"""从 1.xlsx（清洗版，150 条）抽出「有最后一步热处理温度+时间记录」的子集，生成 2.xlsx。

口径（师兄 10-26 20:51「只考虑最后一步热处理的温度和时间」的筛选式读法）：
  保留 加工工艺 列里能读到「温度(℃)+时间(h/min)」的行；
  「最后一步」= 工艺串里最后一个「温度×时间」
  （例 `1100℃固溶2h，水淬 + 850℃×720min等温时效` → 850℃×720min）；
  不再按成形方式、前置工序、冷却方式卡。

列：1.xlsx 已是 25 列（A–Y，无文献信息/备注），原样保留；文献编号逐行填值便于追溯。
版式：宋体 11 / 细边框 / 全居中（强制）/ 列宽照 1.xlsx / 缩放 150 / 不设冻结窗格。
"""
import re
import sys
import collections
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

SRC = sys.argv[1] if len(sys.argv) > 1 else "1.xlsx"
DST = sys.argv[2] if len(sys.argv) > 2 else "2.xlsx"
SHEET = "Sheet1"
KEEP_COLS = 25            # A..Y
COL_PROC = 18             # 1-based，加工工艺
PAT = re.compile(r"(\d{3,4})\s*℃[^，,]{0,10}?(\d+(?:\.\d+)?)\s*(h|min|小时|分钟)")


def parse_tt(text):
    ms = PAT.findall(str(text or ""))
    if not ms:
        return None
    t, v, unit = ms[-1]
    t = int(t)
    v = float(v)
    hours = v / 60.0 if unit in ("min", "分钟") else v
    return t, hours


def copy_style(src, dst):
    dst.font = src.font.copy()
    dst.border = src.border.copy()
    dst.fill = src.fill.copy()
    dst.alignment = Alignment(horizontal="center", vertical="center",
                              wrap_text=src.alignment.wrap_text)
    dst.number_format = src.number_format


def main():
    wb0 = openpyxl.load_workbook(SRC)
    ws0 = wb0[SHEET]

    kept, dropped = [], []
    for r in range(3, ws0.max_row + 1):
        doc = ws0.cell(r, 1).value
        if doc in (None, ""):
            continue
        tt = parse_tt(ws0.cell(r, COL_PROC).value)
        if tt is None:
            dropped.append((r, doc))
        else:
            kept.append((r, doc, tt[0], tt[1]))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    # 表头两行
    for r in (1, 2):
        for c in range(1, KEEP_COLS + 1):
            s = ws0.cell(r, c)
            d = ws.cell(r, c)
            d.value = s.value
            copy_style(s, d)

    # 表头合并（落在 A..Y 内）
    for m in ws0.merged_cells.ranges:
        if m.min_row <= 2 and m.max_col <= KEEP_COLS:
            ws.merge_cells(str(m))

    # 数据行
    out_r = 3
    for r, doc, t, h in kept:
        for c in range(1, KEEP_COLS + 1):
            s = ws0.cell(r, c)
            d = ws.cell(out_r, c)
            d.value = s.value
            copy_style(s, d)
        out_r += 1

    # 合并区被覆盖的格子补一遍对齐，保证整表全居中
    for r in range(1, ws.max_row + 1):
        for c in range(1, KEEP_COLS + 1):
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")

    # 列宽 / 视图
    for c in range(1, KEEP_COLS + 1):
        L = get_column_letter(c)
        if L in ws0.column_dimensions:
            ws.column_dimensions[L].width = ws0.column_dimensions[L].width
    ws.sheet_view.zoomScale = 150
    ws.sheet_view.zoomScaleNormal = 150
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"

    # 说明页
    ws2 = wb.create_sheet("说明")
    by_doc = collections.Counter(d for _, d, _, _ in kept)
    temps = collections.Counter(t for _, _, t, _ in kept)
    hours = collections.Counter(round(h, 2) for _, _, _, h in kept)
    band = [x for x in kept if 1050 <= x[2] <= 1150]
    lines = [
        ("口径", ""),
        ("来源", "%s 的 Sheet1（清洗版 150 条）" % SRC),
        ("筛选条件", "加工工艺列同时含「温度(℃)」与「时间(h/min)」的行"),
        ("最后一步", "取工艺列里最后一个「温度×时间」；如「1100℃固溶2h + 850℃×720min时效」记 850℃×720min"),
        ("不卡的条件", "成形方式、前置工序、冷却方式"),
        ("列", "A–Y 共 25 列，与原表一致；文献编号逐行填值便于追溯"),
        ("", ""),
        ("结果", ""),
        ("总行数（不含表头）", len(kept)),
        ("文献篇数", len(by_doc)),
        ("最后一步 ≥1000℃（固溶）", "%d 条" % sum(v for k, v in temps.items() if k >= 1000)),
        ("最后一步 <1000℃（时效/稳定化）", "%d 条" % sum(v for k, v in temps.items() if k < 1000)),
        ("1050–1150℃ 固溶带", "%d 条" % len(band)),
        ("", ""),
        ("按文献条数", ""),
    ]
    for d in sorted(by_doc):
        lines.append(("文献%s" % d, "%d 条" % by_doc[d]))
    lines += [("", ""), ("最后一步温度分布(℃)", "")]
    for t in sorted(temps):
        lines.append((str(t), "%d 条" % temps[t]))
    lines += [("", ""), ("最后一步时间分布(h)", "")]
    for h in sorted(hours):
        lines.append((str(h), "%d 条" % hours[h]))
    lines += [("", ""), ("未纳入的行（%d 行）" % len(dropped), "原表行号 / 文献编号")]
    for r, doc in dropped:
        lines.append((str(r), "文献%s" % doc))

    for i, (a, b) in enumerate(lines, start=1):
        ws2.cell(i, 1).value = a
        ws2.cell(i, 2).value = b
        ws2.cell(i, 1).font = Font(name="宋体", size=11)
        ws2.cell(i, 2).font = Font(name="宋体", size=11)
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 62

    wb.save(DST)

    print("来源 %s -> 写出 %s" % (SRC, DST))
    print("保留 %d 条 / %d 篇；未纳入 %d 条" % (len(kept), len(by_doc), len(dropped)))
    print("按文献:", dict(sorted(by_doc.items())))
    print("最后一步温度:", dict(sorted(temps.items())))
    print("1050-1150℃:", len(band), "条")


if __name__ == "__main__":
    main()
