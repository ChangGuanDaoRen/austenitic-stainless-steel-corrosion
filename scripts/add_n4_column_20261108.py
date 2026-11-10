# -*- coding: utf-8 -*-
"""给 0.xlsx 增加第 27 列「口径标注」，只标 N4 两步热处理不合规行，不删任何行。

口径依据：口径基准_20261107.md 3.2 A3 —— 两步 = 固溶或退火后再叠敏化/时效/稳定化；
热轧+固溶、冷轧+退火不算两步；收录时标注不合规、不物理删行。

用法：
    python add_n4_column_20261108.py dry      # 只打印将写入的单元格
    python add_n4_column_20261108.py build    # 在 temp\\addcol\\0_加标注校验.xlsx 上生成
    python add_n4_column_20261108.py verify   # 校验校验件与原件：A–Z 逐格值+样式、合并区、列宽
"""
import os
import shutil
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "0.xlsx")
TMPDIR = os.path.join(ROOT, "temp", "addcol")
CHK = os.path.join(TMPDIR, "0_加标注校验.xlsx")

COL = 27  # AA
HEADER = "口径标注"

# 行级 N4 人工核定（对 60 块逐行复核后的结果，38 条）
N4_ROWS = {
    # 块3 砂型/离心铸造，1065-1120℃固溶5h + 550-650℃稳定化2h
    11: "N4-两步：固溶+稳定化",
    15: "N4-两步：固溶+稳定化",
    16: "N4-两步：固溶+稳定化",
    # 块7 304(敏化态)，1060℃保温0.5h水冷 + 650℃保温2h空冷
    26: "N4-两步：固溶+敏化(650℃×2h)",
    # 块16 铸态 1150℃固溶2h水淬 + 850℃敏化2h空冷
    90: "N4-两步：固溶+敏化",
    # 块34 冷轧态→固溶退火(1200℃等)水淬→900℃×0.5h时效空冷（r168 未固溶、未时效，不标）
    169: "N4-两步：固溶退火+时效",
    170: "N4-两步：固溶退火+时效",
    171: "N4-两步：固溶退火+时效",
    172: "N4-两步：固溶退火+时效",
    173: "N4-两步：固溶退火+时效",
    174: "N4-两步：固溶退火+时效",
    # 块49 铸态→1200℃固溶→650℃时效5000h
    229: "N4-两步：固溶+时效",
    230: "N4-两步：固溶+时效",
    # 块53 304-SA + 敏化 945K×1h
    252: "N4-两步：固溶+敏化",
}
# 块26 1100℃固溶2h水淬 + 550/650/750/850℃ × 6 时长 等温时效 = r131–r154
for _r in range(131, 155):
    N4_ROWS[_r] = "N4-两步：固溶+时效"


def nf(s):
    return unicodedata.normalize("NFKC", s if isinstance(s, str) else "")


def style_like(cell, src):
    cell.font = Font(name=src.font.name, sz=src.font.sz, b=src.font.b, i=src.font.i)
    cell.alignment = Alignment(
        horizontal=src.alignment.horizontal,
        vertical=src.alignment.vertical,
        wrap_text=src.alignment.wrap_text,
    )
    b = src.border
    cell.border = Border(
        left=Side(style=b.left.style),
        right=Side(style=b.right.style),
        top=Side(style=b.top.style),
        bottom=Side(style=b.bottom.style),
    )
    cell.fill = openpyxl.styles.PatternFill(fill_type=src.fill.fill_type,
                                            fgColor=src.fill.fgColor.rgb)


def do_dry():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Sheet1"]
    starts = [(r, ws.cell(r, 1).value) for r in range(3, ws.max_row + 1)
              if ws.cell(r, 1).value not in (None, "")]

    def blk(r):
        b = None
        for s, v in starts:
            if s <= r:
                b = v
            else:
                break
        return b

    print("将写入 %d 行：" % len(N4_ROWS))
    for r in sorted(N4_ROWS):
        print("  r%-4d 块%-3s %-22s %s" % (r, blk(r), nf(ws.cell(r, 2).value)[:20], N4_ROWS[r]))


def do_build():
    os.makedirs(TMPDIR, exist_ok=True)
    shutil.copyfile(SRC, CHK)
    wb = openpyxl.load_workbook(CHK)
    ws = wb["Sheet1"]
    z1 = ws.cell(1, 26)
    # 表头：AA1:AA2 合并，样式照 Z1
    ws.cell(1, COL).value = HEADER
    style_like(ws.cell(1, COL), z1)
    style_like(ws.cell(2, COL), ws.cell(2, 26))
    ws.merge_cells(start_row=1, start_column=COL, end_row=2, end_column=COL)
    # 数据行：仅写 N4 行
    n = 0
    for r in range(3, ws.max_row + 1):
        if r in N4_ROWS:
            c = ws.cell(r, COL)
            c.value = N4_ROWS[r]
            style_like(c, ws.cell(r, 26))
            n += 1
        else:
            style_like(ws.cell(r, COL), ws.cell(r, 26))
    ws.column_dimensions[get_column_letter(COL)].width = 30
    wb.save(CHK)
    print("生成 %s，写入 %d 行标注，总列数 %d" % (CHK, n, ws.max_column))


def _sig(c):
    return (
        c.value,
        c.font.name, c.font.sz, c.font.b,
        c.alignment.horizontal, c.alignment.vertical, c.alignment.wrap_text,
        c.border.left.style, c.border.right.style, c.border.top.style, c.border.bottom.style,
        c.fill.fill_type, c.fill.fgColor.rgb,
        c.number_format,
    )


def do_verify():
    a = openpyxl.load_workbook(SRC)
    b = openpyxl.load_workbook(CHK)
    wa, wb2 = a["Sheet1"], b["Sheet1"]
    prob = []
    if (wa.max_row, wa.max_column) != (wb2.max_row, 26):
        prob.append("原表尺寸 %s/%s" % (wa.max_row, wa.max_column))
    if wb2.max_column != 27:
        prob.append("校验件列数 %d != 27" % wb2.max_column)
    diff_val = diff_sty = 0
    for r in range(1, wa.max_row + 1):
        for c in range(1, 27):
            sa, sb = _sig(wa.cell(r, c)), _sig(wb2.cell(r, c))
            if sa[0] != sb[0]:
                diff_val += 1
                if diff_val <= 5:
                    prob.append("值差异 r%d c%d" % (r, c))
            elif sa[1:] != sb[1:]:
                diff_sty += 1
                if diff_sty <= 5:
                    prob.append("样式差异 r%d c%d" % (r, c))
    ma = set(str(m) for m in wa.merged_cells.ranges)
    mb = set(str(m) for m in wb2.merged_cells.ranges)
    added, lost = mb - ma, ma - mb
    if added != {"AA1:AA2"}:
        prob.append("新增合并区异常 %s" % sorted(added))
    if lost:
        prob.append("丢失合并区 %s" % sorted(lost))
    wa_w = {k: v.width for k, v in wa.column_dimensions.items() if k in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" and len(k) == 1}
    wb_w = {k: v.width for k, v in wb2.column_dimensions.items() if len(k) == 1}
    if {k: wa_w.get(k) for k in wa_w} != {k: wb_w.get(k) for k in wa_w}:
        prob.append("列宽变化")
    marked = [(r, wb2.cell(r, COL).value) for r in range(3, wb2.max_row + 1)
              if wb2.cell(r, COL).value not in (None, "")]
    if len(marked) != len(N4_ROWS):
        prob.append("标注行数 %d != %d" % (len(marked), len(N4_ROWS)))
    # 空值检查：除 AA 外不得出现新空格
    blanks = [(r, c) for r in range(3, wb2.max_row + 1) for c in range(1, 27)
              if wa.cell(r, c).value not in (None, "") and wb2.cell(r, c).value in (None, "")]
    if blanks:
        prob.append("出现空格 %d 处" % len(blanks))
    print("A–Z 值差异 %d 格 / 样式差异 %d 格 / 标注 %d 行 / 合并区 %d→%d" %
          (diff_val, diff_sty, len(marked), len(ma), len(mb)))
    print("表头 AA1=%r AA2=%r 合并=%s" % (wb2.cell(1, COL).value, wb2.cell(2, COL).value,
                                          "AA1:AA2" in mb))
    if prob:
        print("问题：")
        for p in prob:
            print("  -", p)
        return 1
    print("校验通过：除新增 AA 列外，A–Z 全表零差异。")
    return 0


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    if mode == "dry":
        do_dry()
    elif mode == "build":
        do_build()
    elif mode == "verify":
        sys.exit(do_verify())
    else:
        print(__doc__)
