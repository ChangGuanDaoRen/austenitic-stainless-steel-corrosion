# -*- coding: utf-8 -*-
"""把 xlsx 的一个 sheet 渲染成 HTML 预览（只为核对版式，不修改任何 xlsx）。

用法：python scripts/preview_xlsx.py <xlsx路径> <sheet名> <输出html>
"""
import sys
import html

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries


def build(src, sheet, out, max_row=None, max_col=None):
    wb = openpyxl.load_workbook(src)
    ws = wb[sheet]
    nrow = max_row or ws.max_row
    ncol = max_col or ws.max_column

    # 合并区：左上角记 (rowspan, colspan)，其余标记跳过
    span = {}
    skip = set()
    for rng in ws.merged_cells.ranges:
        c1, r1, c2, r2 = range_boundaries(str(rng))
        span[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1):
                    skip.add((r, c))

    def css(cell):
        f = cell.font
        st = []
        st.append(f"font-family:'{f.name or '宋体'}',SimSun")
        st.append(f"font-size:{f.sz or 11}pt")
        if f.bold:
            st.append("font-weight:bold")
        if f.color is not None and f.color.rgb not in (None, "00000000"):
            st.append(f"color:#{str(f.color.rgb)[-6:]}")
        fill = cell.fill
        if fill is not None and fill.fill_type == "solid":
            st.append(f"background:#{str(fill.fgColor.rgb)[-6:]}")
        al = cell.alignment
        st.append(f"text-align:{al.horizontal or 'center'}")
        st.append(f"vertical-align:{'middle' if al.vertical in (None, 'center') else al.vertical}")
        b = cell.border
        for side, name in ((b.left, "left"), (b.right, "right"), (b.top, "top"), (b.bottom, "bottom")):
            if side is not None and side.style:
                st.append(f"border-{name}:1px solid #000000")
        st.append("padding:1px 3px")
        st.append("white-space:nowrap" if not al.wrap_text else "white-space:normal")
        return ";".join(st)

    cols = []
    for c in range(1, ncol + 1):
        d = ws.column_dimensions.get(get_column_letter(c))
        # Excel 列宽 -> px 近似
        w = (d.width if d and d.width else 8.43)
        cols.append(f"<col style='width:{w * 7.0:.0f}px'>")

    rows_html = []
    for r in range(1, nrow + 1):
        tds = []
        for c in range(1, ncol + 1):
            if (r, c) in skip:
                continue
            cell = ws.cell(row=r, column=c)
            v = cell.value
            text = "" if v is None else html.escape(str(v))
            rs, cs = span.get((r, c), (1, 1))
            attr = (f" rowspan='{rs}'" if rs > 1 else "") + (f" colspan='{cs}'" if cs > 1 else "")
            tds.append(f"<td{attr} style='{css(cell)}'>{text}</td>")
        h = ws.row_dimensions[r].height
        hh = f" style='height:{h * 1.33:.0f}px'" if h else ""
        rows_html.append(f"<tr{hh}>" + "".join(tds) + "</tr>")

    doc = f"""<!doctype html><html><head><meta charset='utf-8'>
<style>body{{margin:0;background:#fff}} table{{border-collapse:collapse;table-layout:fixed}}</style>
</head><body><table>{''.join(cols)}{''.join(rows_html)}</table></body></html>"""
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    print("written:", out, "rows:", nrow, "cols:", ncol)


if __name__ == "__main__":
    a = sys.argv[1:]
    build(a[0], a[1], a[2],
          int(a[3]) if len(a) > 3 else None,
          int(a[4]) if len(a) > 4 else None)
