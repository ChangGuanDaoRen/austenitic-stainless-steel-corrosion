#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 备份\00_录入_20261109\json\*.json 合并成项目根目录的 00.xlsx。

列结构与 0.xlsx 完全一致：26 列、两行表头、宋体 11、细边框、居中、Y/Z 自动换行、A 列按文献块竖合并。
只新建 00.xlsx，不改 0.xlsx / 1.xlsx。
"""
import json
import glob
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_DIR = os.path.join(ROOT, "备份", "00_录入_20261109", "json")
OUT = os.path.join(ROOT, "00.xlsx")

HEADERS = [
    ("A", "文献编号", None), ("B", "样本（原文命名）", None), ("C", "化学成分(wt.%)", "Q"),
    ("R", "加工工艺", None), ("S", "工作电极(WE)", None), ("T", "参比电极(RE)", None),
    ("U", "对电极(CE)", None), ("V", "腐蚀温度", None), ("W", "Epit(V)", None),
    ("X", "测试前稳定时间", None), ("Y", "文献信息", None), ("Z", "备注", None),
]
ELEMENTS = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni", "Ti", "Nb", "Co", "Sn", "Fe"]
# 列映射：B 样本、C–Q 元素、R 工艺、S 工作电极、T 参比、U 对电极、V 温度、W Epit、
# X 稳定时间、Y 文献信息（文件级）、Z 备注
ROW_KEYS = ["样本"] + ELEMENTS + ["工艺", "WE", "RE", "CE", "温度", "Epit", "稳定时间"]
NOTE_KEY = "备注"
Y_COL, Z_COL = 25, 26
WIDTHS = {"A": 9.0, "B": 20.1, "C": 7.2, "R": 72.3, "S": 22.9, "T": 15.3,
          "U": 11.4, "V": 8.8, "W": 7.9, "X": 9.2, "Y": 14.3, "Z": 70.4}


def normalize_epit(v):
    """把 Epit 统一成 V（只做 mV→V 单位归一，不做参比换算）。

    支持：单值、区间、±、带正负号与约/≈ 前缀。区间与 ± 保持字符串形式，单位统一为 V。
    """
    if isinstance(v, (int, float)):
        return v, False, ""
    s = str(v).strip()
    if not s or s.startswith("N/A"):
        return s, False, ""
    t = s.lstrip("约≈~ ").strip()
    if t.startswith("<") or t.startswith(">"):
        return s, False, ""

    def scale_of(unit, a):
        if unit == "mV":
            return 0.001
        if unit == "V":
            return 1.0
        return 1.0 if abs(a) <= 2 else 0.001

    m = re.match(r"^([+-]?\d+(?:\.\d+)?)\s*±\s*([+-]?\d+(?:\.\d+)?)\s*(mV|V)?$", t)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        k = scale_of(m.group(3), a)
        return f"{a * k:g}±{b * k:g}", True, f"{s} → {a * k:g}±{b * k:g} V"

    m = re.match(r"^([+-]?\d+(?:\.\d+)?)\s*[-–~]\s*([+-]?\d+(?:\.\d+)?)\s*(mV|V)?$", t)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        k = scale_of(m.group(3), a)
        lo, hi = sorted((a * k, b * k))
        return f"{lo:g}-{hi:g}", True, f"{s} → {lo:g}-{hi:g} V"

    m = re.match(r"^([+-]?\d+(?:\.\d+)?)\s*(mV|V)?$", t)
    if m:
        a = float(m.group(1))
        k = scale_of(m.group(2), a)
        return round(a * k, 4), True, f"{s} → {a * k:g} V"

    return s, False, ""


def main():
    files = sorted(glob.glob(os.path.join(JSON_DIR, "*.json")))
    if not files:
        print("no json")
        sys.exit(1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="宋体", size=11)
    center = Alignment(horizontal="center", vertical="center")
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_center = Alignment(vertical="center")

    # 表头
    for col, title, merge_to in HEADERS:
        ws[f"{col}1"] = title
        ws[f"{col}1"].font = font
        ws[f"{col}1"].alignment = center
        ws[f"{col}1"].border = border
        if merge_to:
            ws[f"{col}1"].alignment = center
            ws.merge_cells(f"{col}1:{merge_to}1")
        else:
            ws[f"{col}2"] = title
            ws[f"{col}2"].font = font
            ws[f"{col}2"].alignment = center
            ws[f"{col}2"].border = border
            ws.merge_cells(f"{col}1:{col}2")
    for i, el in enumerate(ELEMENTS):
        c = ws.cell(2, 3 + i, el)
        c.font = font
        c.alignment = center
        c.border = border
    for r in (1, 2):
        for col in ("C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"):
            ws[f"{col}{r}"].border = border
            ws[f"{col}{r}"].font = font
            ws[f"{col}{r}"].alignment = center

    # 数据
    row = 3
    blocks = []
    audit = []
    for f in files:
        d = json.load(open(f, encoding="utf-8"))
        rows = d.get("rows") or []
        if not rows:
            continue
        start = row
        for i, rec in enumerate(rows):
            ws.cell(row, 1, d["no"] if i == 0 else None)
            for j, key in enumerate(ROW_KEYS):
                v = rec.get(key, "未提及")
                if key == "Epit":
                    new, changed, how = normalize_epit(v)
                    if changed:
                        audit.append([d["no"], row, v, new, how])
                    v = new
                cell = ws.cell(row, 2 + j, v)
                if key == "工艺":
                    cell.alignment = left_center
                else:
                    cell.alignment = center
            ws.cell(row, Y_COL, d.get("文献信息", "") if i == 0 else None).alignment = wrap_center
            ws.cell(row, Z_COL, rec.get(NOTE_KEY, "未提及")).alignment = wrap_center
            for c in range(1, 27):
                cell = ws.cell(row, c)
                cell.font = font
                cell.border = border
            row += 1
        if len(rows) > 1:
            ws.merge_cells(start_row=start, start_column=1, end_row=row - 1, end_column=1)
            ws.cell(start, 1).alignment = center
        blocks.append((d["no"], len(rows), start, row - 1))

    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w

    wb.save(OUT)
    audit_path = os.path.join(ROOT, "备份", "00_录入_20261109", "00_Epit单位归一_20261109.csv")
    with open(audit_path, "w", newline="", encoding="utf-8-sig") as fh:
        import csv as _csv
        w = _csv.writer(fh)
        w.writerow(["文献编号", "00.xlsx行号", "原值", "归一后(V)", "说明"])
        w.writerows(audit)
    print(f"[OK] {OUT}")
    print(f"[AUDIT] {audit_path} 改动 {len(audit)} 行")
    print(f"文献块 {len(blocks)} / 数据行 {row - 3} / 列 26")
    print("块前5:", blocks[:5], "块后3:", blocks[-3:])


if __name__ == "__main__":
    main()
