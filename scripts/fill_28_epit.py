# -*- coding: utf-8 -*-
r"""把文献 28（Heliyon 2024, Mo/N 超奥）三行的 Epit 从 N/A 补为读图值。

为什么是填值而不是加行：B 类复核出的 3 条（SD_S1350630725010702，DOI 10.1016/j.heliyon.2024.e25964）
与库内文献 28 是同一篇、同 3 个试样（Mo2N1/Mo1N1/Mo1N2）、同条件（50℃/SCE/铂片），
第 1 批入库时因原文只在 Fig.3(a) 给曲线、无数值而记为 N/A；本轮读图得到 0.960/0.900/0.940。

用法：
  python scripts/fill_28_epit.py           只校验，写 temp\lit\entry2\0_文献28补E_校验.xlsx
  python scripts/fill_28_epit.py --commit  校验通过后写回 0.xlsx
全程从 备份\0_基准_文献28补Epit前_20261104.xlsx 读原件，可重复执行。
"""
import hashlib
import os
import shutil
import sys

import openpyxl

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
BACKUP = os.path.join(WS, "备份", "0_基准_文献28补Epit前_20261104.xlsx")
TARGET = os.path.join(WS, "0.xlsx")
STAGE = os.path.join(WS, "temp", "lit", "entry2", "0_文献28补E_校验.xlsx")

ROWS = {160: ("Mo2N1", 0.960), 161: ("Mo1N1", 0.900), 162: ("Mo1N2", 0.940)}
COL_EPIT, COL_AA = 24, 27
APPEND = ("\n【2026-11-04 19:2x 读图补录】3 条 Epit 原记 N/A（原文只在 p5 Fig.3(a) 给 3.5% NaCl/50℃ 动电位极化曲线，"
          "正文无数值）。本轮按师兄「不做 Tafel 外推、取阳极区拐点」口径渲染该图读数："
          "Mo2N1 0.960 V、Mo1N1 0.900 V、Mo1N2 0.940 V（vs SCE，50℃），读图误差约 ±0.02 V。"
          "同一篇的 A/B 类复核（文件 SD_S1350630725010702）与文献 28 为同文，未另立编号。")


def md5(p):
    with open(p, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def sig(c):
    f, a, b = c.font, c.alignment, c.border
    return (f.name, f.sz, f.b, f.i, a.horizontal, a.vertical, a.wrap_text,
            c.number_format, b.left.style, b.right.style, b.top.style, b.bottom.style)


def build():
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if (ws.max_row, ws.max_column) != (266, 27):
        raise SystemExit("原件结构变了：%d 行 / %d 列" % (ws.max_row, ws.max_column))
    if ws.cell(160, 1).value != 28:
        raise SystemExit("第 160 行不是文献 28 的块首：%r" % ws.cell(160, 1).value)
    for r, (name, _) in ROWS.items():
        if ws.cell(r, 2).value != name:
            raise SystemExit("第 %d 行样本名不符：%r（应为 %s）" % (r, ws.cell(r, 2).value, name))
        if ws.cell(r, COL_EPIT).value != "N/A":
            raise SystemExit("第 %d 行 Epit 不是 N/A：%r（已补过？）" % (r, ws.cell(r, COL_EPIT).value))
        if ws.cell(r, 22).value != "50℃" or ws.cell(r, 20).value != "SCE":
            raise SystemExit("第 %d 行条件不符（温度/参比）：%r %r" % (r, ws.cell(r, 22).value, ws.cell(r, 20).value))
    for r, (_, v) in ROWS.items():
        c = ws.cell(r, COL_EPIT)
        keep = c._style
        c.value = v
        c._style = keep
    aa = ws.cell(160, COL_AA)
    aa.value = (aa.value or "") + APPEND
    wb.save(STAGE)
    return wb


def verify():
    problems = []
    old = openpyxl.load_workbook(BACKUP)
    new = openpyxl.load_workbook(STAGE)
    o, n = old["Sheet1"], new["Sheet1"]
    if (n.max_row, n.max_column) != (266, 27):
        problems.append("行/列数变了：%d/%d" % (n.max_row, n.max_column))
    diff = []
    for r in range(1, 267):
        for c in range(1, 28):
            a, b = o.cell(r, c), n.cell(r, c)
            if a.value != b.value:
                diff.append((r, c))
            elif sig(a) != sig(b):
                problems.append("样式变了：r%d c%d" % (r, c))
    expect = {(r, COL_EPIT) for r in ROWS} | {(160, COL_AA)}
    if set(diff) != expect:
        problems.append("改动格不符：实际 %r，应为 %r" % (sorted(diff), sorted(expect)))
    for r, (name, v) in ROWS.items():
        if n.cell(r, COL_EPIT).value != v:
            problems.append("r%d 写入不符" % r)
        if n.cell(r, 2).value != name:
            problems.append("r%d 样本名被改" % r)
    if set(str(x) for x in o.merged_cells.ranges) != set(str(x) for x in n.merged_cells.ranges):
        problems.append("合并区变了")
    if {k: v.width for k, v in o.column_dimensions.items()} != {k: v.width for k, v in n.column_dimensions.items()}:
        problems.append("列宽变了")
    return problems


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if not os.path.exists(BACKUP):
        raise SystemExit("缺备份：" + BACKUP)
    build()
    problems = verify()
    print("改动：r160/r161/r162 的 Epit 与 r160 的 AA 备注，共 4 格")
    for r, (name, v) in ROWS.items():
        print("  r%d %-6s N/A -> %.3f V" % (r, name, v))
    print("校验件：" + STAGE)
    if problems:
        print("校验未通过：")
        for x in problems[:20]:
            print("  -", x)
        raise SystemExit(2)
    print("校验通过：除 4 格外逐格零差异 / 合并区与列宽不变")
    if "--commit" in sys.argv:
        if md5(TARGET) not in (md5(BACKUP), md5(STAGE)):
            raise SystemExit("0.xlsx 已被改动，停止")
        shutil.copy2(STAGE, TARGET)
        print("[commit] 已写回 0.xlsx，MD5=%s" % md5(TARGET))
    else:
        print("[dry] 未改 0.xlsx")


if __name__ == "__main__":
    main()
