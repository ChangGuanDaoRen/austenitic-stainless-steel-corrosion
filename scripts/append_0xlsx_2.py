# -*- coding: utf-8 -*-
r"""把「准」第 2 批 13 篇 / 58 条按 0.xlsx 原有版式追加录入，文献编号 43-55。

用法：
  D:/Python/python.exe scripts/append_0xlsx_2.py            生成 temp\lit\entry2\0_追加校验2.xlsx 并全量校验
  D:/Python/python.exe scripts/append_0xlsx_2.py --commit   校验通过后才把结果写回 0.xlsx

安全设计（为什么这么做）：
  - 一律从 备份\0_基准_追加前_20261104.xlsx 读原件，不从 0.xlsx 读：脚本可重复执行，
    不会把 49 条追加两遍；该备份与本轮改前的 0.xlsx 字节一致（MD5 已核）。
  - 默认只写 temp\ 校验件，不碰 0.xlsx；只有 --commit 才覆盖，覆盖前再次核验
    0.xlsx 未被别的进程改过，覆盖后再把成品重新载入校验一遍。
  - 前 159 行（含 2 行表头）逐格比对取值与样式，要求零差异；原有 80 个合并区全部保留。
  - 版式沿用现有文献块：A/Z/AA 按块合并，宋体 11 / 细边框 / 全居中，Z、AA 自动换行。
"""
import hashlib
import json
import os
import shutil
import sys
from copy import copy

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
BACKUP = os.path.join(WS, "备份", "0_基准_追加前_20261104_第2批.xlsx")
TARGET = os.path.join(WS, "0.xlsx")
DATA = os.path.join(WS, "temp", "lit", "entry2", "final_rows2.json")
STAGE = os.path.join(WS, "temp", "lit", "entry2", "0_追加校验2.xlsx")

NCOL = 27
OLD_MAX_ROW = 208
TPL_NORMAL_ROW = 208   # 最后一个文献块（42）的非锚点数据行
TPL_ANCHOR_ROW = 207   # 最后一个文献块（42）的锚点行
TPL_WRAP_ROW = 3       # Z / AA 带 wrap_text 的锚点行

# B..Y 共 24 列，顺序与 final_rows.json 的行字段一一对应
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "pH", "Epit", "stable"]
NUM_FIELDS = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
              "Ti", "Nb", "Co", "Sn"]


def md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def aa_text(paper):
    head = ("【本轮入库 2026-11-04｜文献%d.pdf｜原件与逐条取证见 "
            "准文献入库对照_20261104.md】" % paper["no"])
    return head + ("\n" + paper["notes"].strip() if paper["notes"].strip() else "")


def load_data():
    data = json.load(open(DATA, encoding="utf-8"))
    if len(data) != 13:
        raise SystemExit("数据篇数异常：%d（应为 13）" % len(data))
    if [p["no"] for p in data] != list(range(43, 56)):
        raise SystemExit("文献编号不是 43-55 连续序列")
    total = 0
    for p in data:
        n = len(p["rows"])
        total += n
        if n < 1:
            raise SystemExit("文献 %d 无数据行" % p["no"])
        if not p["citation"].strip():
            raise SystemExit("文献 %d 缺 citation" % p["no"])
        for r in p["rows"]:
            for k in FIELDS:
                if k not in r or r[k] is None:
                    raise SystemExit("文献 %d / %s 缺字段 %s" % (p["no"], r["sample"], k))
            if not (isinstance(r["Epit"], (int, float)) or r["Epit"] == "N/A"):
                raise SystemExit("文献 %d / %s Epit 非法：%r" % (p["no"], r["sample"], r["Epit"]))
            for k in NUM_FIELDS:
                v = r[k]
                if not (isinstance(v, (int, float)) or isinstance(v, str)):
                    raise SystemExit("文献 %d / %s / %s 类型非法" % (p["no"], r["sample"], k))
    if total != 58:
        raise SystemExit("数据条数异常：%d（应为 58）" % total)
    return data, total


def build(data):
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if ws.max_row != OLD_MAX_ROW or ws.max_column != NCOL:
        raise SystemExit("原件结构变了：%d 行 / %d 列（应为 159 / 27）" %
                         (ws.max_row, ws.max_column))

    tpl_normal = {c: copy(ws.cell(TPL_NORMAL_ROW, c)._style) for c in range(1, NCOL + 1)}
    tpl_anchor = {c: copy(ws.cell(TPL_ANCHOR_ROW, c)._style) for c in (1, 26, 27)}
    align_z = copy(ws.cell(TPL_WRAP_ROW, 26).alignment)
    align_aa = copy(ws.cell(TPL_WRAP_ROW, 27).alignment)

    r = OLD_MAX_ROW
    blocks = []
    for p in data:
        start = r + 1
        n = len(p["rows"])
        end = r + n
        for i, row in enumerate(p["rows"]):
            r += 1
            for c in range(1, NCOL + 1):
                ws.cell(r, c)._style = copy(tpl_normal[c])
            ws.cell(r, 1).value = p["no"] if i == 0 else None
            for j, key in enumerate(FIELDS):
                ws.cell(r, 2 + j).value = row[key]
            ws.cell(r, 26).value = p["citation"] if i == 0 else None
            ws.cell(r, 27).value = aa_text(p) if i == 0 else None
            if i == 0:
                for c in (1, 26, 27):
                    ws.cell(r, c)._style = copy(tpl_anchor[c])
                ws.cell(r, 26).alignment = copy(align_z)
                ws.cell(r, 27).alignment = copy(align_aa)
        if n > 1:
            for c in (1, 26, 27):
                ws.merge_cells(start_row=start, start_column=c,
                               end_row=end, end_column=c)
        blocks.append(dict(no=p["no"], start=start, end=end, n=n, citation=p["citation"],
                           aa=aa_text(p), samples=[x["sample"] for x in p["rows"]],
                           epit=[x["Epit"] for x in p["rows"]]))

    if r != OLD_MAX_ROW + 58:
        raise SystemExit("写入行数异常：末行 %d（应为 266）" % r)
    wb.save(STAGE)
    return blocks


def style_sig(cell):
    f, a, b = cell.font, cell.alignment, cell.border
    return (f.name, f.sz, f.b, f.i,
            a.horizontal, a.vertical, a.wrap_text,
            cell.number_format,
            b.left.style, b.right.style, b.top.style, b.bottom.style)


def verify(data, blocks, path):
    problems = []
    old = openpyxl.load_workbook(BACKUP)
    new = openpyxl.load_workbook(path)
    if old.sheetnames != new.sheetnames:
        problems.append("sheet 列表变了：%r -> %r" % (old.sheetnames, new.sheetnames))
    o, n = old["Sheet1"], new["Sheet1"]

    if o.max_row != OLD_MAX_ROW:
        problems.append("原件行数变了：%d" % o.max_row)
    if n.max_row != 266:
        problems.append("成品行数不对：%d（应为 266）" % n.max_row)
    if n.max_column != NCOL:
        problems.append("成品列数不对：%d（应为 27）" % n.max_column)

    changed = 0
    for r in range(1, OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            a, b = o.cell(r, c), n.cell(r, c)
            if a.value != b.value:
                changed += 1
                if changed <= 5:
                    problems.append("原有格取值被改：r%d c%d %r -> %r" %
                                    (r, c, a.value, b.value))
            elif style_sig(a) != style_sig(b):
                changed += 1
                if changed <= 5:
                    problems.append("原有格样式被改：r%d c%d" % (r, c))
    if changed:
        problems.append("前 159 行共有 %d 处差异" % changed)

    om = set(str(x) for x in o.merged_cells.ranges)
    nm = set(str(x) for x in n.merged_cells.ranges)
    if not om.issubset(nm):
        problems.append("原有合并区丢失：%s" % sorted(om - nm))
    multi = sum(1 for b in blocks if b["n"] > 1)
    if len(nm) != len(om) + 3 * multi:
        problems.append("新增合并区数量不对：%d（应为 %d）" %
                        (len(nm), len(om) + 3 * multi))

    ow = {k: v.width for k, v in o.column_dimensions.items()}
    nw = {k: v.width for k, v in n.column_dimensions.items()}
    if ow != nw:
        problems.append("列宽变了")
    if o.sheet_view.zoomScale != n.sheet_view.zoomScale:
        problems.append("缩放比例变了：%r -> %r" %
                        (o.sheet_view.zoomScale, n.sheet_view.zoomScale))

    for p, b in zip(data, blocks):
        s, e, cnt = b["start"], b["end"], b["n"]
        if n.cell(s, 1).value != p["no"]:
            problems.append("文献 %d 编号错：%r" % (p["no"], n.cell(s, 1).value))
        if n.cell(s, 26).value != b["citation"]:
            problems.append("文献 %d 文献信息错" % p["no"])
        if n.cell(s, 27).value != b["aa"]:
            problems.append("文献 %d 备注错" % p["no"])
        for rr in range(s + 1, e + 1):
            for c in (1, 26, 27):
                if n.cell(rr, c).value is not None:
                    problems.append("文献 %d 块内 r%d c%d 不该有值" % (p["no"], rr, c))
        for c in (1, 26, 27):
            rng = "A" if c == 1 else ("Z" if c == 26 else "AA")
            if cnt > 1 and ("%s%d:%s%d" % (rng, s, rng, e)) not in nm:
                problems.append("文献 %d 缺合并区 %s%d:%s%d" % (p["no"], rng, s, rng, e))
        for i, row in enumerate(p["rows"]):
            rr = s + i
            if n.cell(rr, 2).value != row["sample"]:
                problems.append("文献 %d 样本名错：r%d" % (p["no"], rr))
            if n.cell(rr, 24).value != row["Epit"]:
                problems.append("文献 %d Epit 错：r%d" % (p["no"], rr))
            for j, key in enumerate(FIELDS):
                if n.cell(rr, 2 + j).value != row[key]:
                    problems.append("文献 %d / %s 字段 %s 写入不符" %
                                    (p["no"], row["sample"], key))
                    break
            if n.cell(rr, 2).font.name != "宋体" or n.cell(rr, 2).font.sz != 11:
                problems.append("文献 %d / r%d 字体不是宋体 11" % (p["no"], rr))

    for s in ("Sheet2", "Sheet3"):
        if (old[s].max_row, old[s].max_column) != (new[s].max_row, new[s].max_column):
            problems.append("%s 尺寸变了" % s)
    return problems


def main():
    data, total = load_data()
    if not os.path.exists(BACKUP):
        raise SystemExit("缺备份：" + BACKUP)
    if md5(TARGET) not in (md5(BACKUP), md5(STAGE) if os.path.exists(STAGE) else ""):
        raise SystemExit("0.xlsx 已被其它进程改动，停止；先人工核对再执行")

    blocks = build(data)
    problems = verify(data, blocks, STAGE)

    print("数据：13 篇 / %d 条，编号 43-55，写入行 209-266" % total)
    for b in blocks:
        print("  文献 %d | 行 %d-%d | %d 条 | Epit=%s" %
              (b["no"], b["start"], b["end"], b["n"], b["epit"]))
    print("校验件：" + STAGE)
    if problems:
        print("校验未通过，共 %d 项：" % len(problems))
        for x in problems[:40]:
            print("  -", x)
        raise SystemExit(2)
    _nm = len(openpyxl.load_workbook(STAGE)["Sheet1"].merged_cells.ranges)
    _om = len(openpyxl.load_workbook(BACKUP)["Sheet1"].merged_cells.ranges)
    print("校验通过：前 %d 行零差异 / 原有 %d 个合并区全保留 / 新增 %d 个 / 列宽与缩放不变 / 58 行逐格回读一致"
          % (OLD_MAX_ROW, _om, _nm - _om))

    if "--commit" in sys.argv:
        if md5(TARGET) not in (md5(BACKUP), md5(STAGE)):
            raise SystemExit("提交前发现 0.xlsx 被改动，停止")
        shutil.copy2(STAGE, TARGET)
        problems2 = verify(data, blocks, TARGET)
        if problems2:
            print("写回后校验失败（原件仍在 备份\\ 可恢复）：")
            for x in problems2[:20]:
                print("  -", x)
            raise SystemExit(3)
        print("[commit] 已写回 0.xlsx，MD5=%s；写回后复验通过" % md5(TARGET))
    else:
        print("[dry] 未改 0.xlsx；确认无误后加 --commit 提交")


if __name__ == "__main__":
    main()
