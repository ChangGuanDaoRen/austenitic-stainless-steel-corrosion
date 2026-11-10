# -*- coding: utf-8 -*-
r"""第 3 批入库：第 6 轮检索 8 篇 / 25 条 -> 0.xlsx，文献编号 56-63。

用法：
  D:/Python/python.exe scripts/append_0xlsx_3.py            生成 temp\lit\entry3\0_追加校验3.xlsx 并全量校验
  D:/Python/python.exe scripts/append_0xlsx_3.py --commit   校验通过后才写回 0.xlsx

与第 2 批脚本的差异：本版 0.xlsx 是 26 列（pH 列在 11-05 删酸时移除），
B..X 为 23 个数据字段，Y=文献信息，Z=备注。
一律从 备份\0_基准_第3批入库前_20261105.xlsx 读原件，保证可重复执行、不会追加两遍。
"""
import hashlib
import json
import os
import shutil
import sys
from copy import copy

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

WS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP = os.path.join(WS, "备份", "0_基准_第3批入库前_20261105.xlsx")
TARGET = os.path.join(WS, "0.xlsx")
DATA = os.path.join(WS, "temp", "lit", "entry3", "final_rows3.json")
STAGE = os.path.join(WS, "temp", "lit", "entry3", "0_追加校验3.xlsx")

NCOL = 26
OLD_MAX_ROW = 260
NEW_ROWS = 25
NEW_MAX_ROW = OLD_MAX_ROW + NEW_ROWS          # 285
TPL_NORMAL_ROW = 260                           # 文献55 非锚点数据行
TPL_ANCHOR_ROW = 259                           # 文献55 锚点行
TPL_WRAP_ROW = 3                               # Y/Z 带 wrap_text 的数据行

FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]             # B..X = 2..24


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def aa_text(p):
    head = ("【本轮入库 2026-11-05｜文献%d.pdf｜原件与逐条取证见 "
            "第6轮检索_新候选_20261105.md】" % p["no"])
    return head + "\n" + p["notes"].strip()


def load_data():
    data = json.load(open(DATA, encoding="utf-8"))
    if len(data) != 8:
        raise SystemExit("数据篇数异常：%d（应为 8）" % len(data))
    if [p["no"] for p in data] != list(range(56, 64)):
        raise SystemExit("文献编号不是 56-63 连续序列")
    total = 0
    for p in data:
        n = len(p["rows"])
        total += n
        if n < 1 or not p["citation"].strip() or not p["notes"].strip():
            raise SystemExit("文献 %d 缺数据/题录/备注" % p["no"])
        for r in p["rows"]:
            for k in FIELDS:
                if k not in r or r[k] is None or r[k] == "":
                    raise SystemExit("文献 %d / %s 缺字段 %s" % (p["no"], r["sample"], k))
            if not (isinstance(r["Epit"], (int, float)) or r["Epit"] == "<-0.35"):
                raise SystemExit("文献 %d / %s Epit 非法：%r" % (p["no"], r["sample"], r["Epit"]))
    if total != NEW_ROWS:
        raise SystemExit("数据条数异常：%d（应为 25）" % total)
    return data


def build(data):
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if ws.max_row != OLD_MAX_ROW or ws.max_column != NCOL:
        raise SystemExit("原件结构变了：%d 行 / %d 列（应为 260 / 26）" %
                         (ws.max_row, ws.max_column))

    tpl_normal = {c: copy(ws.cell(TPL_NORMAL_ROW, c)._style) for c in range(1, NCOL + 1)}
    tpl_anchor = {c: copy(ws.cell(TPL_ANCHOR_ROW, c)._style) for c in (1, 25, 26)}
    align_y = copy(ws.cell(TPL_WRAP_ROW, 25).alignment)
    align_z = copy(ws.cell(TPL_WRAP_ROW, 26).alignment)

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
            ws.cell(r, 25).value = p["citation"] if i == 0 else None
            ws.cell(r, 26).value = aa_text(p) if i == 0 else None
            if i == 0:
                for c in (1, 25, 26):
                    ws.cell(r, c)._style = copy(tpl_anchor[c])
                ws.cell(r, 25).alignment = copy(align_y)
                ws.cell(r, 26).alignment = copy(align_z)
        if n > 1:
            for c in (1, 25, 26):
                ws.merge_cells(start_row=start, start_column=c,
                               end_row=end, end_column=c)
        blocks.append(dict(no=p["no"], start=start, end=end, n=n,
                           citation=p["citation"], aa=aa_text(p),
                           samples=[x["sample"] for x in p["rows"]],
                           epit=[x["Epit"] for x in p["rows"]]))

    if r != NEW_MAX_ROW:
        raise SystemExit("写入行数异常：末行 %d（应为 %d）" % (r, NEW_MAX_ROW))
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

    if n.max_row != NEW_MAX_ROW:
        problems.append("成品行数不对：%d（应为 %d）" % (n.max_row, NEW_MAX_ROW))
    if n.max_column != NCOL:
        problems.append("成品列数不对：%d（应为 26）" % n.max_column)

    changed = 0
    for r in range(1, OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            a, b = o.cell(r, c), n.cell(r, c)
            if a.value != b.value:
                changed += 1
                if changed <= 5:
                    problems.append("原有格取值被改：r%d c%d %r -> %r" % (r, c, a.value, b.value))
            elif style_sig(a) != style_sig(b):
                changed += 1
                if changed <= 5:
                    problems.append("原有格样式被改：r%d c%d" % (r, c))
    if changed:
        problems.append("前 %d 行共有 %d 处差异" % (OLD_MAX_ROW, changed))

    om = set(str(x) for x in o.merged_cells.ranges)
    nm = set(str(x) for x in n.merged_cells.ranges)
    if not om.issubset(nm):
        problems.append("原有合并区丢失：%s" % sorted(om - nm))
    multi = sum(1 for b in blocks if b["n"] > 1)
    if len(nm) != len(om) + 3 * multi:
        problems.append("新增合并区数量不对：%d（应为 %d）" % (len(nm), len(om) + 3 * multi))

    ow = {k: v.width for k, v in o.column_dimensions.items()}
    nw = {k: v.width for k, v in n.column_dimensions.items()}
    if ow != nw:
        problems.append("列宽变了")
    if o.sheet_view.zoomScale != n.sheet_view.zoomScale:
        problems.append("缩放比例变了")

    for p, b in zip(data, blocks):
        s, e, cnt = b["start"], b["end"], b["n"]
        if n.cell(s, 1).value != p["no"]:
            problems.append("文献 %d 编号错：%r" % (p["no"], n.cell(s, 1).value))
        if n.cell(s, 25).value != b["citation"]:
            problems.append("文献 %d 文献信息错" % p["no"])
        if n.cell(s, 26).value != b["aa"]:
            problems.append("文献 %d 备注错" % p["no"])
        for rr in range(s + 1, e + 1):
            for c in (1, 25, 26):
                if n.cell(rr, c).value is not None:
                    problems.append("文献 %d 块内 r%d c%d 不该有值" % (p["no"], rr, c))
        for c, rng in ((1, "A"), (25, "Y"), (26, "Z")):
            if cnt > 1 and ("%s%d:%s%d" % (rng, s, rng, e)) not in nm:
                problems.append("文献 %d 缺合并区 %s%d:%s%d" % (p["no"], rng, s, rng, e))
        for i, row in enumerate(p["rows"]):
            rr = s + i
            for j, key in enumerate(FIELDS):
                if n.cell(rr, 2 + j).value != row[key]:
                    problems.append("文献 %d / %s 字段 %s 写入不符（r%d c%d）" %
                                    (p["no"], row["sample"], key, rr, 2 + j))
                    break
            if n.cell(rr, 2).font.name != "宋体" or n.cell(rr, 2).font.sz != 11:
                problems.append("文献 %d / r%d 字体不是宋体 11" % (p["no"], rr))
            if n.cell(rr, 25).alignment.wrap_text is not True and i == 0:
                problems.append("文献 %d / r%d 题录未自动换行" % (p["no"], rr))

    for s in ("Sheet2", "Sheet3"):
        if (old[s].max_row, old[s].max_column) != (new[s].max_row, new[s].max_column):
            problems.append("%s 尺寸变了" % s)
    return problems


def main():
    data = load_data()
    # 首次执行创建基准备份；已有则不动，保证脚本可重复执行
    if not os.path.exists(BACKUP):
        shutil.copy2(TARGET, BACKUP)
        print("已创建基准备份：" + BACKUP)
    if md5(TARGET) not in (md5(BACKUP), md5(STAGE) if os.path.exists(STAGE) else ""):
        raise SystemExit("0.xlsx 已被其它进程改动，停止；先人工核对再执行")

    blocks = build(data)
    problems = verify(data, blocks, STAGE)

    print("数据：8 篇 / 25 条，编号 56-63，写入行 261-285")
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
    print("校验通过：前 %d 行零差异 / 原有 %d 个合并区全保留 / 新增 %d 个 / 列宽与缩放不变 / 25 行逐格回读一致"
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
