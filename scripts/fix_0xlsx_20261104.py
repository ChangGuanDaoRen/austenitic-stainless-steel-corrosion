# -*- coding: utf-8 -*-
r"""修正 0.xlsx 第一批（28–42）复核发现的 4 处硬错，并同步 final_rows.json 与备注。

修正内容（行号为 0.xlsx 当前行号）：
  165 R  文献30 1#      工艺改为「商用 6 mm 热轧+固溶退火态」，不套 2#–4# 的真空熔炼路线
  171 E  文献32 N       0.075 -> 未提及（原文无 N 列，0.075 是 V）
  171 H  文献32 Cu      0.05  -> 0.11
  171 N  文献32 Nb      0.002 -> 0.02
  171 O  文献32 Co      0.11  -> 0.05
  173 R  文献34 A29     工艺改为「冷轧态（as received，未固溶、未时效）」
  188 K  文献38 204Cu  Mo   未提及 -> 0.03
  160/165/171/185/191 AA 备注同步修正（28 的 N/A 矛盾、30 的 1# 说明、32 的两行表头说明、
                              38 的 Mo 归属、39 的模板残留句）

安全流程：备份已由外部完成（备份\0_基准_改前_20261104_1912.xlsx）；本脚本只写
temp\lit\entry\0_修正校验.xlsx，校验通过且源文件 MD5 未变才覆盖 0.xlsx。
"""
import hashlib
import json
import os
import shutil
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
TARGET = os.path.join(WS, "0.xlsx")
BACKUP = os.path.join(WS, "备份", "0_基准_改前_20261104_1912.xlsx")
STAGE = os.path.join(WS, "temp", "lit", "entry", "0_修正校验.xlsx")
JSON_PATH = os.path.join(WS, "temp", "lit", "entry", "final_rows.json")
JSON_BACKUP = os.path.join(WS, "备份", "final_rows_改前_20261104_1912.json")

CELL_FIXES = {
    "R165": "商用 6 mm 热轧+固溶退火态（太钢，原文未给固溶温度/时间）",
    "E171": "未提及",
    "H171": 0.11,
    "N171": 0.02,
    "O171": 0.05,
    "R173": "冷轧态（as received，未固溶、未时效）",
    "K188": 0.03,
}

NOTE_FIXES = {
    "AA160": [
        ("3 条 Epit 一律记「未提及」", "3 条 Epit 一律记「N/A」"),
        ("仅在 p6 Fig.3(a) 图中", "仅在 p5–p6 Fig.3(a) 图中"),
    ],
    "AA165": [
        ("原文另有 6%FeCl3 浸泡试验，本表只录 3.5% NaCl 极化数据。",
         "原文另有 6%FeCl3 浸泡试验，本表只录 3.5% NaCl 极化数据。"
         "1# 为太钢商用 6 mm 热轧固溶退火态，工艺与 2#–4# 的真空熔炼+锻造+热轧+1150℃×1h 固溶不同，"
         "原文未给 1# 的固溶温度/时间。"),
    ],
    "AA171": [
        ("可再补 3 条。",
         "可再补 3 条。成分表为两行表头（Fe/C/Si/Mn/P/S/Cr/Ni + Mo/Al/Cu/Co/Ti/Nb/V/W），"
         "无 N 列；Cu 0.11、Co 0.05、Nb 0.02，V 0.075 不入表。"),
    ],
    "AA185": [
        ("Mo 列 txt 只留 4 个值（0.09/0.11/0.21/0.03）——后三组（204Cu、204Cu/Nb、304）在 PDF 表格中疑为合并单元格共用 0.03，按列序还原均为 0.03，数值 0.03 本身可在 txt 找到，但逐格对应关系存疑，已在 notes 标注，请主程序复核。",
         "Mo 列 4 个值 0.09/0.11/0.21/0.03 依次属 201A/201B/201Cu/204Cu；204Cu/Nb 与 304 两格空白，记「未提及」。"),
        ("【主程复核】p3 Table 1 的 Mo 行文字层只有 4 个数值(0.09/0.11/0.21/0.03)，后 3 组(204Cu/204Cu-Nb/304)归属不明(疑为合并单元格)，本轮改记「未提及」，不猜。",
         "【主程复核】p3 Table 1 的 Mo 行 4 个数值(0.09/0.11/0.21/0.03)按列对齐 201A/201B/201Cu/204Cu；204Cu/Nb 与 304 空白，记「未提及」。"),
    ],
    "AA191": [
        ("另：本篇亦为易切削/高 S 无关，但为 303Cu 之外的高合金",
         "另：本篇与易切削/高硫钢无关，是 303Cu 之外的另一类高合金"),
    ],
}

JSON_NOTE_FIXES = {
    28: NOTE_FIXES["AA160"],
    30: NOTE_FIXES["AA165"],
    32: NOTE_FIXES["AA171"],
    38: NOTE_FIXES["AA185"],
    39: NOTE_FIXES["AA191"],
}

JSON_ROW_FIXES = {
    30: {"process": CELL_FIXES["R165"]},
    32: {"N": "未提及", "Cu": 0.11, "Nb": 0.02, "Co": 0.05},
    34: {"process": CELL_FIXES["R173"]},
    38: {"Mo": 0.03},
}


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def style_tuple(cell):
    return (
        cell.font.name, cell.font.sz, cell.font.b,
        cell.alignment.horizontal, cell.alignment.vertical, cell.alignment.wrap_text,
        cell.border.left.style, cell.border.right.style,
        cell.border.top.style, cell.border.bottom.style,
    )


def main():
    src_md5 = md5(TARGET)
    assert src_md5 == md5(BACKUP), "0.xlsx 与 19:12 备份不一致，先停下"

    wb = openpyxl.load_workbook(TARGET)
    ws = wb.active
    old_dims = (ws.max_row, ws.max_column)
    old_merges = {str(m) for m in ws.merged_cells.ranges}
    old_values = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                  for r in range(1, ws.max_row + 1)]
    old_styles = {addr: style_tuple(ws[addr]) for addr in CELL_FIXES}

    for addr, new in CELL_FIXES.items():
        ws[addr] = new

    for addr, pairs in NOTE_FIXES.items():
        cell = ws[addr]
        text = cell.value or ""
        for old, new in pairs:
            if text.count(old) != 1:
                raise SystemExit("%s 里找不到唯一匹配：%r（出现 %d 次）"
                                 % (addr, old, text.count(old)))
            text = text.replace(old, new, 1)
        cell.value = text

    os.makedirs(os.path.dirname(STAGE), exist_ok=True)
    if os.path.exists(STAGE):
        raise SystemExit("校验件已存在，先确认是否上一次残留：%s" % STAGE)
    wb.save(STAGE)

    # 重新载入校验件，逐格比对
    wb2 = openpyxl.load_workbook(STAGE)
    ws2 = wb2.active
    assert (ws2.max_row, ws2.max_column) == old_dims, "行列数变了"
    assert {str(m) for m in ws2.merged_cells.ranges} == old_merges, "合并区变了"

    diffs = []
    for r in range(1, old_dims[0] + 1):
        for c in range(1, old_dims[1] + 1):
            new = ws2.cell(r, c).value
            if new != old_values[r - 1][c - 1]:
                diffs.append("%s%d" % (openpyxl.utils.get_column_letter(c), r))

    expected = sorted(list(CELL_FIXES) + list(NOTE_FIXES))
    if sorted(diffs) != expected:
        raise SystemExit("差异格不符：实际 %s；预期 %s" % (sorted(diffs), expected))

    for addr in CELL_FIXES:
        assert style_tuple(ws2[addr]) == old_styles[addr], "%s 样式变了" % addr

    # 抽查第二批（43–55）未被触碰
    for addr in ("A209", "B237", "K266", "AA255", "Z261"):
        assert ws2[addr].value == old_values[ws2[addr].row - 1][ws2[addr].column - 1], \
            "%s 被误改" % addr

    if md5(TARGET) != src_md5:
        raise SystemExit("校验期间 0.xlsx 被别的进程改了，停止覆盖")
    shutil.copy2(STAGE, TARGET)
    print("0.xlsx 已修正；行列 %s；合并区 %d 个未变；改动 %d 格"
          % (old_dims, len(old_merges), len(diffs)))

    # JSON
    assert md5(JSON_PATH) == md5(JSON_BACKUP), "final_rows.json 与备份不一致"
    data = json.load(open(JSON_PATH, encoding="utf-8"))
    assert len(data) == 15 and [d["no"] for d in data] == list(range(28, 43))

    for doc in data:
        no = doc["no"]
        for old, new in JSON_NOTE_FIXES.get(no, []):
            if doc["notes"].count(old) != 1:
                raise SystemExit("JSON 文献%d notes 找不到唯一匹配：%r" % (no, old))
            doc["notes"] = doc["notes"].replace(old, new, 1)

    for doc in data:
        no = doc["no"]
        if no not in JSON_ROW_FIXES:
            continue
        rows = doc["rows"]
        if no == 30:
            assert rows[0]["sample"] == "1#"
            rows[0].update(JSON_ROW_FIXES[no])
        elif no == 32:
            assert len(rows) == 1
            rows[0].update(JSON_ROW_FIXES[no])
        elif no == 34:
            assert rows[0]["sample"].startswith("A29")
            rows[0].update(JSON_ROW_FIXES[no])
        elif no == 38:
            target = [r for r in rows if r["sample"] == "204Cu"]
            assert len(target) == 1
            target[0].update(JSON_ROW_FIXES[no])

    with open(JSON_PATH, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print("final_rows.json 已同步修正")


if __name__ == "__main__":
    main()
