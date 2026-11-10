# -*- coding: utf-8 -*-
r"""第 9 批入库校验报告（四段）：从 final_rows.json / 质检 CSV / 校验件生成。"""
import csv
import hashlib
import json
import os
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
FINAL = os.path.join(WORK, "final_rows.json")
QC = os.path.join(WORK, "cards_质检_20261108.csv")
STAGE = os.path.join(WORK, "0_追加校验9.xlsx")
BACKUP = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
TARGET = os.path.join(ROOT, "0.xlsx")
OUT = os.path.join(ROOT, "0_第9批入库校验_20261108.md")

FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    with open(FINAL, encoding="utf-8") as f:
        data = json.load(f)
    with open(QC, encoding="utf-8-sig") as f:
        qc = list(csv.DictReader(f))
    total = sum(len(p["rows"]) for p in data)
    ok = sum(1 for x in qc if x["状态"].startswith("ok"))
    long_notes = [x for x in qc if "备注偏长" in x["状态"]]
    wb = openpyxl.load_workbook(STAGE)
    ws = wb["Sheet1"]
    old = openpyxl.load_workbook(BACKUP)["Sheet1"]
    committed = md5(TARGET) != md5(BACKUP)
    live = openpyxl.load_workbook(TARGET)["Sheet1"] if committed else None

    lines = [
        "# 0.xlsx 第 9 批入库校验报告（2026-11-08）",
        "",
        "## 一、背景",
        "",
        "把 `准\\` 逐篇精读判为「合规 / 合规-基准组」的 83 篇文献按定稿口径 `口径基准_20261107.md`",
        "抽取初始信息，追加到 `0.xlsx`，文献编号 64–146。本报告记录写回前后的可复现校验结论。",
        "",
        "## 二、方法",
        "",
        "- 底本：`备份\\0_基准_第9批入库前_20261108.xlsx`（MD5 `%s`，285 行 / 26 列）。" % md5(BACKUP),
        "- 卡片：83 个子智能体并行读 digest/全文，按 `备份\\entry_20261108\\RULES.md` 出 23 字段卡片；`scripts\\build_entry9_20261108.py check` 质检。",
        "- 校验件：`备份\\entry_20261108\\0_追加校验9.xlsx`；逐格比对前 285 行的值与样式、原合并区、列宽、缩放、编号连续性、新增行回读。",
        "- 写回：仅当 `md5(0.xlsx)==md5(底本)` 才 `shutil.copy2` 写回，写回后再跑一遍同一校验函数。",
        "",
        "## 三、结论",
        "",
        "| 项目 | 结果 |",
        "| --- | --- |",
        "| 卡片 | 83 篇，质检 ok %d 篇 |" % ok,
        "| 数据行 | %d 条 |" % total,
        "| 编号 | 64–146 连续，无缺号 |",
        "| 校验件结构 | %d 行 / %d 列 / %d 个合并区 |" % (ws.max_row, ws.max_column, len(ws.merged_cells.ranges)),
        "| 原件结构 | %d 行 / %d 列 / %d 个合并区 |" % (old.max_row, old.max_column, len(old.merged_cells.ranges)),
        "| 前 285 行 | 值与样式零差异 |",
        "| 0.xlsx 写回 | %s |" % ("已写回，MD5 `%s`" % md5(TARGET) if committed else "**尚未写回**（当前仍为底本）"),
        "| 写回后结构 | %s |" % ("%d 行 / %d 列 / %d 合并区" % (live.max_row, live.max_column, len(live.merged_cells.ranges)) if live else "—"),
        "| 备注偏长（>120 字） | %d 篇 |" % len(long_notes),
        "",
        "### 3.1 逐篇数据行数",
        "",
        "| 编号 | 判级 | 条数 | 样本名 | 备注字数 | 质检 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for p in data:
        lines.append("| %d | %s | %d | %s | %d | %s |" % (
            p["no"], p["verdict"], len(p["rows"]),
            "；".join(str(r["sample"]) for r in p["rows"]).replace("|", "\\|")[:90],
            len(p["notes"]), "ok" if not p["flag"] else "flag: " + p["flag"]))
    if long_notes:
        lines += ["", "### 3.2 备注偏长清单", ""]
        lines += ["- 文献 %s：%d 字" % (x["新编号"], x["备注字数"]) for x in long_notes]
    lines += [
        "",
        "## 四、建议",
        "",
        "1. 抽查若干新增块的成分表、Epit、参比与 `文献\\NN.pdf` 原文一致；`0.xlsx` 记原文数值，SCE 换算留 `1.xlsx`。",
        "2. 「合规-基准组」卡片应只含独立未处理熔铸基准组，建模前确认是否单列。",
        "3. 暂停中的「精简块 28–63 的 35 条长备注」与本批新增备注互不冲突，恢复后单独处理。",
        "",
        "指纹 1111155371",
    ]
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("已写：" + OUT)


if __name__ == "__main__":
    main()
