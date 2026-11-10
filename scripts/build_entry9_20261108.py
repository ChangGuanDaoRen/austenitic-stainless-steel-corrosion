# -*- coding: utf-8 -*-
r"""第 9 批入库（准\ 83 篇 -> 0.xlsx，文献编号 64-146）的确定性构建器。

三个模式：
  check   读 cards\064.json..146.json，质检并生成 final_rows.json + 质检 CSV（不碰 xlsx）
  build   以 备份\0_基准_第9批入库前_20261108.xlsx 为底生成校验件并全量逐格校验
  commit  校验通过且 0.xlsx 未被改动时，写回 0.xlsx 并二次复验

硬约束：既有的 285 行 A–Z 一格不动；新增行照 26 列模板（当前 0.xlsx 实测为 26 列，无 AA 列）。
"""
import csv
import hashlib
import json
import os
import re
import shutil
import sys
from copy import copy

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
CARD_DIR = os.path.join(WORK, "cards")
BACKUP = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
TARGET = os.path.join(ROOT, "0.xlsx")
STAGE = os.path.join(WORK, "0_追加校验9.xlsx")
FINAL = os.path.join(WORK, "final_rows.json")
QC = os.path.join(WORK, "cards_质检_20261108.csv")

NCOL = 26
OLD_MAX_ROW = 285
TPL_NORMAL_ROW = 285
TPL_ANCHOR_ROW = 283
TPL_WRAP_ROW = 3
FIRST_NO, LAST_NO = 64, 146

FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]          # B..X = 2..24

NUM_INT = re.compile(r"^-?\d+$")
NUM_FLOAT = re.compile(r"^-?\d*\.\d+(?:[eE][+-]?\d+)?$|^-?\d+[eE][+-]?\d+$")


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


NUMERIC_FIELDS = set(["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
                      "Ti", "Nb", "Co", "Sn", "Fe", "Epit"])


def cellval(v, key):
    """照录原文；只有成分列与 Epit 的纯数字转数值型（与前段人工表一致）。
    样本名/工艺/电极/温度/稳定时间一律保留文本，避免 '304' 被写成 304。"""
    if v is None:
        return "未提及"
    t = str(v).strip()
    if not t:
        return "未提及"
    if key in NUMERIC_FIELDS:
        if NUM_INT.match(t):
            try:
                return int(t)
            except ValueError:
                return t
        if NUM_FLOAT.match(t):
            try:
                return float(t)
            except ValueError:
                return t
    return t


def load_cards():
    with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    by_no = {m["newno"]: m for m in manifest}

    data, qc = [], []
    for no in range(FIRST_NO, LAST_NO + 1):
        idx = "%03d" % no
        path = os.path.join(CARD_DIR, idx + ".json")
        item = {"新编号": no, "卡片": idx + ".json", "原文件名": by_no[no]["file"],
                "判级": by_no[no]["verdict"], "状态": "", "行数": 0,
                "题录字数": 0, "备注字数": 0, "备注": "", "问题": ""}
        if not os.path.exists(path):
            item["状态"] = "缺卡片"
            item["问题"] = "card file missing"
            qc.append(item)
            continue
        try:
            with open(path, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            item["状态"] = "JSON 解析失败"
            item["问题"] = "%s: %s" % (type(e).__name__, e)
            qc.append(item)
            continue

        problems = []
        if c.get("newno") not in (no, str(no)):
            problems.append("newno=%r 与 %d 不符" % (c.get("newno"), no))
        if not str(c.get("file") or "").strip():
            problems.append("file 空")
        elif str(c.get("file")).strip() != by_no[no]["file"]:
            problems.append("file 与台账不符：%r" % c.get("file"))
        citation = str(c.get("citation") or "").strip()
        notes = str(c.get("notes") or "").strip()
        if not citation:
            problems.append("citation 空")
        if not notes:
            problems.append("notes 空")
        rows = c.get("rows")
        if not isinstance(rows, list) or len(rows) < 1:
            problems.append("rows 为空")
            rows = []
        norm_rows = []
        for i, r in enumerate(rows):
            if not isinstance(r, dict):
                problems.append("row %d 非对象" % i)
                continue
            missing = [k for k in FIELDS if k not in r or r[k] is None or str(r[k]).strip() == ""]
            if missing:
                problems.append("row %d 缺 %s" % (i, ",".join(missing)))
            extra = [k for k in r if k not in FIELDS]
            if extra:
                problems.append("row %d 多字段 %s（已忽略）" % (i, ",".join(extra)))
            norm_rows.append({k: cellval(r.get(k), k) for k in FIELDS})
        flag = str(c.get("flag") or "").strip()

        item["行数"] = len(norm_rows)
        item["题录字数"] = len(citation)
        item["备注字数"] = len(notes)
        item["备注"] = notes.replace("\n", " ")
        if flag:
            item["状态"] = "flag（占位行，待补证）"
            item["问题"] = flag
        elif problems:
            item["状态"] = "需修"
            item["问题"] = "；".join(problems)
        else:
            item["状态"] = "ok"
        if len(notes) > 120:
            item["状态"] += "／备注偏长"
            item["问题"] = (item["问题"] + "；" if item["问题"] else "") + "备注 %d 字" % len(notes)

        qc.append(item)
        data.append({
            "no": no, "idx": idx, "file": by_no[no]["file"], "verdict": by_no[no]["verdict"],
            "citation": citation, "notes": notes, "flag": flag, "rows": norm_rows,
        })

    with open(QC, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(qc[0].keys()))
        w.writeheader()
        w.writerows(qc)
    return data, qc


def check():
    data, qc = load_cards()
    bad = [x for x in qc if not x["状态"].startswith("ok")]
    flags = [x for x in qc if x["状态"].startswith("flag")]
    total = sum(len(p["rows"]) for p in data)
    print("卡片 %d 篇 / 数据 %d 条 / 编号 %d-%d" % (len(data), total, FIRST_NO, LAST_NO))
    print("状态分布：", {s: sum(1 for x in qc if x["状态"] == s) for s in sorted(set(x["状态"] for x in qc))})
    print("质检表：" + QC)
    if bad:
        print("非 ok / 非 flag 共 %d 篇：" % len(bad))
        for x in bad:
            print("  -", x["新编号"], x["原文件名"][:45], "|", x["状态"], "|", x["问题"][:160])
    if flags:
        print("flag 占位 %d 篇：" % len(flags))
        for x in flags:
            print("  -", x["新编号"], x["原文件名"][:45], "|", x["问题"][:120])
    with open(FINAL, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print("已写：" + FINAL)
    empty = [x for x in qc if x["行数"] < 1]
    blocking = [x for x in qc if not x["状态"].startswith("ok")] + empty
    if blocking:
        print("质检未通过：%d 篇需修/缺卡片/flag/零数据行，禁止进入 build："
              % len(blocking))
        for x in blocking[:15]:
            print("  -", x["新编号"], x["原文件名"][:45], "|", x["状态"], "|", x["问题"][:120])
        raise SystemExit(2)
    print("质检通过：83 篇 / %d 条数据，全部可入库" % total)
    return data


def style_sig(cell):
    f, a, b = cell.font, cell.alignment, cell.border
    return (f.name, f.sz, f.b, f.i,
            a.horizontal, a.vertical, a.wrap_text,
            cell.number_format,
            b.left.style, b.right.style, b.top.style, b.bottom.style)


def build(data):
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if ws.max_row != OLD_MAX_ROW or ws.max_column != NCOL:
        raise SystemExit("原件结构变了：%d 行 / %d 列（应为 %d / %d）"
                         % (ws.max_row, ws.max_column, OLD_MAX_ROW, NCOL))

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
            ws.cell(r, 26).value = p["notes"] if i == 0 else None
            if i == 0:
                for c in (1, 25, 26):
                    ws.cell(r, c)._style = copy(tpl_anchor[c])
                ws.cell(r, 25).alignment = copy(align_y)
                ws.cell(r, 26).alignment = copy(align_z)
        if n > 1:
            for c in (1, 25, 26):
                ws.merge_cells(start_row=start, start_column=c, end_row=end, end_column=c)
        blocks.append({"no": p["no"], "start": start, "end": end, "n": n,
                       "citation": p["citation"], "notes": p["notes"],
                       "samples": [x["sample"] for x in p["rows"]]})
    wb.save(STAGE)
    return blocks


def verify(data, blocks, path):
    problems = []
    old = openpyxl.load_workbook(BACKUP)
    new = openpyxl.load_workbook(path)
    if old.sheetnames != new.sheetnames:
        problems.append("sheet 列表变了：%r -> %r" % (old.sheetnames, new.sheetnames))
    o, n = old["Sheet1"], new["Sheet1"]

    exp_rows = OLD_MAX_ROW + sum(len(p["rows"]) for p in data)
    if n.max_row != exp_rows:
        problems.append("成品行数不对：%d（应为 %d）" % (n.max_row, exp_rows))
    if n.max_column != NCOL:
        problems.append("成品列数不对：%d（应为 %d）" % (n.max_column, NCOL))

    changed = 0
    for rr in range(1, OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            a, b = o.cell(rr, c), n.cell(rr, c)
            if a.value != b.value or style_sig(a) != style_sig(b):
                changed += 1
                if changed <= 5:
                    problems.append("原有格被改：r%d c%d %r -> %r" % (rr, c, a.value, b.value))
    if changed:
        problems.append("前 %d 行共有 %d 处差异" % (OLD_MAX_ROW, changed))

    om = set(str(x) for x in o.merged_cells.ranges)
    nm = set(str(x) for x in n.merged_cells.ranges)
    if not om.issubset(nm):
        problems.append("原有合并区丢失：%s" % sorted(om - nm)[:5])
    multi = sum(1 for b in blocks if b["n"] > 1)
    if len(nm) != len(om) + 3 * multi:
        problems.append("新增合并区数量不对：%d（应为 %d）" % (len(nm), len(om) + 3 * multi))

    ow = {k: v.width for k, v in o.column_dimensions.items()}
    nw = {k: v.width for k, v in n.column_dimensions.items()}
    if ow != nw:
        problems.append("列宽变了")
    if o.sheet_view.zoomScale != n.sheet_view.zoomScale:
        problems.append("缩放比例变了")

    nos = [n.cell(b["start"], 1).value for b in blocks]
    if nos != list(range(FIRST_NO, LAST_NO + 1)):
        problems.append("文献编号不连续：%r" % nos[:5])
    for p, b in zip(data, blocks):
        s, e, cnt = b["start"], b["end"], b["n"]
        if n.cell(s, 25).value != p["citation"]:
            problems.append("文献 %d 题录写入不符" % p["no"])
        if n.cell(s, 26).value != p["notes"]:
            problems.append("文献 %d 备注写入不符" % p["no"])
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
                    problems.append("文献 %d / %s 字段 %s 写入不符（r%d c%d）"
                                    % (p["no"], row["sample"], key, rr, 2 + j))
                    break
            if n.cell(rr, 2).font.name != "宋体" or n.cell(rr, 2).font.sz != 11:
                problems.append("文献 %d / r%d 字体不是宋体 11" % (p["no"], rr))
            if i == 0 and n.cell(rr, 25).alignment.wrap_text is not True:
                problems.append("文献 %d / r%d 题录未自动换行" % (p["no"], rr))
    return problems


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    if mode == "check":
        check()
        return

    if not os.path.exists(FINAL):
        raise SystemExit("缺 final_rows.json，先跑 check")
    with open(FINAL, encoding="utf-8") as f:
        data = json.load(f)
    if [p["no"] for p in data] != list(range(FIRST_NO, LAST_NO + 1)):
        raise SystemExit("final_rows.json 编号不连续（疑似质检未通过时留下的旧文件），重跑 check")
    bad = [p["no"] for p in data
           if p.get("flag") or not p.get("citation") or not p.get("notes")
           or not isinstance(p.get("rows"), list) or len(p["rows"]) < 1
           or any(any(k not in r or r[k] is None or str(r[k]).strip() == "" for k in FIELDS)
                  for r in p["rows"])]
    if bad:
        raise SystemExit("final_rows.json 有 %d 篇不合格（%s），重跑 check 并补齐卡片"
                         % (len(bad), bad[:10]))
    total = sum(len(p["rows"]) for p in data)

    if mode == "build":
        if md5(TARGET) != md5(BACKUP):
            raise SystemExit("0.xlsx 与第9批入库前备份不一致，停止（先人工核对）")
        blocks = build(data)
        problems = verify(data, blocks, STAGE)
        print("数据：%d 篇 / %d 条，编号 %d-%d，写入行 %d-%d"
              % (len(data), total, FIRST_NO, LAST_NO, OLD_MAX_ROW + 1,
                 OLD_MAX_ROW + total))
        print("校验件：" + STAGE)
        if problems:
            print("校验未通过，共 %d 项：" % len(problems))
            for x in problems[:40]:
                print("  -", x)
            raise SystemExit(2)
        _nm = len(openpyxl.load_workbook(STAGE)["Sheet1"].merged_cells.ranges)
        _om = len(openpyxl.load_workbook(BACKUP)["Sheet1"].merged_cells.ranges)
        print("校验通过：前 %d 行零差异 / 原有 %d 个合并区全保留 / 新增 %d 个 / 列宽与缩放不变 / %d 行逐格回读一致"
              % (OLD_MAX_ROW, _om, _nm - _om, total))
        print("[dry] 未改 0.xlsx；确认无误后跑 commit")
        return

    if mode == "commit":
        if md5(TARGET) != md5(BACKUP):
            raise SystemExit("提交前发现 0.xlsx 被改动，停止")
        if not os.path.exists(STAGE):
            raise SystemExit("缺校验件，先跑 build")
        blocks = build(data)   # 重建，保证 blocks 与 STAGE 一致
        problems = verify(data, blocks, STAGE)
        if problems:
            print("校验件复验失败：")
            for x in problems[:20]:
                print("  -", x)
            raise SystemExit(2)
        shutil.copy2(STAGE, TARGET)
        problems2 = verify(data, blocks, TARGET)
        if problems2:
            print("写回后校验失败（原件仍在 备份\\ 可恢复）：")
            for x in problems2[:20]:
                print("  -", x)
            raise SystemExit(3)
        print("[commit] 已写回 0.xlsx，MD5=%s；%d 篇 / %d 条；写回后复验通过"
              % (md5(TARGET), len(data), total))
        return

    raise SystemExit("未知模式：%s（应为 check / build / commit）" % mode)


if __name__ == "__main__":
    main()
