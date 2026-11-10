# -*- coding: utf-8 -*-
r"""第12步（重建 1.xlsx）只读盘点：把 0.xlsx 各列写法、块级 Y/Z、Epit/参比/
温度/稳定时间/对电极/工艺/元素的所有不同取值落盘，供清洗脚本设计解析规则。

只写 备份\清洗1_20261109\audit_inputs.json，不碰 0.xlsx / 1.xlsx / temp。
用法：D:\Python\python.exe -X utf8 scripts\audit_1_inputs_20261109.py
"""
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "0.xlsx")
OUT_DIR = os.path.join(ROOT, "备份", "清洗1_20261109")
OUT = os.path.join(OUT_DIR, "audit_inputs.json")

C_ID, C_NAME = 1, 2
C_ELEM = list(range(3, 17))       # C..Sn
C_FE = 17
C_PROC, C_WE, C_RE, C_CE = 18, 19, 20, 21
C_TEMP, C_EPIT, C_STAB = 22, 23, 24
C_Y, C_Z = 25, 26


def norm(s):
    if s is None:
        return ""
    return unicodedata.normalize("NFKC", str(s)).strip()


def classify_epit(v):
    """粗分 Epit 形态，给设计 parse_epit 用。"""
    s = norm(v)
    if s == "":
        return "空"
    up = s.upper()
    if up.startswith("N/A") or up in ("NA", "NAN", "NONE", "-", "—"):
        return "N/A类"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return "数值"
    if re.fullmatch(r"[-−‑‒–—+]?\d+(?:\.\d+)?", s):
        return "纯数字文本"
    if "±" in s or "+/-" in s or "+-" in s:
        return "带±"
    if re.search(r"[<≤≤≦]", s):
        return "比较符"
    if re.search(r"mV|mVSCE|mVSCE|MILLIV", s, re.I):
        return "mV"
    if re.search(r"\bV\b|VSCE|V_SCE|Vvs", s):
        return "V"
    return "其它"


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Sheet1"]

    # 表头
    h1 = [ws.cell(1, c).value for c in range(1, 27)]
    h2 = [ws.cell(2, c).value for c in range(1, 27)]

    # 合并区里 A/Y/Z 的块首 -> 值（用于逐行回填）
    merge_first = {}
    for rng in ws.merged_cells.ranges:
        key = (rng.min_row, rng.min_col)
        merge_first.setdefault(key, []).append(str(rng))

    rows = []
    cur_id = None
    cur_y = None
    cur_z = None
    block_of_row = {}
    blocks = {}   # pid -> {y,z,first_row,count,notes_len}
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, C_ID).value
        if a not in (None, ""):
            try:
                cur_id = int(float(a))
            except (TypeError, ValueError):
                cur_id = a
        y = ws.cell(r, C_Y).value
        z = ws.cell(r, C_Z).value
        if y not in (None, ""):
            cur_y = norm(y)
        if z not in (None, ""):
            cur_z = str(z)
        name = ws.cell(r, C_NAME).value
        if name in (None, ""):
            continue
        block_of_row[r] = cur_id
        if cur_id not in blocks:
            blocks[cur_id] = {
                "first_row": r, "count": 0,
                "y": cur_y or "", "z": cur_z or "",
                "z_len": len(cur_z or ""),
            }
        blocks[cur_id]["count"] += 1
        rec = {
            "row": r, "pid": cur_id, "name": norm(name),
            "Y": cur_y or "", "Z": cur_z or "",
        }
        for c in C_ELEM + [C_FE, C_PROC, C_WE, C_RE, C_CE, C_TEMP, C_EPIT, C_STAB]:
            rec["c%d" % c] = ws.cell(r, c).value
        rows.append(rec)

    def uniq(col, transform=None):
        cnt = Counter()
        for rec in rows:
            v = rec["c%d" % col]
            s = transform(v) if transform else norm(v)
            cnt[s] += 1
        return cnt

    data = {
        "meta": {
            "src": SRC, "sheet": "Sheet1",
            "dims": [ws.max_row, ws.max_column],
            "merges": len(ws.merged_cells.ranges),
            "data_rows": len(rows), "blocks": len(blocks),
            "h1": [norm(x) for x in h1], "h2": [norm(x) for x in h2],
            "col_widths": {chr(64 + c): ws.column_dimensions[chr(64 + c)].width
                           for c in range(1, 27) if chr(64 + c) in ws.column_dimensions},
            "zoom": ws.sheet_view.zoomScale, "freeze": ws.freeze_panes,
        },
        "epit_forms": dict(uniq(C_EPIT)),
        "epit_class": dict(Counter(classify_epit(rec["c%d" % C_EPIT]) for rec in rows)),
        "re_forms": dict(uniq(C_RE)),
        "temp_forms": dict(uniq(C_TEMP)),
        "stab_forms": dict(uniq(C_STAB)),
        "ce_forms": dict(uniq(C_CE)),
        "we_forms": dict(uniq(C_WE)),
        "fe_forms": dict(uniq(C_FE)),
        "proc_forms": dict(uniq(C_PROC)),
        "elem_forms": {str(c): dict(uniq(c)) for c in C_ELEM},
        "blocks": {str(k): v for k, v in blocks.items()},
        "long_notes": {str(k): {"len": v["z_len"], "z": v["z"]}
                       for k, v in blocks.items() if v["z_len"] > 120},
        "merge_ranges": sorted(str(x) for x in ws.merged_cells.ranges),
    }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    print("audit ->", OUT, os.path.getsize(OUT), "B")
    print("dims", data["meta"]["dims"], "data_rows", len(rows), "blocks", len(blocks))
    print("h1", data["meta"]["h1"])
    print("h2", data["meta"]["h2"])
    print("epit_class", data["epit_class"])
    print("epit_forms", len(data["epit_forms"]), "re_forms", len(data["re_forms"]),
          "temp_forms", len(data["temp_forms"]), "stab_forms", len(data["stab_forms"]),
          "ce_forms", len(data["ce_forms"]), "proc_forms", len(data["proc_forms"]))
    print("long_notes", len(data["long_notes"]), "max", max((v["z_len"] for v in blocks.values()), default=0))


if __name__ == "__main__":
    main()
