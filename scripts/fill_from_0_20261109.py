#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""用 0.xlsx 已有的 Epit 记录回填 00.xlsx 中「N/A（待读图）」的行。

硬约束（避免瞎填）：
1. 只认「文献编号相同 + 样本名（NFKC/去空格/小写）完全一致」；
2. 0.xlsx 该行 Epit 必须是数值；
3. 两边参比电极必须一致（任一侧为未提及/空才放行）；
4. 0.xlsx 的 Epit 列本身混着 mV 与 V（|值|>3 视为 mV），统一折成 V 写入；
5. 只改 JSON 里的 Epit 与备注，其他列不动；改动逐行写审计 CSV。

用法：python fill_from_0_20261109.py dry|apply
"""
import csv
import json
import os
import re
import sys
import zipfile

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JDIR = os.path.join(ROOT, "备份", "00_录入_20261109", "json")
OUTDIR = os.path.join(ROOT, "备份", "00_录入_20261109")


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("（", "(").replace("）", ")").replace("　", " ")
    return re.sub(r"\s+", "", s).lower()


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def collect():
    ws = openpyxl.load_workbook(os.path.join(ROOT, "00.xlsx"))["Sheet1"]
    todo, cur, idx = [], None, 0
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is not None:
            cur, idx = int(a), 0
        v = ws.cell(r, 23).value
        if isinstance(v, str) and "待读图" in v:
            todo.append(dict(no=cur, idx=idx, xrow=r, sample=ws.cell(r, 2).value,
                             re=ws.cell(r, 20).value, temp=ws.cell(r, 22).value, fig=v))
        idx += 1
    return todo


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    w0 = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"))["Sheet1"]
    rows0, cur = {}, None
    for r in range(3, w0.max_row + 1):
        a = w0.cell(r, 1).value
        if a is not None:
            cur = int(a)
        rows0.setdefault(cur, []).append(dict(
            row=r, sample=w0.cell(r, 2).value, epit=w0.cell(r, 23).value,
            re=w0.cell(r, 20).value, temp=w0.cell(r, 22).value))

    todo = collect()
    fills, pend = [], []
    for t in todo:
        cands = [c for c in rows0.get(t["no"], []) if is_num(c["epit"])]
        if not cands:
            pend.append((t, "0.xlsx 该文献也无 Epit 数值"))
            continue
        ex = [c for c in cands if norm(c["sample"]) == norm(t["sample"])]
        if not ex:
            pend.append((t, "0.xlsx 该文献有值但样本名不一致"))
            continue
        if len(ex) > 1:
            tt = [c for c in ex if norm(c["temp"]) == norm(t["temp"]) and norm(t["temp"]) not in ("", "未提及")]
            vals = {round(float(c["epit"]), 4) for c in ex}
            if len(tt) == 1:
                ex = tt
            elif len(vals) == 1:
                ex = ex[:1]
            else:
                pend.append((t, "0.xlsx 同名多行且条件无法区分"))
                continue
        c = ex[0]
        if (norm(t["re"]) != norm(c["re"])) and norm(t["re"]) not in ("", "未提及") and norm(c["re"]) not in ("", "未提及"):
            pend.append((t, f"参比不一致（00={t['re']}，0={c['re']}）"))
            continue
        raw = float(c["epit"])
        unit = "mV" if abs(raw) > 3 else "V"
        val = round(raw / 1000, 4) if unit == "mV" else raw
        fills.append((t, c, raw, unit, val))

    print(f"待读图 {len(todo)} 行；可回填 {len(fills)} 行；留待读图 {len(pend)} 行")

    if mode == "dry":
        for t, c, raw, unit, val in fills[:50]:
            print(f"  文献{t['no']} 行{t['xrow']} {t['sample']} ← 0行{c['row']} {raw} {unit} → {val}")
        for t, why in pend[:5]:
            print(f"  [留] 文献{t['no']} {t['sample']}：{why}")
        return

    # 备份
    import shutil
    bak = os.path.join(OUTDIR, "00_填图前_20261109_2005.xlsx")
    if not os.path.exists(bak):
        shutil.copy2(os.path.join(ROOT, "00.xlsx"), bak)
    zpath = os.path.join(OUTDIR, "json_填图前_20261109_2005.zip")
    if not os.path.exists(zpath):
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for fn in sorted(os.listdir(JDIR)):
                z.write(os.path.join(JDIR, fn), fn)

    # 写回 JSON
    cache = {}
    for t, c, raw, unit, val in fills:
        f = os.path.join(JDIR, f"{t['no']:03d}.json")
        d = cache.get(f) or json.load(open(f, encoding="utf-8"))
        row = d["rows"][t["idx"]]
        note = row.get("备注", "")
        add = f"；Epit 取自 0.xlsx 同名样本记录值（原文 {raw:g} {unit}，参比一致），00 重录时该值仅在图中未重读"
        row["Epit"] = val
        row["备注"] = (note + add).strip("；")
        d["_epit_from_0"] = True
        cache[f] = d
    for f, d in cache.items():
        json.dump(d, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    # 审计
    audit = os.path.join(OUTDIR, "读图回填来源_20261109.csv")
    with open(audit, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["文献编号", "00.xlsx行号", "样本", "0.xlsx行号", "0原值", "0单位", "写入Epit(V)", "参比", "图号"])
        for t, c, raw, unit, val in fills:
            w.writerow([t["no"], t["xrow"], t["sample"], c["row"], f"{raw:g}", unit, val, t["re"], t["fig"]])

    # 剩余待办
    todo_csv = os.path.join(OUTDIR, "读图待办_20261109.csv")
    ws = openpyxl.load_workbook(os.path.join(ROOT, "00.xlsx"))["Sheet1"]
    with open(todo_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["文献编号", "00.xlsx行号", "样本", "图号", "参比", "温度", "备注线索",
                    "Epit(V)【填这里】", "不确定度±(V)", "取点说明", "状态"])
        for t, why in pend:
            w.writerow([t["no"], t["xrow"], t["sample"], t["fig"], t["re"], t["temp"],
                        str(ws.cell(t["xrow"], 26).value or "")[:120], "", "", "", why])
    print(f"[OK] 回填 {len(fills)} 行 → {audit}")
    print(f"[OK] 待办 {len(pend)} 行 → {todo_csv}")


if __name__ == "__main__":
    main()
