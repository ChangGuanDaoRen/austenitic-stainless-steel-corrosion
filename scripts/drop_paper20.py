# -*- coding: utf-8 -*-
"""剔除文献 20（原文无元素成分数据）：只改清洗版新表，不改原件、不动 00_原始数据 页。"""
import os
import shutil
import sys
import time

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

P = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版.xlsx"
DROP = 20
wb = openpyxl.load_workbook(P)

removed = {}
for name in list(wb.sheetnames):
    if name in ("00_原始数据", "说明"):
        continue
    ws = wb[name]
    hits = [r for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=1).value == DROP]
    if hits:
        removed[name] = [(r, ws.cell(row=r, column=2).value) for r in hits]
        for r in sorted(hits, reverse=True):
            ws.delete_rows(r, 1)
    # 删完后若没有数据行，整页移除（表头保留无意义）
    has_data = any(ws.cell(row=r, column=2).value not in (None, "") for r in range(3, ws.max_row + 1))
    if not has_data and name not in ("01_清洗总表",):
        wb.remove(ws)
        removed.setdefault("_empty_sheets", []).append(name)

doc = wb["说明"]
start = doc.max_row + 2
for i, t in enumerate([
    "六、文献20 整篇剔除（追加）",
    "剔除原因：20.pdf（付海成. 新型奥氏体不锈钢组织性能研究[D]. 燕山大学, 2022）原文未给出元素成分数值——316LN 一条成分原文未列，4 条新成分试样只写“以 Z3CN20-09M 为基础，调整 Cr、Ni 配比”，无具体数值，无法计算 Fe。",
    "处理方式：5 条记录从 01_清洗总表 及所属工艺表整条删除（121 → 116 条）；原表中因整篇只有这些行而变空的“6_熔炼铸锭”“8_锻造”两页一并移除。",
    "未动原件：根目录原始 xlsx 一字未改；文献20 的 5 条原始记录仍完整保留在本文件最后一页 00_原始数据 中，随时可查。",
]):
    doc.cell(row=start + i, column=1).value = t

tmp = os.path.join(os.path.dirname(P), "_tmp_drop20.xlsx")
wb.save(tmp)
for i in range(8):
    try:
        shutil.copyfile(tmp, P)
        os.remove(tmp)
        print("已写入正式件")
        break
    except PermissionError:
        print(f"文件被占用，{i*10}s 后重试…")
        time.sleep(10)
else:
    print("写入失败，临时文件保留:", tmp)

for k, v in removed.items():
    print(k, "->", v)
