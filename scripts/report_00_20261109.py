#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成 00.xlsx 的配套台账与校验报告（只读 00.xlsx / 0.xlsx / json，写 csv+md）。"""
import csv
import glob
import json
import os
from collections import Counter, defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JDIR = os.path.join(ROOT, "备份", "00_录入_20261109", "json")
OUTDIR = os.path.join(ROOT, "备份", "00_录入_20261109")

ELEMENTS = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni", "Ti", "Nb", "Co", "Sn", "Fe"]


def load_json():
    out = []
    for f in sorted(glob.glob(os.path.join(JDIR, "*.json"))):
        out.append(json.load(open(f, encoding="utf-8")))
    return out


def blocks_of(ws, col=1):
    cnt = defaultdict(int)
    cur = None
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, col).value
        if a is not None:
            cur = int(a)
        if cur is not None:
            cnt[cur] += 1
    return cnt


def main():
    js = load_json()
    wb0 = openpyxl.load_workbook(os.path.join(ROOT, "00.xlsx"))
    ws0 = wb0["Sheet1"]
    wb1 = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"))
    ws1 = wb1["Sheet1"]

    c00 = blocks_of(ws0)
    c0 = blocks_of(ws1)

    # 1 跳过台账
    skip = [d for d in js if not d.get("conforms")]
    with open(os.path.join(OUTDIR, "跳过台账_20261109.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["文献编号", "文件", "跳过原因"])
        for d in skip:
            w.writerow([d["no"], d["file"], d.get("exclude_reason", "")])

    # 2 与 0.xlsx 差异
    allno = sorted(set(c0) | set(c00))
    diffs = [(n, c0.get(n, 0), c00.get(n, 0), c00.get(n, 0) - c0.get(n, 0)) for n in allno]
    diffs.sort(key=lambda x: (-abs(x[3]), x[0]))
    with open(os.path.join(OUTDIR, "00_与0差异对照_20261109.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["文献编号", "0.xlsx行数", "00.xlsx行数", "差值(00-0)"])
        for row in diffs:
            if row[3] != 0:
                w.writerow(row)

    # 3 统计
    rows = []
    for r in range(3, ws0.max_row + 1):
        rows.append([ws0.cell(r, c).value for c in range(1, 27)])
    total = len(rows)
    epit_num = sum(1 for r in rows if isinstance(r[22], (int, float)))
    epit_na = sum(1 for r in rows if isinstance(r[22], str) and "待读图" in r[22])
    epit_range = sum(1 for r in rows if isinstance(r[22], str) and ("±" in r[22] or "-" in r[22]) and "N/A" not in r[22])
    re_miss = sum(1 for r in rows if r[19] == "未提及")
    temp_miss = sum(1 for r in rows if r[21] == "未提及")
    stab_miss = sum(1 for r in rows if r[23] == "未提及")
    fe_miss = sum(1 for r in rows if r[16] == "未提及")
    two_step = sum(1 for r in rows if isinstance(r[25], str) and "两步" in r[25])
    doubt = sum(1 for r in rows if isinstance(r[25], str) and ("存疑" in r[25] or "待核" in r[25] or "人工" in r[25]))
    sea = sum(1 for r in rows if isinstance(r[25], str) and ("海水" in r[25]))
    na_no = sum(1 for r in rows if isinstance(r[22], str) and "无点蚀" in r[22])

    reason_cnt = Counter()
    for d in skip:
        t = d.get("exclude_reason", "")
        for k in ["增材", "焊接", "表面", "双相", "铁素体", "无 Epit", "无点蚀", "介质", "无成分", "标准", "无数据"]:
            if k in t:
                reason_cnt[k] += 1
                break
        else:
            reason_cnt["其它"] += 1

    md = []
    md.append("# 00.xlsx 重录校验报告（2026-11-09）\n")
    md.append("## 一、背景\n")
    md.append("用户判断前序录入可能有问题，决定暂停后续推进、重新录入。本轮按 `文献\\01.pdf`–`177.pdf` 逐篇重新提取，"
              "列结构与写法对齐 `0.xlsx`（26 列、两行表头、宋体 11、细边框、居中、A 列按文献块竖合并），"
              "口径以师兄微信原话与 `口径基准_20261107.md` 为准。用户 11-09 17:54 裁定：图上 Epit 记 `N/A（待读图，图X）`；不符合口径的整篇跳过。\n")
    md.append("## 二、方法\n")
    md.append("1. 先读 `0.xlsx` 前 20 篇，提取列定义与写法惯例（`未提及` 填充、`余量` 表示 Fe、工艺按「成形方式，热处理制度，冷却方式」全角逗号分隔、Y 只写块首行）。")
    md.append("2. 写 `备份\\00_录入_20261109\\RULES_00.md`：四条收录口径、24 个字段定义、6 步录入流程、JSON 输出格式、禁止事项。")
    md.append("3. 写 `scripts\\pdf_read_00.py`（只读 PDF）：brief / grep / pages / tables 四种模式，供子智能体按需定位成分表与电化学段。")
    md.append("4. 以 01、03 两篇做试点，与 `0.xlsx` 对账（成分、工艺、行数一致）后再批量派发。")
    md.append("5. 177 篇拆成 20 批（每批 9 篇，末批 6 篇），20 个子智能体并行提取，每篇写一个 JSON 到 `备份\\00_录入_20261109\\json\\NNN.json`。")
    md.append("6. `scripts\\build_00_20261109.py` 合并生成 `00.xlsx`；`scripts\\report_00_20261109.py` 生成本报告与两份台账。"
              "Epit 列做了单位归一（mV→V，含区间/±写法），参比电极一律不做换算，改动逐行记 `00_Epit单位归一_20261109.csv`。")
    md.append("7. 20:20 补做回填：`scripts\\fill_from_0_20261109.py` 按「文献编号 + 样本名完全一致 + 参比一致」从 `0.xlsx` 搬 Epit，回填 50 行，审计 `读图回填来源_20261109.csv`；"
              "剩余读图清单 `读图待办_20261109.csv`，作业说明 `读图作业说明_20261109.md`。\n")
    md.append("## 三、结论\n")
    md.append(f"- **规模**：177 篇全部处理；`00.xlsx` 收录 **{len(c00)} 篇 / {total} 行 / 26 列**；跳过 **{len(skip)} 篇**。`0.xlsx` 为 {len(c0)} 块 / {sum(c0.values())} 条。")
    md.append(f"- **Epit 覆盖**：数值 {epit_num} 行（{epit_num/total*100:.1f}%）、`N/A（待读图）` {epit_na} 行（{epit_na/total*100:.1f}%）、"
              f"`N/A（无点蚀）` {na_no} 行、区间/± 写法 {epit_range} 行。数值范围 {min(r[22] for r in rows if isinstance(r[22],(int,float)))} ~ {max(r[22] for r in rows if isinstance(r[22],(int,float)))} V。")
    md.append(f"- **字段缺失**：参比电极 `未提及` {re_miss} 行、腐蚀温度 `未提及` {temp_miss} 行、稳定时间 `未提及` {stab_miss} 行、Fe `未提及` {fe_miss} 行。")
    md.append("- **`0.xlsx` 量级问题**：`0.xlsx` 的 Epit 列有 50 格记的是 mV 数字（最大 603、最小 −169），与同列 V 值混排。回填时按 `|值|>3` 判为 mV 折算，若直接拿 `0.xlsx` 建模会得到差 1000 倍的结果。")
    md.append(f"- **特殊标注**：备注含「两步」{two_step} 行、含「存疑/待核/人工」{doubt} 行、含「海水」{sea} 行。")
    md.append(f"- **跳过原因分布**（按首命中）：{dict(reason_cnt)}。逐篇清单见 `跳过台账_20261109.csv`。")
    md.append(f"- **与 `0.xlsx` 的差异**：{sum(1 for x in diffs if x[3] != 0)} 篇行数不同，明细见 `00_与0差异对照_20261109.csv`；差异最大的前 12 篇：")
    md.append("")
    md.append("| 文献 | 0.xlsx 行数 | 00.xlsx 行数 | 差值 |")
    md.append("|---|---|---|---|")
    for n, a, b, d in [x for x in diffs if x[3] != 0][:12]:
        md.append(f"| {n} | {a} | {b} | {d:+d} |")
    md.append("")
    md.append("## 四、建议\n")
    md.append("1. **先确认跳过名单**：26 篇整篇未录，其中增材/焊接/表面改性/双相/介质不符属口径判定，建议抽 5 篇复核 `跳过台账_20261109.csv`。")
    md.append(f"2. **优先补图**：仍有 {epit_na} 行 Epit 只在图上，是本表最大缺口。已出 `读图待办_20261109.csv`（177 行 / 37 篇，按三批拆分）与 `读图作业说明_20261109.md`，图号已在 W 列写明。")
    md.append("3. **重点复核不可读件**：文献 21、27 为 CNKI 字体、数字层丢失；22 为扫描件；08 为无文本层、由 OCR 重建成分表。这 4 篇的数值建议对照纸本或渲染页人工确认。")
    md.append("4. **差异大的篇目逐篇看**：文献 5、19、22、48、79、83、87、89 在 `0.xlsx` 有数据而 `00.xlsx` 为空（多为无 Epit、焊接、渗氮、钝化处理）；文献 10、45、77、131、162、177 行数大幅增加（旧表只收了部分试样）。这些是本次重录最值得核对的地方。")
    md.append("5. **后续清洗另起**：`00.xlsx` 只记原文初始信息，`1.xlsx` 的换算、补 0、参比归一等步骤建议基于 00 重新生成，不与旧 `1.xlsx` 混用。")
    md.append("")
    md.append("## 五、数据来源\n")
    md.append("`文献\\01.pdf`–`177.pdf`（2313 页，文本层 6.61 M 字符）、`0.xlsx`、`备份\\00_录入_20261109\\json\\`（177 个 JSON）、"
              "`备份\\00_录入_20261109\\RULES_00.md`、`scripts\\pdf_read_00.py`、`scripts\\build_00_20261109.py`。`0.xlsx`、`1.xlsx`、`文献\\` 本轮未改。")

    with open(os.path.join(OUTDIR, "00_录入校验报告_20261109.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(md))

    print("跳过", len(skip), "收录", len(c00), "行", total)
    print("Epit 数值", epit_num, "待读图", epit_na, "无点蚀", na_no, "区间±", epit_range)
    print("参比未提及", re_miss, "温度未提及", temp_miss, "稳定未提及", stab_miss, "Fe未提及", fe_miss)
    print("两步", two_step, "存疑", doubt, "海水", sea)
    print("差异篇数", sum(1 for x in diffs if x[3] != 0))


if __name__ == "__main__":
    main()
