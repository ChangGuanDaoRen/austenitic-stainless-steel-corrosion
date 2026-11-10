#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成《读图作业说明_20260909.md》：把剩余待读图行按文献整理成可执行清单。"""
import csv
import os
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "备份", "00_录入_20261109")
TODO = os.path.join(OUT, "读图待办_20261109.csv")
FILL = os.path.join(OUT, "读图回填来源_20261109.csv")

rows = list(csv.DictReader(open(TODO, encoding="utf-8-sig")))
fills = list(csv.DictReader(open(FILL, encoding="utf-8-sig")))
by = defaultdict(list)
for r in rows:
    by[int(r["文献编号"])].append(r)

n_doc = len(by)
re_miss = sum(1 for r in rows if r["参比"] == "未提及")
t_miss = sum(1 for r in rows if r["温度"] == "未提及")
name_bad = sum(1 for r in rows if "样本名不一致" in r["状态"])
fig = Counter(r["图号"].replace("N/A（待读图，", "").rstrip("）") for r in rows)

L = []
L.append("# 读图作业说明（2026-11-09 20:20）\n")
L.append("## 一、背景\n")
L.append("`00.xlsx` 是 11-09 按 `文献\\01.pdf`–`177.pdf` 重录的数据件（151 篇 / 743 行）。"
         "重录时按你的裁定：点蚀电位 Epit 只在图上出现、正文和表格都没有数值的，先记 `N/A（待读图，图X）`，不猜值。"
         "这类行原有 **227 行**。\n")
L.append("本轮先从 `0.xlsx` 里把已有记录搬过来：按「文献编号 + 样本名完全一致 + 参比一致」逐行比对，"
         f"成功回填 **{len(fills)} 行**，回填清单见 `读图回填来源_20261109.csv`（含 0.xlsx 原行号与原始数值，可逐条核对）。\n")
L.append(f"剩下的 **{len(rows)} 行、分布在 {n_doc} 篇文献** 是 `0.xlsx` 也拿不到值的："
         f"其中 {sum(1 for r in rows if '无 Epit 数值' in r['状态'])} 行是 `0.xlsx` 那篇本身也没记数值，"
         f"{name_bad} 行是 `0.xlsx` 存了值但样本名和这次重录对不上（不是同一批试样，硬凑会张冠李戴）。"
         "这 177 行只能靠看图取点。\n")
L.append("## 二、方法\n")
L.append("### 2.1 你要做的一件事\n")
L.append("打开 `读图待办_20261109.csv`，对每行：打开对应 PDF → 找到「图号」那一列写的图 → 在阳极极化曲线上读出该试样的点蚀电位 → "
         "填进第 8 列 `Epit(V)`，第 9 列写不确定度，第 10 列写一句取点说明。填完把 CSV 发回，我用脚本合并进 `00.xlsx` 并重算校验。\n")
L.append("### 2.2 怎么在一篇 PDF 里快速定位那张图\n")
L.append("1. 用 PDF 阅读器打开 `文献\\NN.pdf`（NN 就是 CSV 第 1 列文献编号，两位数补零，如 3 → `03.pdf`，131 → `131.pdf`）。")
L.append("2. 按 `Ctrl+F` 搜 `Fig.` 加上图号（如 `Fig.6`），或搜 `图6`；中文/日文文献搜 `図`。")
L.append("3. 图注一般在图的下方，先看图注确认横纵轴与参比标注，再读点。\n")
L.append("### 2.3 取点规则（和库里其他行保持一致）\n")
L.append("- **Epit 取阳极极化曲线上电流密度突增的拐点**，不做 Tafel 外推。原文若给了判据（如 `10 μA/cm²`、`100 μA/cm²`、`Ec`、`Eb`、`Ebreakdown`、`E_pit`），按原文判据取，并在「取点说明」里写清用的是哪个判据。")
L.append("- **曲线有多条时**按 CSV 的「样本」列去找图例对应的那条；图例只有编号没有材料名时，按图的排列顺序 + 正文描述判断，判断不了就留空并在「取点说明」写「图中无法区分」。")
L.append("- **单位**：图轴若为 mV，填 V（例如读到 350 mV 就填 0.35）；已有 500 行统一是 V。")
L.append("- **参比不换算**：图注写的参比如果和 CSV「参比」列一致，直接填；如果不一致（例如图上是 Ag/AgCl、表里记 SCE），**照图上的数值填、不要自己加减换算**，在「取点说明」写「图上为 Ag/AgCl，未换算」。")
L.append("- **不确定度**：肉眼读数一般 ±0.02 V；曲线陡直、刻度和数据点清晰时写 ±0.01；曲线平缓、判据模糊时写 ±0.05。")
L.append("- **读不出来就留空**：图上没有该试样、曲线没出现击穿、被水印遮挡，一律留空并在「取点说明」写原因。留空比猜值有用——`00.xlsx` 建模前会剔除空值，猜错的值会污染模型。\n")
L.append("### 2.4 五个必须由你顺手确认的字段\n")
L.append("1. **参比**：本批有 5 行 `参比` 列是 `未提及`，读图时看图上写的参比，一并填进「取点说明」，我据此回填 T 列。")
L.append("2. **温度**：29 行 `温度` 是 `未提及`，图上或图注若写了温度，一并记下。")
L.append("3. **判据**：图上给的是 Ec 还是 Eb，两者定义不同，别混。")
L.append("4. **是否是同一试样**：40 行在 `0.xlsx` 里存在同名不同值的记录，读图时以本次 PDF 为准，不要参考旧值。")
L.append("5. **图上无点蚀**：若曲线全程钝化、没有击穿，写 `无点蚀` 而不是留空，这与「没读到」是两回事。\n")
L.append("### 2.5 交付方式（三种任选）\n")
L.append("- **推荐**：填 `读图待办_20261109.csv` 的三列后发回，我合并（不动原表结构，最安全）。")
L.append("- 也可以直接在 `00.xlsx` 的 W 列按「00.xlsx 行号」定位改值，X 列写 ±，Z 列补说明。")
L.append("- 分批发也行：先发一部分，我先合一部分，不要求一次做完。\n")
L.append("## 三、结论：还剩多少、卡在哪\n")
L.append(f"- **总量**：{len(rows)} 行 / {n_doc} 篇。最多的是文献 "
         + "、".join(f"{k}（{v} 行）" for k, v in Counter({int(k): len(v) for k, v in by.items()}).most_common(5))
         + "。")
L.append(f"- **图上位置**：图 2 最多（{fig.get('图2', 0)} 行），其次图 6（{fig.get('图6', 0)}）、图 1（{fig.get('图1', 0)}）、图 3（{fig.get('图3', 0)}）。")
L.append("- **已有 50 行不用管**：已从 `0.xlsx` 回填，清单在 `读图回填来源_20261109.csv`，建议抽 5 行核对（尤其文献 3 的 S12/S13，`0.xlsx` 里顺序与 `00.xlsx` 相反，是按样本名对的，不是按行号）。")
L.append("- **注意 `0.xlsx` 的量级问题**：`0.xlsx` 的 Epit 列有 50 个格子记的是 mV 数字（最大 603、最小 −169），不是 V。回填时已按 `|值|>3` 判为 mV 折算，逐行记在审计表里。后续若有人直接拿 `0.xlsx` 建模，会得到差 1000 倍的结果。\n")
L.append("### 3.1 按文献的待办明细\n")
L.append("| 文献 | 待读行数 | 涉及图号 | 特殊提示 |")
L.append("|---|---|---|---|")
for no in sorted(by):
    rs = by[no]
    figs = "、".join(sorted({r["图号"].replace("N/A（待读图，", "").rstrip("）") for r in rs}))
    tips = []
    if any(r["参比"] == "未提及" for r in rs):
        tips.append("参比需确认")
    if any(r["温度"] == "未提及" for r in rs):
        tips.append("温度未提及")
    if any("样本名不一致" in r["状态"] for r in rs):
        tips.append("与 0.xlsx 样本名不同")
    if len(rs) >= 8:
        tips.append("行数多，建议单独排时间")
    L.append(f"| {no} | {len(rs)} | {figs} | {'；'.join(tips) if tips else '—'} |")
L.append("")
L.append("### 3.2 按工作量分三批（推荐顺序）\n")
small = [n for n in sorted(by) if len(by[n]) <= 3]
mid = [n for n in sorted(by) if 3 < len(by[n]) <= 7]
big = [n for n in sorted(by) if len(by[n]) > 7]
L.append(f"- **第一批（热身）**：{len(small)} 篇、{sum(len(by[n]) for n in small)} 行，每篇 1–3 行，用来熟悉取点标准。文献：{'、'.join(map(str, small))}。")
L.append(f"- **第二批（主体）**：{len(mid)} 篇、{sum(len(by[n]) for n in mid)} 行，每篇 4–7 行。文献：{'、'.join(map(str, mid))}。")
L.append(f"- **第三批（重活）**：{len(big)} 篇、{sum(len(by[n]) for n in big)} 行，每篇 8 行以上，多为温度/变形量系列，同一张图上多条曲线。文献：{'、'.join(map(str, big))}。\n")
L.append("## 四、建议\n")
L.append("1. **先做第一批 20 行左右**，我拿你的读数做一次一致性检查（量级、参比、与同文献其他行是否自相矛盾），确认标准对齐后再往下推，避免 177 行返工。")
L.append("2. **每篇文献只读一次**：同一篇的 8–14 行通常在同一张图上，一次性读完比来回翻页快。")
L.append("3. **时间预估**：单行 1–2 分钟（含翻页定位），177 行约 4–6 小时；按三批拆开，每批 1–2 小时。")
L.append("4. **不要动 W 列以外的内容**：成分、工艺、参比、温度这些是逐篇从原文核过的，改错会连带建模特征出错。")
L.append("5. **读数回填后我做三件事**：格式与量级校验 → 参比一致性校验（同文献内是否出现两种参比混用）→ 更新校验报告与 0/00 差异对照。")
L.append("6. **如果某一篇你判断根本读不出**：告诉我文献编号，我按「图上不可判」整篇标注，并在报告里单列，不硬凑数。\n")
L.append("## 五、文件位置\n")
L.append("- 待填清单：`备份\\00_录入_20261109\\读图待办_20261109.csv`（177 行，空着的是第 8–10 列）")
L.append("- 已回填清单：`备份\\00_录入_20261109\\读图回填来源_20261109.csv`（50 行，含 0.xlsx 原行号）")
L.append("- 数据件：`00.xlsx`（Sheet1，W 列 Epit，第 3 行起为数据）")
L.append("- 原始 PDF：`文献\\01.pdf`–`177.pdf`（只读，不要改名）")

open(os.path.join(ROOT, "读图作业说明_20261109.md"), "w", encoding="utf-8").write("\n".join(L))
print("生成 读图作业说明_20261109.md")
print("待办", len(rows), "篇数", n_doc, "三批", len(small), len(mid), len(big))
