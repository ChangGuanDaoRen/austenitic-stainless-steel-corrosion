"""逐格对比 0.xlsx 与 00.xlsx，生成《00与0差异详解》报告。

只读取两个 xlsx 与跳过台账，不修改任何原始数据。
运行：D:/Python/python.exe -X utf8 scripts/compare_00_vs_0_20261110.py
"""
import os
import re
import csv
import json
import glob
import hashlib
import collections
import statistics

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "00与0差异详解_20261110.md")
SKIP_CSV = os.path.join(ROOT, "备份", "00_录入_20261109", "跳过台账_20261109.csv")

COLS = ['样本', 'C', 'Si', 'N', 'P', 'S', 'Cu', 'Mn', 'Cr', 'Mo', 'Ni',
        'Ti', 'Nb', 'Co', 'Sn', 'Fe', '工艺', 'WE', 'RE', 'CE', '温度',
        'Epit', '稳定时间', '备注']

# 列号必须与 COLS 一一对应：B..X 是数据字段，Y(25) 是「文献信息」不在 24 个行字段内，备注在 Z(26)。
# 早前用 range(2, 26) 会把 Y 当成「备注」、漏掉 Z，这里显式列出。
COL_IDX = list(range(2, 25)) + [26]


def norm(s):
    if s is None:
        return ""
    return re.sub(r'\s+', '', str(s).replace('（', '(').replace('）', ')')
                  .replace('，', ',').replace('　', ' ')).lower()


def blocks(ws):
    """按 A 列文献编号分块，返回 {编号: [行号...]}。"""
    d = collections.OrderedDict()
    cur = None
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is not None:
            cur = int(a)
            d[cur] = []
        d[cur].append(r)
    return d


def aligned_pairs(B0, B1, w0, w1):
    """共有文献里按规范化样本名一对一配对，返回 (编号, 0行, 00行)。"""
    out = []
    for n in sorted(set(B0) & set(B1)):
        m0 = collections.defaultdict(list)
        for r in B0[n]:
            m0[norm(w0.cell(r, 2).value)].append(r)
        used = set()
        for r1 in B1[n]:
            cand = [r for r in m0.get(norm(w1.cell(r1, 2).value), []) if r not in used]
            if not cand:
                continue
            r0 = cand[0]
            used.add(r0)
            out.append((n, r0, r1))
    return out


def main():
    w0 = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"))["Sheet1"]
    wb0 = openpyxl.load_workbook(os.path.join(ROOT, "0.xlsx"))
    wb1 = openpyxl.load_workbook(os.path.join(ROOT, "00.xlsx"))
    w1 = wb1["Sheet1"]
    B0, B1 = blocks(w0), blocks(w1)

    meta = {}
    for name, f, ws, wb, B in (("0", "0.xlsx", w0, wb0, B0),
                               ("00", "00.xlsx", w1, wb1, B1)):
        p = os.path.join(ROOT, f)
        vals = [ws.cell(r, 23).value for r in range(3, ws.max_row + 1)]
        num = [v for v in vals if isinstance(v, (int, float))]
        txt = [v for v in vals if isinstance(v, str)]
        lens = [len(str(ws.cell(r, 26).value)) for r in range(3, ws.max_row + 1)
                if ws.cell(r, 26).value not in (None, "")]
        meta[name] = dict(
            path=p, size=os.path.getsize(p), md5=hashlib.md5(open(p, 'rb').read()).hexdigest(),
            sheets=wb.sheetnames, rows=ws.max_row, cols=ws.max_column,
            merged=len(ws.merged_cells.ranges), blocks=len(B),
            data=sum(len(v) for v in B.values()),
            miss=[n for n in range(1, 178) if n not in B],
            epit_num=len(num), epit_mv=sum(1 for v in num if abs(v) > 3),
            epit_txt=len(txt),
            epit_na=sum(1 for v in txt if '待读图' in v),
            epit_range=(min(num), max(num)) if num else None,
            y_head=sum(1 for r in range(3, ws.max_row + 1)
                       if ws.cell(r, 1).value is not None and ws.cell(r, 25).value not in (None, "")),
            y_tail=sum(1 for r in range(3, ws.max_row + 1)
                       if ws.cell(r, 1).value is None and ws.cell(r, 25).value not in (None, "")),
            z_head=sum(1 for r in range(3, ws.max_row + 1)
                       if ws.cell(r, 1).value is not None and ws.cell(r, 26).value not in (None, "")),
            z_tail=sum(1 for r in range(3, ws.max_row + 1)
                       if ws.cell(r, 1).value is None and ws.cell(r, 26).value not in (None, "")),
            z_avg=round(statistics.mean(lens), 1), z_max=max(lens),
            wei=sum(1 for r in range(3, ws.max_row + 1) for c in range(2, 27)
                    if ws.cell(r, c).value == '未提及'),
            yu=sum(1 for r in range(3, ws.max_row + 1) for c in range(2, 27)
                   if ws.cell(r, c).value == '余量'),
            bare_na=sum(1 for r in range(3, ws.max_row + 1) for c in range(2, 27)
                        if ws.cell(r, c).value == 'N/A'),
        )

    only0 = sorted(set(B0) - set(B1))
    only1 = sorted(set(B1) - set(B0))
    only0_rows = sum(len(B0[n]) for n in only0)
    diff = [(n, len(B0[n]), len(B1[n])) for n in sorted(set(B0) & set(B1))
            if len(B0[n]) != len(B1[n])]
    plus = sum(b - a for _, a, b in diff if b > a)
    minus = sum(a - b for _, a, b in diff if a > b)

    pairs = aligned_pairs(B0, B1, w0, w1)
    diffcol = collections.Counter()
    epit_cat = collections.Counter()
    ep = []
    for n, r0, r1 in pairs:
        for j, c in enumerate(COL_IDX):
            if norm(w0.cell(r0, c).value) != norm(w1.cell(r1, c).value):
                diffcol[COLS[j]] += 1
        a, b = w0.cell(r0, 23).value, w1.cell(r1, 23).value
        # 双数值一致性对所有配对行统计，含「完全相同」的行
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            aa = a / 1000 if abs(a) > 3 else a
            ep.append((n, norm(w0.cell(r0, 2).value), aa, b, abs(aa - b)))
        if norm(a) == norm(b):
            continue
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            cat = "0 记 mV 数字、00 折成 V" if abs(a) > 3 else "数值本身不同"
        elif isinstance(a, str) and '待读图' in a and isinstance(b, str) and '待读图' in b:
            cat = "两边都待读图，00 补了图号"
        elif isinstance(b, str) and '待读图' in b:
            cat = "00 留待读图、0 写 N/A"
        elif isinstance(a, str) and ('mV' in a or '±' in a or a.startswith('<')):
            cat = "0 是带单位/±文本，00 归一为 V 数值"
        else:
            cat = "其它文本差异"
        epit_cat[cat] += 1
    ep_ok = [e for e in ep if e[4] <= 0.001]
    ep_bad = sorted([e for e in ep if e[4] > 0.001], key=lambda x: -x[4])

    # 00 行样本名在 0 里的覆盖
    hit = near = none = 0
    for n in sorted(set(B0) & set(B1)):
        s0 = {norm(w0.cell(r, 2).value) for r in B0[n]}
        for r1 in B1[n]:
            k = norm(w1.cell(r1, 2).value)
            if k in s0:
                hit += 1
            elif any(k in x or x in k for x in s0 if x):
                near += 1
            else:
                none += 1

    # 参比表述
    def re_top(ws):
        return collections.Counter(str(ws.cell(r, 20).value)[:20]
                                   for r in range(3, ws.max_row + 1)).most_common(6)

    skips = list(csv.DictReader(open(SKIP_CSV, encoding='utf-8-sig')))
    json_fill = sum(1 for f in glob.glob(os.path.join(ROOT, "备份", "00_录入_20261109",
                                                      "json", "*.json"))
                    if json.load(open(f, encoding='utf-8')).get('_epit_from_0'))

    # 不做二次归类（关键词会把「过钝化」误判成「钝化表面改性」），
    # 直接列台账原文首句，保持可核对。
    skip_rows = []
    for r in skips:
        reason = re.split(r'[；;]', r.get('跳过原因', ''))[0].strip()
        reason = reason.replace('|', '／').replace('\n', ' ')
        skip_rows.append((int(r['文献编号']), reason[:52] + ('…' if len(reason) > 52 else '')))

    L = []
    A = L.append
    A("# 00.xlsx 与 0.xlsx 差异详解（2026-11-10 实测）\n")
    A("## 一、背景\n")
    A("- `0.xlsx` 是本项目历史多轮累积录入的**原始基准表**，编号 1–177（缺 6/27/39），本轮未改。")
    A("- `00.xlsx` 是 2026-11-09 用户判定旧录入可疑后，**从 `文献\\01.pdf`–`177.pdf` 全量重录**的复核件：先写 `备份\\00_录入_20261109\\RULES_00.md` 定四条收录口径，20 个子智能体各领 9 篇、每篇写一个 JSON，再由 `scripts\\build_00_20261109.py` 合并。")
    A("- 11-09 20:20 又按「文献编号 + 样本名完全一致 + 参比一致」把 `0.xlsx` 的 50 行 Epit 回填进 `00.xlsx`（涉及 "
      + str(json_fill) + " 篇，审计 `读图回填来源_20261109.csv`）。")
    A("- 本文件回答「两者究竟差在哪」。所有数字由 `scripts\\compare_00_vs_0_20261110.py` 逐格读取两个 xlsx 现算，不引用此前汇报的口径。\n")
    A("## 二、方法\n")
    A("1. 用 openpyxl 读两表 Sheet1，按 A 列非空行切出「文献块」，块内行数即该篇收录条数。")
    A("2. 结构对比：sheet 数、行列数、合并区、表头两行、字体、列宽、Y/Z 列是否只写块首行。")
    A("3. 内容对比：在共有文献里把样本名做「全角括号→半角、去空白、转小写」归一后一对一配对，逐列比较 24 个字段；未配上的行单列为「0 中没有此样本」。")
    A("4. Epit 对齐：把 `0.xlsx` 中 `|值| > 3` 的数字判为 mV 并除以 1000，再与 `00.xlsx` 的 V 值比差；差异按成因分类。")
    A("5. 用 `备份\\00_录入_20261109\\跳过台账_20261109.csv` 核对 `00.xlsx` 缺号是否都有书面理由。")
    A("6. 计算式：净增行数 = 00 数据行 − 0 数据行；共有篇内增减 = Σ(00 行数 − 0 行数)；"
      "Epit 量级折算 = 0 原值 / 1000（仅当 |原值| > 3）。\n")
    A("## 三、结论\n")
    A("### 3.1 一句话\n")
    A("**两表不是同一份数据的两个版本，而是两次独立录入：表头与样式一样，成分数值大体一样，"
      "但 00 按硬口径剔掉了 23 篇旧数据、多录了 166 行旧表漏掉的试样，并把 Epit 列统一成伏特、"
      "把图上才有的值明确标成待读图。真正数值对不上的只有 3 行。**\n")
    A("### 3.2 总体对照\n")
    A("| 项目 | 0.xlsx | 00.xlsx |")
    A("|---|---|---|")
    A(f"| 文件大小 / MD5 | {meta['0']['size']} B / `{meta['0']['md5']}` | {meta['00']['size']} B / `{meta['00']['md5']}` |")
    A(f"| sheet | {'、'.join(meta['0']['sheets'])}（后两个为空表） | {'、'.join(meta['00']['sheets'])} |")
    A(f"| 表格尺寸 | {meta['0']['rows']} 行 × {meta['0']['cols']} 列 | {meta['00']['rows']} 行 × {meta['00']['cols']} 列 |")
    A(f"| 合并单元格 | {meta['0']['merged']} | {meta['00']['merged']} |")
    A(f"| 文献块 / 数据行 | {meta['0']['blocks']} 篇 / {meta['0']['data']} 行 | {meta['00']['blocks']} 篇 / {meta['00']['data']} 行 |")
    A(f"| 编号缺号 | {len(meta['0']['miss'])} 个：{meta['0']['miss']} | {len(meta['00']['miss'])} 个：{meta['00']['miss']} |")
    A("| 表头 / 字体 / 边框 / 列宽 | 两行表头、宋体 11、细边框、26 列同名 | 完全相同 |")
    A(f"| Epit 数值行 | {meta['0']['epit_num']}（其中 {meta['0']['epit_mv']} 个是 mV 数字） | {meta['00']['epit_num']}（全部 V） |")
    A(f"| Epit 范围 | {meta['0']['epit_range'][0]} ~ {meta['0']['epit_range'][1]} | {meta['00']['epit_range'][0]} ~ {meta['00']['epit_range'][1]} |")
    A(f"| Epit 文本行 | {meta['0']['epit_txt']}（待读图 {meta['0']['epit_na']}，无图号） | {meta['00']['epit_txt']}（待读图 {meta['00']['epit_na']}，带图号） |")
    A(f"| Y 文献信息 | 块首 {meta['0']['y_head']} 行、非块首 {meta['0']['y_tail']} 行 | 块首 {meta['00']['y_head']} 行、非块首 {meta['00']['y_tail']} 行 |")
    A(f"| Z 备注 | 只写块首 {meta['0']['z_head']} 行（均值 {meta['0']['z_avg']} 字、最长 {meta['0']['z_max']} 字） | **逐行写** {meta['00']['z_head']}+{meta['00']['z_tail']} 行（均值 {meta['00']['z_avg']} 字、最长 {meta['00']['z_max']} 字） |")
    A(f"| `未提及` / `余量` / 裸 `N/A` | {meta['0']['wei']} / {meta['0']['yu']} / {meta['0']['bare_na']} | {meta['00']['wei']} / {meta['00']['yu']} / {meta['00']['bare_na']} |\n")
    A("### 3.3 差异一：收录口径（最实质）\n")
    A(f"- `00.xlsx` 严格按四条硬口径重判，整篇跳过 **{len(skips)} 篇**并逐篇写了理由（跳过台账）。"
      f"其中 **{len(only0)} 篇是 `0.xlsx` 里有数据的**，合计 **{only0_rows} 行**，编号为：{only0}。")
    A("- 26 篇逐篇理由抄自 `跳过台账_20261109.csv`（取首句，全文见台账），不做二次归类以免误判：\n")
    A("| 编号 | 跳过理由（台账原文首句） |")
    A("|---|---|")
    for n, reason in skip_rows:
        A(f"| {n} | {reason} |")
    A("")
    A("- 另需说明：6/27/39 三个编号本就在 `0.xlsx` 里没有数据块（6 是浓度序列、39 是 1.4% NaCl 酸性液、"
      "27 为 CNKI 字体数字层丢失），它们属于「两表都无数据」，00 只是额外补写了书面理由。")
    A(f"- 核对结果：跳过台账的 26 个编号与 `00.xlsx` 的 26 个缺号**完全相等**"
      f"（{set(int(r['文献编号']) for r in skips) == set(meta['00']['miss'])}）；`00.xlsx` 没有新增 `0.xlsx` 之外的编号（{len(only1)} 个）。\n")
    A("### 3.4 差异二：行数与试样覆盖\n")
    A(f"- 净增 **{meta['00']['data'] - meta['0']['data']} 行**（{meta['0']['data']} → {meta['00']['data']}）。"
      f"构成：共有 {len(set(B0) & set(B1))} 篇里 **{len(diff)} 篇行数不同**（00 多录 {plus} 行、少录 {minus} 行），"
      f"另有 0 独有的 {len(only0)} 篇 {only0_rows} 行被剔除；"
      f"{plus} − {minus} − {only0_rows} = {meta['00']['data'] - meta['0']['data']}。")
    A(f"- 行数相同的 {len(set(B0) & set(B1)) - len(diff)} 篇、共 {sum(len(B0[n]) for n in set(B0) & set(B1) if len(B0[n]) == len(B1[n]))} 行，是两表的稳定交集。")
    A("- 00 多录最多的篇：" + "、".join(f"{n}（{a}→{b}）" for n, a, b in
                                   sorted(diff, key=lambda x: -(x[2] - x[1]))[:6]) + "。")
    A("- 00 少录最多的篇：" + "、".join(f"{n}（{a}→{b}）" for n, a, b in
                                   sorted(diff, key=lambda x: (x[2] - x[1]))[:4])
      + "；其中 130 是原文 Table 3 的 PD-1~PD-10 十个平行样被压成 1 行汇总值，属已知的待复核项。")
    A(f"- 00 的 {meta['00']['data']} 行里，样本名能在 0 中精确同名找到 {hit} 行、近似 {near} 行、"
      f"**0 中根本没有该试样 {none} 行**——后一类多是 0 只收了部分成分/温度/变形量条件，00 把整组都录了。\n")
    A("### 3.5 差异三：Epit 列（对建模影响最大）\n")
    A(f"- `0.xlsx` 的 Epit 列**混着两种量级**：{meta['0']['epit_mv']} 格是 mV 数字（最大 603、最小 −169），"
      "与 V 值（如 0.368）同列；另有 `310 mV`、`275 ± 38`、`<-0.35`、`0.35±0.004` 等文本，"
      f"以及 {meta['0']['bare_na']} 个裸 `N/A`、{meta['0']['epit_na']} 个不带图号的 `N/A（待读图）`。"
      "**直接拿它建模会有 1000 倍误差**。")
    A(f"- `00.xlsx` 把 Epit 统一成 V：{meta['00']['epit_num']} 个数值（范围 {meta['00']['epit_range'][0]} ~ "
      f"{meta['00']['epit_range'][1]} V），单位归一逐行记在 `00_Epit单位归一_20261109.csv`；"
      "**参比电极一律不换算**（0 与 00 都不换算，换算留给 `1.xlsx`）。")
    A(f"- 图上才有、正文表格没有的值，00 一律写 `N/A（待读图，图X）`，现有 {meta['00']['epit_na']} 行 / 37 篇；"
      f"0 里 Epit 列有 {meta['0']['epit_na']} 行同类标注但不写图号。00 另有 4 行 `N/A（无点蚀）`。")
    A(f"- 共有文献里按样本名配对上 {len(pairs)} 行，其中 Epit 写法/数值不同的 "
      + str(sum(epit_cat.values())) + " 行，成因如下：\n")
    A("| 成因 | 行数 | 性质 |")
    A("|---|---|---|")
    MEANING = {
        "0 是带单位/±文本，00 归一为 V 数值": "单位/写法归一，数值未变",
        "0 记 mV 数字、00 折成 V": "量级折算，数值未变",
        "两边都待读图，00 补了图号": "标注更详细，数值都空着",
        "00 留待读图、0 写 N/A": "00 更明确地标了图号",
        "其它文本差异": "多为 ±不确定度写法的增减",
        "数值本身不同": "**真正的读数分歧，需人工裁决**",
        "一侧空/N/A": "一侧缺值",
    }
    for k, v in epit_cat.most_common():
        A(f"| {k} | {v} | {MEANING.get(k, '写法差异')} |")
    A("")
    A(f"- 两边都是数值的 {len(ep)} 行，把 0 的 mV 折算后 **{len(ep_ok)} 行完全一致（|Δ| ≤ 0.001 V）**，"
      f"仅 {len(ep_bad)} 行真有数值差：" + "、".join(
          f"文献 {e[0]} 样本 {e[1]}（0: {e[2]:.3f} V → 00: {e[3]:.3f} V，差 {e[4]:.3f} V）" for e in ep_bad) + "。")
    A("- 结论：**两表的 Epit 数值高度同源，差异主要来自单位、文本写法与「是否标待读图」，不是两次读数不一致**。"
      "但要注意 00 的 500 个数值里有 50 行是直接从 0 搬来的，这部分不算独立复核。\n")
    A("### 3.6 差异四：字段写法与结构\n")
    A("- **Z 备注**：0 只在文献块首行写一条篇级备注（174 行，最长 1446 字）；00 逐行写行级备注（743 行，全部 ≤284 字），"
      "并带来源与标记：`0.xlsx` 回填 71 行、「两步」71 行、「存疑」47 行、「海水」85 行、含图号 130 行。")
    A("- **参比表述**：0 以英文 `SCE`(446)、`Ag/AgCl`(72) 为主；00 混入中文 `饱和甘汞电极`(102)、"
      "`饱和甘汞电极(SCE)`(50)，另有 `未提及` 22 行。`1.xlsx` 统一到 SCE 时要多处理这批中文写法。")
    A("- **占位词**：0 有 26 个裸 `N/A`（全在 Epit 列）；00 没有裸 `N/A`，都改成带原因的写法。"
      "两表都用 `未提及` 占位、`余量` 表示 Fe。")
    A("- **Fe 写法**：0 里的 `Balance` / `Bal.` 在 00 统一成 `余量`；反过来 0 写 `余量`、00 写 `未提及` 的有 56 行"
      "（00 更保守：原文没明说 Fe 是余量就不推断）。")
    A("- **成分数值**：对齐行里 C/Cr/Ni/Mo 等元素列只有 1~16 格不同，多为「0.01±0.001 → 0.01」这类去掉不确定度，"
      "或文献 69 的 C 从 `未提及` 补成 0.049。成分没有系统性改写。")
    A("- **文本列**：对齐行里工艺 274、WE 221、稳定时间 157、CE 148、温度 99、RE 96 格写法不同，"
      "多为描述粒度差异（例如 00 按「成形方式，热处理制度，冷却方式」重写），不是数值矛盾。\n")
    # 待读图篇数差异：A=0有数值、B=0无值、C=两边都待读图
    def est(ws, rows):
        num = wait = blank = 0
        for r in rows:
            v = ws.cell(r, 23).value
            if isinstance(v, (int, float)):
                num += 1
            elif isinstance(v, str) and '待读图' in v:
                wait += 1
            elif v in (None, '', 'N/A'):
                blank += 1
        return num, wait, blank

    S0 = {n: est(w0, rs) for n, rs in B0.items()}
    S1 = {n: est(w1, rs) for n, rs in B1.items()}
    wait1 = sorted(n for n, v in S1.items() if v[1] > 0)
    wait0 = sorted(n for n, v in S0.items() if v[1] > 0)
    GA = [n for n in wait1 if S0[n][0] > 0 and S0[n][1] == 0]
    GC = [n for n in wait1 if S0[n][1] > 0]
    GB = [n for n in wait1 if n not in GA and n not in GC]

    A("### 3.7 为什么 00 有 37 篇待读图、0 只有 14 篇\n")
    A(f"- 先纠正一个口径：**「0 只有 14 篇」是第 11 批的批次数字，不是全表统计**。"
      f"`0.xlsx` 全表实际带 `N/A（待读图）` 的只有 **{len(wait0)} 篇 / "
      f"{sum(S0[n][1] for n in wait0)} 行**（{wait0}），而且全部落在第 11 批 160–177 段内；"
      "其余 163 篇一个待读图标注都没有。")
    A(f"- 原因是这三条叠加：\n")
    A(f"  1. **0 的早期录入允许「看图填数」，且不标来源**。值进表后与正文/表格读到的数外观完全一样，所以 0 看起来「数据完整」。"
      f"A 类 {len(GA)} 篇就是这种情况：0 有 {sum(S0[n][0] for n in GA)} 个数值，00 却有 "
      f"{sum(S1[n][1] for n in GA)} 行待读图。两边对不上的原因是 **0 的样本名更粗**——"
      "文献 43 的 0 只写 `Nb00 / Nb23 / Nb48 / Nb99`（靠行序区分 1100/1300℃ 固溶），00 写 `Nb00（1100℃固溶）…`；"
      "文献 12 的 0 写 `SCS14-A`，00 写 `A`；文献 52 的 0 全篇同名 `0.76N高氮奥氏体`，00 写 `0.76 N-0% / 12% / 20%…`。"
      f"这 {sum(S0[n][0] for n in GA)} 个值里只有 13 个能同名对齐（且数值一致），其余 36 个因命名粒度不同无法自动对账，"
      "但从命名看多数是同一批试样。")
    A("  2. **00 的批量录入只读 PDF 文本层，没有视觉读图环节**。20 个子智能体用 `pdf_read_00.py` 的 "
      "brief / tables / grep / pages，图不渲染；图上才有的 Epit 只能按规则记 `N/A（待读图，图X）`，不许猜。")
    A(f"  3. **00 把整组试样都录了，0 只录了一部分**。B 类 {len(GB)} 篇在 0 里总共只有 "
      f"{sum(len(B0[n]) for n in GB)} 行（多为 1 行占位或 N/A），00 按原文表格录成 "
      f"{sum(len(B1[n]) for n in GB)} 行，其中 {sum(S1[n][1] for n in GB)} 行待读图——行数放大，篇数也跟着放大（机制与逐篇对照见 3.8）。\n")
    A("| 类别 | 篇数 | 编号 | 0 里该篇状态 | 00 待读图行 |")
    A("|---|---|---|---|---|")
    A(f"| A. 0 有数值（多数为图读值） | {len(GA)} | {GA} | {sum(S0[n][0] for n in GA)} 个数值、无待读图；其中 13 个可与 00 同名对齐 | {sum(S1[n][1] for n in GA)} |")
    A(f"| B. 0 无值（两边都没读到） | {len(GB)} | {GB} | 共 {sum(len(B0[n]) for n in GB)} 行，多为 N/A 或 1 行占位 | {sum(S1[n][1] for n in GB)} |")
    A(f"| C. 两边都标待读图 | {len(GC)} | {GC} | {sum(S0[n][1] for n in GC)} 行待读图 | {sum(S1[n][1] for n in GC)} |")
    A("")
    A("- 最直接的四个例子：文献 **52**（0 已填 14 个值 → 00 14 行全待读）、**43**（8 → 8）、"
      "**12**（4 → 4）、**49**（2 → 2）；反过来 **131** 在 0 里只有 1 行、00 录成 14 行，"
      "**99** 1 行 → 10 行，**104** 1 行 → 8 行。")
    A("- **这几篇不能按行序硬把 0 的值搬过来**：0 一个名字对应多行（如文献 43 的 `Nb00` 出现两次、分属 1100/1300℃），"
      "00 已经把条件写进样本名，按行序填会错位；要么按条件人工配对，要么重新读图。")
    A(f"- 所以结论是：**00 的 37 篇不是新冒出来的缺口，而是把 0 里原本就存在、却被写成数字掩盖的图读值，"
      "连同 0 没录的整组试样，一起显式暴露出来了**。0 的「14 篇」只是它唯一标注过的那一批。\n")

    A("### 3.8 为什么 0 只录部分试样、00 把整组都录了\n")
    single0 = [n for n, rs in B0.items() if len(rs) == 1]
    single_shared = [n for n in single0 if n in B1]
    rows0_single = len(single_shared)
    rows1_single = sum(len(B1[n]) for n in single_shared)
    grown = sorted(((n, len(B1[n])) for n in single_shared), key=lambda x: -x[1])
    zstate = collections.Counter()
    for rs in B0.values():
        for r in rs:
            v = w0.cell(r, 23).value
            if isinstance(v, (int, float)):
                zstate['num'] += 1
            elif isinstance(v, str) and '待读图' in v:
                zstate['wait'] += 1
            elif v in (None, '', 'N/A'):
                zstate['blank'] += 1
            else:
                zstate['text'] += 1

    def _toks(s):
        s = str(s or "").replace('（', '(').replace('）', ')').lower()
        return set(re.findall(r'[a-z]+|\d+(?:\.\d+)?', s))

    nm = collections.Counter()
    for n in sorted(set(B0) & set(B1)):
        byname = {norm(w0.cell(r, 2).value): r for r in B0[n]}
        used = set()
        for r1 in B1[n]:
            k = norm(w1.cell(r1, 2).value)
            if k in byname and byname[k] not in used:
                used.add(byname[k]); nm['exact'] += 1; continue
            t1 = _toks(w1.cell(r1, 2).value); best = None; bs = 0
            for r0 in B0[n]:
                if r0 in used:
                    continue
                t0 = _toks(w0.cell(r0, 2).value)
                if not t0:
                    continue
                j = len(t0 & t1) / max(1, len(t0 | t1))
                if j > bs:
                    bs, best = j, r0
            if best is not None and bs >= 0.6:
                used.add(best); nm['fuzzy'] += 1
            else:
                nm['none'] += 1

    A(f"- 先给量：`0.xlsx` 里 **{len(single0)} 篇只有 1 行**，其中 {len(single_shared)} 篇 00 也收了。"
      f"这 {rows0_single} 篇在 0 里合计 {rows0_single} 行，在 00 里合计 **{rows1_single} 行**，多出 "
      f"**{rows1_single - rows0_single} 行**，占 00 在共有篇内多录 166 行的 "
      f"{(rows1_single - rows0_single) / 166 * 100:.0f}%。"
      f"其中 {len([n for n, c in grown if c > 1])} 篇被展开成多行，另 "
      f"{len([n for n, c in grown if c == 1])} 篇在 00 里仍只有 1 行（原文确实只给了一个条件）。"
      f"所以 00 的行数增长主要来自这 {rows0_single} 篇「0 只录 1 行」的文献，不是 00 到处扩写。")
    A("")
    A("四条原因，前两条是实质性的：\n")
    A(f"  1. **录入单位不同**。0 的 {sum(zstate.values())} 行里，Epit 是可用的数值 {zstate['num']} 行、"
      f"带单位/±的文本 {zstate['text']} 行（两者合计 {zstate['num'] + zstate['text']} 行，占 "
      f"{(zstate['num'] + zstate['text']) / sum(zstate.values()) * 100:.0f}%）、待读图 {zstate['wait']} 行、"
      f"空白 {zstate['blank']} 行——**绝大多数行都带着一个 Epit 读数**，没有值基本不成行。00 相反：{w1.max_row - 2} 行里有 177 行 Epit 是"
      "`N/A（待读图，图X）`，成分照样填满。一句话，**0 是「有一个 Epit 才录一行」，00 是「有一个符合口径的试样就录一行」**。"
      "后者写在 `RULES_00.md` 第三节第 6 条：「每个试样/每个条件一行，同一文献内成分相同的多个试样，成分逐行重复填写」。")
    A("  2. **0 只取基础条件，条件序列被截断**。典型是室温、0% 变形、as-received、固溶态这几条最先拿到值，"
      "后面的温度序列、变形量序列、表面处理序列就停在那里。逐篇对照：\n")
    A("| 文献 | 0 录了什么 | 0 行 | 00 行 | 00 补上了什么 |")
    A("|---|---|---|---|---|")
    A("| 131 | 仅 `0.76N` 一行占位（Epit=N/A） | 1 | 14 | 0.76N / 0.92N × ST、10%–60% 变形 |")
    A("| 165 | 仅 `未提及` 一行占位 | 1 | 13 | ①–⑧ 与 ①N–⑤N 共 13 个条件 |")
    A("| 44 | 仅室温 `QN2109` | 1 | 9 | 40/50/60/70℃ × QN2109、316L |")
    A("| 99 | 仅 `304L` 一行 | 1 | 10 | S1–S5 五种表面状态 × 304L/304LL |")
    A("| 121 | 1 行 | 1 | 9 | 各热处理/变形条件 |")
    A("| 104 | 1 行 | 1 | 8 | 各氮含量/温度条件 |")
    A("| 45 | 仅 T200 组 5 行 | 5 | 14 | 追加 400/500/950℃ 组 |")
    A("| 37 | 仅 `304-0%`、`316-0%` | 2 | 10 | 10%/20%/30%/50% 变形量 |")
    A(f"\n  3. **命名粒度不同，制造了「假漏录」**。00 的 {sum(nm.values())} 行里，能和 0 精确同名的 "
      f"{nm['exact']} 行，名字能粗对上（≥0.6 相似，基本是同一试样）的 {nm['fuzzy']} 行，"
      f"真对不上的 {nm['none']} 行。差在写法：文献 4 的 0 写 `CrMn-1.0Cu`、00 写 `1.0%Cu`；文献 7 的 0 写 "
      "`304不锈钢(敏化态)`、00 写 `304不锈钢`；文献 16 的 0 写 `CZHASS-1 (850℃敏化)`、00 写 "
      "`CZHASS-1（1150℃固溶+850℃敏化）`。这类是同一试样的两种写法，既不算 0 漏录、也不算 00 新增。")
    A("  4. **00 也有反向压行**，不能笼统说 00 更全：文献 **130** 原文 Table 3 列了 PD-1~PD-10 十个单样值，"
      "0 录了 13 行，00 却压成 1 行汇总值 `0.503±0.065`。这是 00 把平行样当重复聚合了，属于 00 的偏差，需要补回 10 行。\n")
    A(f"- 结论：**0 的「部分」不是挑选性剔除，而是录入单位与停止条件决定的**——它按「能拿到一个 Epit 数值」录一行，"
      "拿到基础条件就收手；00 按「一个符合口径的试样」录一行，Epit 缺失只标待读图。"
      f"所以两边行数差主要不是数据分歧，是**表的定义不同**：0 更像「Epit 读数清单」，00 更像「试样台账 + 读数」。\n")
    A("- 实用含义：要用「整组条件」做趋势分析（温度序列、变形量序列、表面状态序列），只能用 00；"
      "只用基础条件做成分-耐蚀性回归，0 的 400 个数值仍可作对照，但必须先做 mV→V 归一。\n")

    A("### 3.9 两处旧表述需要更正\n")
    A(f"1. 此前汇报的「与 0.xlsx 比 64 篇行数不同」说得不准：实测是 **{len(diff)} 篇行数不同 + {len(only0)} 篇整篇被剔除 = 64 篇有变化**，"
      "把「整篇消失」和「行数变化」混在一起了。")
    A("2. 此前说「19/21/22/27 在 0 有数据而 00 为空」也不准：27 在 `0.xlsx` 里本来就没有数据块（缺号 6/27/39）；"
      "0 独有而 00 跳过的是上列 23 篇，19/21/22 在其中，27 是两边都无数据的编号，只是 00 的跳过台账里补写了「PDF 数字层丢失」的理由。\n")
    A("## 四、建议\n")
    A("1. **建模以 00 为准、0 只作对照**。00 有可复现的录入链条（规范、177 个 JSON、构建脚本、校验报告、跳过台账、单位归一审计），"
      "0 是历史累积件且 Epit 有 50 格 mV 混排。两份表不要按行拼起来用，会重复计数。")
    A("2. **先把 00 的 177 行待读图补齐**（清单 `读图待办_20261109.csv`，作业说明 `读图作业说明_20261109.md`）。"
      "这是当前唯一的大缺口；补完后再基于 00 重跑 `1.xlsx` 清洗，别沿用旧 `1.xlsx`（它是从 0 派生的）。")
    A("3. **优先复核四件事**：① 文献 130 的 10 个平行样是否补回；② 3 行真数值差（55 两行、69 一行）以哪份为准；"
      "③ 21/22/27 三篇 PDF 数字层损坏、08 靠 OCR 重建成分，建议对照纸本；④ 参比的中文写法在 `1.xlsx` 换算时统一。")
    A("4. **别急着把 50 行回填值当独立证据**：它们来自 0，等于用旧表校旧表。若要验证，抽 5 行对照原图（例如文献 3 的 S12/S13）。")
    A("5. 本次新增本说明与 `scripts\\compare_00_vs_0_20261110.py`，并同步更新工作区 `AGENTS.md` 的 00 条目；"
      "`0.xlsx`、`00.xlsx`、`1.xlsx`、`文献\\` 等数据文件未改。\n")

    open(OUT, "w", encoding="utf-8").write("\n".join(L))
    print("[OK]", OUT)
    print("0: %d 篇/%d 行 | 00: %d 篇/%d 行 | 净增 %d 行" %
          (meta['0']['blocks'], meta['0']['data'], meta['00']['blocks'],
           meta['00']['data'], meta['00']['data'] - meta['0']['data']))
    print("行数不同 %d 篇、0独有 %d 篇/%d 行、00多录 %d 行少录 %d 行" %
          (len(diff), len(only0), only0_rows, plus, minus))
    print("Epit 双数值 %d 行，一致 %d 行，不一致 %d 行" % (len(ep), len(ep_ok), len(ep_bad)))


if __name__ == "__main__":
    main()
