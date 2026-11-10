# -*- coding: utf-8 -*-
"""按师兄 2026-10-26 口径统计：只考虑最后一步热处理的温度和时间。

数据源 1.xlsx（150 条）。解析逻辑沿用 scripts\group_last_heat.py。
用法：python scripts/group_last_heat_1.py [dry|write]
  dry   只打印
  write 在 1.xlsx 里新增/替换「热处理分组」sheet
"""
import collections
import re
import sys

import openpyxl

SRC = "1.xlsx"
COL_ID, COL_NAME, COL_PROC, COL_EPIT = 1, 2, 18, 23

STEP_RE = re.compile(
    r"(\d{3,4}(?:\s*[-~]\s*\d{3,4})?)\s*℃[^0-9]{0,14}?"
    r"(\d+(?:\.\d+)?(?:\s*[-~]\s*\d+(?:\.\d+)?)?)\s*(h|小时|min|分钟)"
)
NO_HEAT = "无热处理记录（或未写温度/时间）"


def norm_temp(s):
    s = re.sub(r"\s+", "", s)
    if "-" in s or "~" in s:
        a, b = re.split(r"[-~]", s)
        return f"{a}–{b}"
    return s


def norm_time(num, unit):
    num = re.sub(r"\s+", "", num)
    if "-" in num or "~" in num:
        a, b = re.split(r"[-~]", num)
        return f"{a}–{b}{'h' if unit in ('h', '小时') else 'min'}"
    v = float(num)
    if unit in ("min", "分钟"):
        # 720min / 1440min 这类整小时写法统一成 h，避免与 12h / 24h 分到两组
        if v and v % 60 == 0:
            return f"{v / 60:g}h"
        return f"{v:g}min"
    return f"{v:g}h"


def steps(text):
    return [(norm_temp(t), norm_time(tm, u)) for t, tm, u in STEP_RE.findall(text)]


def temp_mid(s):
    """温度键取中值（用于 ±25℃ 放宽合并）"""
    nums = [float(x) for x in re.split(r"[-–~]", s)]
    return sum(nums) / len(nums)


def to_minutes(t):
    m = re.match(r"^([\d.]+)(?:–[\d.]+)?(h|min)$", t)
    v = float(m.group(1))
    return v * 60 if m.group(2) == "h" else v


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb["Sheet1"]
    rows = [(ws.cell(row=r, column=COL_ID).value,
             ws.cell(row=r, column=COL_NAME).value,
             ws.cell(row=r, column=COL_PROC).value,
             ws.cell(row=r, column=COL_EPIT).value)
            for r in range(3, ws.max_row + 1)
            if ws.cell(row=r, column=COL_NAME).value]

    groups = collections.defaultdict(list)
    detail = []
    for pid, name, proc, epit in rows:
        st = steps(str(proc))
        if st:
            key = f"{st[-1][0]}℃×{st[-1][1]}"
        else:
            key = NO_HEAT
        groups[key].append((pid, name, epit))
        detail.append((pid, name, proc, key, epit))

    strict = {k: v for k, v in groups.items() if k != NO_HEAT}
    order = sorted(strict.items(), key=lambda x: (-len(x[1]), x[0]))

    print(f"总条数 {len(rows)}")
    print(f"\n【严格口径】最后一步热处理 温度×时间")
    print(f"{'分组':<24}{'条数':>4}  文献编号")
    print("-" * 60)
    for k, v in order:
        papers = sorted({str(x[0]) for x in v}, key=int)
        print(f"{k:<24}{len(v):>4}  {','.join(papers)}")
    print(f"{NO_HEAT:<24}{len(groups.get(NO_HEAT, [])):>4}")

    with_n = sum(len(v) for v in strict.values())
    print(f"\n有最后一步热处理记录：{with_n} 条 / {len(strict)} 组")
    print(f"最大组：{order[0][0]} = {len(order[0][1])} 条")
    for th in (5, 4, 3, 2):
        sel = [(k, len(v)) for k, v in order if len(v) >= th]
        print(f"≥{th} 条：{len(sel)} 组 / {sum(n for _, n in sel)} 条 -> {sel}")

    # 温度 ±25℃ 放宽合并（时间须一致）
    loose = collections.defaultdict(list)
    for (t, tm), items in [(tuple(re.split("℃×", k)), v) for k, v in strict.items()]:
        pass
    loose = collections.defaultdict(list)
    loose_items = []
    for k, v in strict.items():
        t, tm = k.split("℃×")
        loose_items.append((temp_mid(t), to_minutes(tm), tm, v))
    used = [False] * len(loose_items)
    merged = []
    for i in range(len(loose_items)):
        if used[i]:
            continue
        ti, mi, tm, vi = loose_items[i]
        cur = list(vi)
        used[i] = True
        for j in range(i + 1, len(loose_items)):
            if used[j]:
                continue
            tj, mj, tmj, vj = loose_items[j]
            if abs(ti - tj) <= 25 and mi == mj:
                cur += vj
                used[j] = True
        merged.append((f"≈{ti:g}℃×{tm}", cur))
    merged.sort(key=lambda x: (-len(x[1]), x[0]))
    print(f"\n【温度 ±25℃ 放宽、时间须一致】")
    for k, v in merged[:15]:
        papers = sorted({str(x[0]) for x in v}, key=int)
        print(f"{k:<24}{len(v):>4}  {','.join(papers)}")
    for th in (10, 5):
        sel = [(k, len(v)) for k, v in merged if len(v) >= th]
        print(f"≥{th} 条：{len(sel)} 组 / {sum(n for _, n in sel)} 条 -> {sel}")

    # 补充视角 1：只卡温度（±25℃），时长不限
    t_items = [(temp_mid(k.split("℃×")[0]), v) for k, v in strict.items()]
    t_items.sort()
    tclusters = []
    for ti, v in t_items:
        if tclusters and abs(ti - tclusters[-1][0]) <= 25:
            tclusters[-1][1] += v
        else:
            tclusters.append([ti, list(v)])
    tclusters.sort(key=lambda x: -len(x[1]))
    print(f"\n【补充：只卡温度 ±25℃，时长不限】")
    for ti, v in tclusters:
        papers = sorted({str(x[0]) for x in v}, key=int)
        print(f"≈{ti:g}℃{'':<10}{len(v):>4}  {','.join(papers)}")
    sel = [(f"≈{ti:g}℃", len(v)) for ti, v in tclusters if len(v) >= 10]
    print(f"≥10 条：{len(sel)} 组 / {sum(n for _, n in sel)} 条 -> {sel}")

    # 补充视角 2：无热处理记录的构成
    noh = groups.get(NO_HEAT, [])
    by_paper = collections.Counter(str(x[0]) for x in noh)
    print(f"\n【补充：无最后一步热处理记录 {len(noh)} 条的构成】")
    for k, v in sorted(by_paper.items(), key=lambda kv: -kv[1]):
        print(f"  文献{k}：{v} 条")

    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    if mode != "write":
        print("\n[dry] 未写盘")
        return

    if "热处理分组" in wb.sheetnames:
        del wb["热处理分组"]
    g = wb.create_sheet("热处理分组")
    g.append(["A. 严格口径：最后一步热处理 温度×时间"])
    g.append(["分组", "条数", "文献编号", "Epit 范围(V vs SCE)"])
    for k, v in order:
        eps = [x[2] for x in v if isinstance(x[2], (int, float))]
        rng = f"{min(eps):.3f}~{max(eps):.3f}" if eps else "-"
        g.append([k, len(v), ",".join(sorted({str(x[0]) for x in v}, key=int)), rng])
    g.append([NO_HEAT, len(groups.get(NO_HEAT, [])), "", ""])
    g.append([])
    g.append(["B. 温度 ±25℃ 放宽（时间须一致）"])
    g.append(["分组", "条数", "文献编号"])
    for k, v in merged:
        g.append([k, len(v), ",".join(sorted({str(x[0]) for x in v}, key=int))])
    g.append([])
    g.append(["C. 补充：只卡温度 ±25℃、时长不限"])
    g.append(["分组", "条数", "文献编号"])
    for ti, v in tclusters:
        g.append([f"≈{ti:g}℃", len(v), ",".join(sorted({str(x[0]) for x in v}, key=int))])
    g.append([])
    g.append(["D. 无最后一步热处理记录的构成"])
    g.append(["文献编号", "条数"])
    for k, v in sorted(by_paper.items(), key=lambda kv: int(kv[0])):
        g.append([k, v])
    g.append([])
    g.append(["E. 逐条明细"])
    g.append(["文献编号", "样本", "工艺原文", "解析出的最后一步", "Epit(V)"])
    for d in detail:
        g.append(list(d))
    wb.save(SRC)
    print(f"\n[write] 已写入 1.xlsx 的「热处理分组」sheet；sheets={wb.sheetnames}")


if __name__ == "__main__":
    main()
