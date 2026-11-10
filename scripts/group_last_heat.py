# -*- coding: utf-8 -*-
"""按「最后一步热处理的温度+时间」给 116 条数据分组（师兄 2026-10-26 20:5x 的口径）。

用法：python scripts/group_last_heat.py [dry]
只读 _清洗版_v2.xlsx 的 01_清洗总表，不改任何文件。
"""
import re
import sys
import collections
import openpyxl

SRC = "铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版_v2.xlsx"

# 温度：单个（1100℃）或区间（1065-1120℃）；时间：单个（2h / 30min）或区间（0.5-5h）
STEP_RE = re.compile(
    r"(\d{3,4}(?:\s*[-~]\s*\d{3,4})?)\s*℃[^0-9]{0,14}?"
    r"(\d+(?:\.\d+)?(?:\s*[-~]\s*\d+(?:\.\d+)?)?)\s*(h|小时|min|分钟)"
)


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
        return f"{v:g}min"
    return f"{v:g}h"


def steps(text):
    out = []
    for t, tm, u in STEP_RE.findall(text):
        out.append((norm_temp(t), norm_time(tm, u)))
    return out


def main():
    ws = openpyxl.load_workbook(SRC, data_only=True)["01_清洗总表"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[1]]

    groups = collections.defaultdict(list)
    for r in rows:
        st = steps(str(r[17]))
        key = f"{st[-1][0]}℃×{st[-1][1]}" if st else "无热处理记录（或未写温度时间）"
        groups[key].append(r)

    print(f"总条数 {len(rows)}\n")
    print(f"{'最后一步热处理':<30}{'条数':>4}  文献")
    print("-" * 78)
    for k, v in sorted(groups.items(), key=lambda x: -len(x[1])):
        papers = sorted({str(x[0]) for x in v}, key=int)
        print(f"{k:<30}{len(v):>4}  {','.join(papers)}")

    with_steps = sum(len(v) for k, v in groups.items() if k != "无热处理记录（或未写温度时间）")
    strict = [(k, v) for k, v in groups.items() if k != "无热处理记录（或未写温度时间）"]
    strict.sort(key=lambda x: -len(x[1]))
    print(f"\n有最后一步热处理记录的：{with_steps} 条，分 {len(strict)} 组")
    print(f"最大一组：{strict[0][0]} = {len(strict[0][1])} 条")
    print(f"≥5 条的组：{[(k, len(v)) for k, v in strict if len(v) >= 5]}")
    print(f"≥2 条的组共 {sum(len(v) for k, v in strict if len(v) >= 2)} 条 / "
          f"{len([1 for k, v in strict if len(v) >= 2])} 组")


if __name__ == "__main__":
    main()
