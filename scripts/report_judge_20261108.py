# -*- coding: utf-8 -*-
"""由 逐篇判级_20261108.csv 生成四段式报告 逐篇判级报告_20261108.md。

用法：D:\\Python\\python.exe -X utf8 scripts\\report_judge_20261108.py
"""
import csv
import os
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "逐篇判级_20261108.csv")
OUT = os.path.join(ROOT, "逐篇判级报告_20261108.md")


def load():
    with open(SRC, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main():
    rows = load()
    c = Counter(r["判定"] for r in rows)
    total = len(rows)
    good = [r for r in rows if r["判定"] in ("合规", "合规-基准组")]
    doubt = [r for r in rows if r["判定"] == "存疑-待补证"]
    bad = [r for r in rows if r["判定"] in ("不符", "重复-已入库")]

    clause = Counter()
    for r in bad:
        t = r["触发条款"].split(" ")[0].split("；")[0].strip() or "其它"
        clause[t] += 1
    dclause = Counter()
    for r in doubt:
        t = r["触发条款"].split(" ")[0].split("；")[0].strip() or "其它"
        dclause[t] += 1

    L = []
    L.append("# 待复筛 262 篇逐篇精读判级报告（2026-11-08）\n")
    L.append("> 口径依据：`口径基准_20261107.md` 定稿 N1–N4（用户 11-07 21:00 批准）。")
    L.append("> 方法：每篇起一个只读判级子智能体，读该篇全文证据摘要（标题/摘要 + 介质·成分·电位·参比·工艺·组织关键词原文段 + 结尾），逐篇写判定 JSON，再由本脚本汇总。")
    L.append("> 落地：合规 → `准\\`；存疑-待补证 → `待\\`；不符/重复 → `待\\_不符合\\`。只移动不删除，台账 `逐篇判级_移动台账_20261108.csv`。\n")

    L.append("## 一、背景\n")
    L.append("`待\\` 与 `准\\` 共 262 篇候选（11-08 复筛后的全部未入库 PDF）。上一轮用「规则初筛 + 人工抽核」，用户要求改为**一篇一篇全部精读**，因此本轮对 262 篇逐篇独立判级，重排三分类。\n")

    L.append("## 二、方法\n")
    L.append("1. `scripts\\prep_judge_20261108.py` 从 `temp\\rescreen_20261108\\text_all.jsonl` 为 262 篇各生成一份证据摘要（平均 11.8 千字符），3 篇文字层抽取失败的用 `scripts\\ocr_failed_20261108.py` 走 RapidOCR 300 dpi 重抽。")
    L.append("2. 262 篇各起一个只读子智能体，按 N1–N4 判级，必须抄录「介质原文 / 成形与组织原文 / 热处理原文 / Epit 与参比」四项证据，写 `temp\\judge_20261108\\verdicts\\pNNN.json`。")
    L.append("3. `scripts\\apply_verdicts_20261108.py` 汇总为 `逐篇判级_20261108.csv`，并按三分类移动文件、写回滚台账。\n")

    L.append("## 三、结论\n")
    L.append("| 判定 | 篇数 | 占比 | 落地位置 |")
    L.append("|---|---|---|---|")
    for k, where in [("合规", "`准\\`"), ("合规-基准组", "`准\\`"), ("存疑-待补证", "`待\\`"),
                     ("不符", "`待\\_不符合\\`"), ("重复-已入库", "`待\\_不符合\\`"),
                     ("重复-库内重复", "`待\\_不符合\\`")]:
        n = c.get(k, 0)
        L.append("| %s | %d | %.1f%% | %s |" % (k, n, 100.0 * n / total if total else 0, where))
    L.append("| **合计** | **%d** | 100%% | |\n" % total)

    L.append("### 3.1 判「不符 / 重复」的触发条款分布\n")
    L.append("| 条款 | 篇数 |")
    L.append("|---|---|")
    for k, n in clause.most_common():
        L.append("| %s | %d |" % (k, n))
    L.append("")

    L.append("### 3.2 判「存疑-待补证」的触发条款分布\n")
    L.append("| 条款 | 篇数 |")
    L.append("|---|---|")
    for k, n in dclause.most_common():
        L.append("| %s | %d |" % (k, n))
    L.append("")

    L.append("### 3.3 判「合规 / 合规-基准组」逐篇\n")
    L.append("| # | 文件 | 介质 | 成形与组织 | 热处理 | Epit 与参比 | 依据 |")
    L.append("|---|---|---|---|---|---|---|")
    for r in good:
        L.append("| %s | %s | %s | %s | %s | %s | %s |" % (
            r["idx"], r["文件"], r["介质原文"][:40], r["成形与组织原文"][:40],
            r["热处理原文"][:40], r["Epit与参比"][:40], r["依据"][:60]))
    L.append("")

    L.append("### 3.4 判「不符 / 重复」逐篇\n")
    L.append("| # | 文件 | 条款 | 依据 |")
    L.append("|---|---|---|---|")
    for r in bad:
        L.append("| %s | %s | %s | %s |" % (r["idx"], r["文件"], r["触发条款"][:50], r["依据"][:60]))
    L.append("")

    L.append("### 3.5 判「存疑-待补证」逐篇（前 60 条，全量见 CSV）\n")
    L.append("| # | 文件 | 条款 | 依据 |")
    L.append("|---|---|---|---|")
    for r in doubt[:60]:
        L.append("| %s | %s | %s | %s |" % (r["idx"], r["文件"], r["触发条款"][:50], r["依据"][:60]))
    L.append("")

    L.append("## 四、建议\n")
    L.append("1. **存疑件优先补介质**：存疑里以「介质写法非 3.5 wt.% 原样」和「摘要证据不足」为主，回原文方法段/表格核一次即可大批转正。")
    L.append("2. **不符件集中复核 N1 与 N2**：增材/焊接/涂层与双相/马氏体是主要排除原因，抽 10 篇回原文复核，确认不是关键词误命中。")
    L.append("3. **合规件进入取证排队**：合规件按 `文献\\` 编号顺序取证入库，先出 `final_rows` 校验件再写 `0.xlsx`/`1.xlsx`。")
    L.append("4. **重复件不删**：`待\\_不符合\\` 内的重复件保留原件，等入库时再统一处理。")
    L.append("")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("写出", OUT, "行数", len(L))


if __name__ == "__main__":
    main()
