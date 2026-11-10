# -*- coding: utf-8 -*-
"""2026-11-08 生成 待\\ 复筛报告（四段式）。只读 CSV，不移动文件。
输出: 待_复筛报告_20261108.md
"""
import csv
import os
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "待_复筛判定_20261108.csv")
LOG = os.path.join(ROOT, "准_移入台账_20261108.csv")
OUT = os.path.join(ROOT, "待_复筛报告_20261108.md")


def load(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    rows = load(SRC)
    moved = {r["文件名"] for r in load(LOG)} if os.path.isfile(LOG) else set()
    c = Counter(r["判定"] for r in rows)
    n = len(rows)
    L = []
    L.append("# `待\\` 文件夹按定稿口径复筛报告")
    L.append("")
    L.append("> 生成日期 2026-11-08 ｜ 判据 `口径基准_20261107.md`（定稿）")
    L.append("> 对象 `待\\` 根目录 **%d 篇 PDF**（1156.9 MB）｜ 对照库 `文献\\` 63 篇" % n)
    L.append("> 逐篇明细 `待_复筛判定_20261108.csv`（262 行，UTF-8-SIG）；移入台账 `准_移入台账_20261108.csv`")
    L.append("> **本轮只移动了判「合规」的 %d 篇，其余 %d 篇原样留在 `待\\`，未删除、未改名。**" % (c["合规"], n - c["合规"]))
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("`待\\` 是历轮检索（第 8/9/10 轮及更早）筛剩的待定件，按旧口径判为不合格或存疑，从未按 11-07 定稿的四条新硬口径复筛过。")
    L.append("用户 11-08 06:48 要求「按照这一口径继续核查待中所有文件，符合的放入准文件夹」，`准\\` 当时为空（0 篇），故合规件直接移入。")
    L.append("")
    L.append("四条硬口径：N1 铸造类（熔炼—凝固路线，排除增材/粉末/焊接/表面改性）；N2 奥氏体不锈钢；N3 3.5 wt.% NaCl；N4 不要两步热处理。")
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("原计划开 6 个子智能体并行（`files[i::6]` 六切片），6 个全部在 07:20–07:21 中断失败，只留下中间件。改为**自建可重跑流水线**：")
    L.append("")
    L.append("| 步 | 脚本 | 产出 |")
    L.append("| --- | --- | --- |")
    L.append("| 1 抽文本 | `scripts\\extract_text_20261108.py` | `temp\\rescreen_20261108\\text_all.jsonl`（325 篇 = 待 262 + 文献 63） |")
    L.append("| 2 规则初筛 | `scripts\\triage_20261108.py` | `temp\\rescreen_20261108\\triage_all.csv`（262 行） |")
    L.append("| 3 抽证据 | `scripts\\dump_doubt_20261108.py` | `review_doubt_A.txt`(39) / `review_doubt_B.txt`(38) |")
    L.append("| 4 人工裁定 | `scripts\\decide_20261108.py` | `待_复筛判定_20261108.csv` |")
    L.append("| 5 移入 | `scripts\\move_compliant_20261108.py run` | `准\\` + `准_移入台账_20261108.csv` |")
    L.append("")
    L.append("要点：PyMuPDF 抽全文 → 规则层按 N1–N4 打标签 → **误报降级**（排除关键词若只出现在参考文献、未在标题/摘要前 1500 字命中，降为「存疑」而不是「不符」）")
    L.append("→ 对 77 篇存疑件逐篇读证据段落人工裁定 → 45 篇合规逐篇写一句判定依据。")
    L.append("")
    L.append("误报规模：初筛「不符」176 篇中，双相 45 篇里仅 23 篇标题/摘要真命中、涂层 60 篇里仅 10 篇、增材 49 篇里仅 8 篇，其余多为参考文献提及。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("### 3.1 判定分布（262 篇）")
    L.append("")
    L.append("| 判定 | 篇数 | 占比 | 处置 |")
    L.append("| --- | --- | --- | --- |")
    disp = {"合规": "已移入 `准\\`", "存疑-待补证": "留在 `待\\`，待补介质/参比证据",
            "不符": "留在 `待\\`", "重复-已入库": "留在 `待\\`（`文献\\` 已有同篇）",
            "重复-库内重复": "留在 `待\\`（`待\\` 内另有同 DOI 件）"}
    for k in ("合规", "存疑-待补证", "不符", "重复-已入库", "重复-库内重复"):
        L.append("| %s | %d | %.1f%% | %s |" % (k, c[k], 100.0 * c[k] / n, disp[k]))
    L.append("| 合计 | %d | 100%% | — |" % n)
    L.append("")
    L.append("### 3.2 已移入 `准\\` 的 %d 篇" % c["合规"])
    L.append("")
    L.append("| # | 文件 | 介质 | 成形 | 热处理 | 判定依据 |")
    L.append("| --- | --- | --- | --- | --- | --- |")
    for i, r in enumerate([x for x in rows if x["判定"] == "合规"], 1):
        f = r["文件名"]
        if len(f) > 46:
            f = f[:43] + "…"
        L.append("| %d | %s | %s | %s | %s | %s |" % (
            i, f.replace("|", "/"), (r["介质证据"] or "—").replace("|", "/")[:34],
            (r["成形证据"] or "—").replace("|", "/")[:26],
            (r["热处理证据"] or "—").replace("|", "/")[:26],
            (r["依据"] or "—").replace("|", "/")[:60]))
    L.append("")
    L.append("### 3.3 重复件 %d 篇" % (c["重复-已入库"] + c["重复-库内重复"]))
    L.append("")
    L.append("| 文件 | 判定 | 说明 |")
    L.append("| --- | --- | --- |")
    for r in rows:
        if r["判定"].startswith("重复"):
            L.append("| %s | %s | %s |" % (r["文件名"].replace("|", "/"), r["判定"], (r["依据"] or "").replace("|", "/")[:70]))
    L.append("")
    L.append("### 3.4 不符 %d 篇——按触发条款" % c["不符"])
    L.append("")
    tb = Counter()
    for r in rows:
        if r["判定"] != "不符":
            continue
        t = r["依据"]
        for key in ("N1-增材", "N1-表面改性", "N1-粉末", "N1-焊接", "N2-双相", "N2-马氏体", "N2-铁素体", "N3", "N4"):
            if key in t:
                tb[key] += 1
                break
        else:
            tb["其它"] += 1
    L.append("| 触发条款 | 篇数 |")
    L.append("| --- | --- |")
    for k, v in tb.most_common():
        L.append("| %s | %d |" % (k, v))
    L.append("")
    L.append("明细（文件名 + 触发理由）：")
    L.append("")
    L.append("| 文件 | 理由 |")
    L.append("| --- | --- |")
    for r in rows:
        if r["判定"] == "不符":
            L.append("| %s | %s |" % (r["文件名"].replace("|", "/"), (r["依据"] or "").replace("|", "/")[:60]))
    L.append("")
    L.append("### 3.5 存疑 %d 篇——分类" % c["存疑-待补证"])
    L.append("")
    sc = Counter()
    for r in rows:
        if r["判定"] != "存疑-待补证":
            continue
        t = r["依据"]
        if "人工海水" in t or "模拟" in t or "SME" in t:
            sc["人工/模拟海水，A2 待补证"] += 1
        elif "参比" in t or "Ag/AgCl" in t or "SSCE" in t:
            sc["参比电极待换算/未标明"] += 1
        elif "Tafel" in t:
            sc["Epit 来源为 Tafel 外推，须核拐点"] += 1
        elif "介质" in t or "3.5" in t:
            sc["介质表述需回原文确认"] += 1
        elif "两步" in t or "时效" in t or "敏化" in t:
            sc["疑似两步热处理，须核"] += 1
        else:
            sc["其它（组织/成形/文本质量）"] += 1
    L.append("| 类别 | 篇数 |")
    L.append("| --- | --- |")
    for k, v in sc.most_common():
        L.append("| %s | %d |" % (k, v))
    L.append("")
    L.append("## 四、建议")
    L.append("")
    L.append("1. **`准\\` 45 篇先入库取证**：这批已逐篇核过介质、成形、热处理与参比，可直接按 `0.xlsx` 版式录入，预计可补 130–180 条。")
    L.append("2. **存疑 " + str(c["存疑-待补证"]) + " 篇优先级**：先处理「人工/模拟海水」与「参比未标明」两类（各占大头），前者需师兄裁定是否等同 3.5 wt.% NaCl，后者只需回原文换算式。")
    L.append("3. **重复 9 篇建议物理清理**：与 `文献\\` 同一 DOI 的 9 篇属重复投递，可按 `scripts\\trash_delete.py` 送回收站（须用户确认后再做）。")
    L.append("4. **库内重复 5 组保留前者**：同 DOI 两份只留一份，另一份标「重复-库内重复」，建议同样送回收站（须确认）。")
    L.append("5. **不要一次性把 `待\\` 清空**：%d 篇不符件是历轮检索的负面样本，保留可避免后续重复检索同一批论文。" % c["不符"])
    L.append("6. **子智能体不再用于长流程**：本次 6 个全部中断，自建流水线可随时重跑（`extract → triage → decide → move`），建议后续沿用。")
    L.append("")
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print("wrote", OUT, len(L), "lines")


if __name__ == "__main__":
    main()
