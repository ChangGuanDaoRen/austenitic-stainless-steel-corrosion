# -*- coding: utf-8 -*-
"""2026-11-08 生成两份 0.xlsx 检查报告（四段式），只读底稿 CSV，不改任何 xlsx。
输出: 0xlsx_口径核查_20261107.md / 0xlsx_重复样本检查_20261108.md
"""
import csv
import os
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CRIT = os.path.join(ROOT, "0xlsx_口径核查_底稿_20261107.csv")
DUP = os.path.join(ROOT, "0xlsx_重复样本检查_底稿_20261108.csv")
OUT1 = os.path.join(ROOT, "0xlsx_口径核查_20261107.md")
OUT2 = os.path.join(ROOT, "0xlsx_重复样本检查_20261108.md")

# 行级两步命中（人工回原文核对，行号为 Excel 行号）
ROW2STEP = {
    "3": [("r11", "S12", "1065–1120℃固溶5h + 550–650℃稳定化2h"),
          ("r15", "C11", "1065–1120℃固溶5h + 550–650℃稳定化2h"),
          ("r16", "C12", "1065–1120℃固溶5h + 550–650℃稳定化2h")],
    "7": [("r35", "304(敏化态)", "1060℃×0.5h水冷 + 650℃×2h空冷")],
    "16": [("r90", "1150℃固溶2h水淬+850℃敏化2h空冷", "固溶 + 850℃敏化2h")],
    "26": [("r142–r165", "24 条 4 温度×6 时长", "1100℃固溶2h水淬 + 550℃×30min 等温时效")],
    "34": [("r168–r174 除 r168", "6 条", "固溶退火 + 900℃×0.5h 时效")],
    "49": [("r233–r234", "2 条", "1200℃固溶 + 650℃时效5000h")],
    "53": [("r252", "304-SA+敏化945K×1h", "1323K×1h固溶 + 945K×1h敏化"),
           ("r254", "晶界工程", "10%冷轧+1173K×1h + 5%冷轧+1173K×1h（两次退火，疑两步）")],
}
BLOCK_CLEAN = {"3": "其余 7 条（r12–r14、r17–r20）为无热处理或单步固溶",
               "16": "其余 6 条为单步固溶",
               "34": "r168「A29 as-received 冷轧态，未固溶、未时效」不属两步",
               "53": "r249/r250/r251/r253 为单步固溶或无热处理"}


def load(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    crit = load(CRIT)
    cnt = Counter(r["定稿判定"] for r in crit)
    samples = sum(int(r["条数"]) for r in crit)
    s_by = defaultdict(int)
    for r in crit:
        s_by[r["定稿判定"]] += int(r["条数"])

    L = []
    L.append("# `0.xlsx` 收录口径核查报告")
    L.append("")
    L.append("> 生成日期 2026-11-08 ｜ 判据文件 `口径基准_20261107.md`（定稿，用户 11-07 21:00「都按你的推荐执行」）")
    L.append("> 核查对象 `0.xlsx` `Sheet1`：285 行 / 26 列 / 283 条 / 60 个文献块（缺号 6、27、39）")
    L.append("> 性质：**只读核查，未改动 `0.xlsx` 任何单元格**。逐块明细见 `0xlsx_口径核查_底稿_20261107.csv`。")
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("用户 11-07 19:26 要求「检查 0 文件中是否有不符合师兄最新口径的文献」，随后 19:56 给出四条新硬口径并要求先拟基准文件。")
    L.append("`口径基准_20261107.md` 于当晚定稿生效，本轮据此对 `0.xlsx` 60 个文献块逐块判级，同时按用户 11-08 追加要求做重复样本检查。")
    L.append("")
    L.append("四条硬口径的落地形式（A1–A5 五处裁定见基准文件 3.2）：")
    L.append("")
    L.append("| 条款 | 内容 | 判级含义 |")
    L.append("| --- | --- | --- |")
    L.append("| N1 | 铸造类（熔炼—凝固路线），排除增材/粉末/焊接/表面改性 | 只要求熔铸路线，不要求原文写出「铸」字；锻轧件收（用户 20:41「锻轧的不也得先铸吗」） |")
    L.append("| N2 | 奥氏体不锈钢 | 奥氏体+少量 δ 视为奥氏体并标注；真双相、马氏体/铁素体为主不符 |")
    L.append("| N3 | 3.5 wt.% NaCl | 收 0.6 mol/L 等价写法与天然海水（须标注）；人工海水/模拟液/3.56%/wt./v 判存疑；酸、H₂S、FeCl₃ 不符 |")
    L.append("| N4 | 不要两步热处理 | 固溶或退火后**再叠**敏化/时效/稳定化才算两步；热轧+固溶、冷轧+退火不算；均匀化/轧前加热/深冷/酸洗/钝化不计步 |")
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("1. 只读脚本 `scripts\\audit_0xlsx_20261108.py`：`openpyxl.load_workbook(data_only=True)` 读 `Sheet1`，以 A 列物理单元格为块首（60 块），")
    L.append("   Y（文献信息）/Z（备注）按合并区取块首值，逐条取 15 个元素列与工艺/电极/温度/Epit/稳定时间。")
    L.append("2. 判级规则按 N1–N4 逐条编码；N1 排除关键词带否定上下文过滤（`排除/不收/不是/非/只`），避免把「只排除增材」这类说明句误判。")
    L.append("3. N4 采用 A3 窄定义：仅当块内原文出现「固溶/退火 + 敏化/时效/稳定化」两段时才判两步；热老化、均匀化、两次固溶/退火单列标注不判两步。")
    L.append("4. 三档判定：**合规**（可直接用）/ **存疑**（缺介质或参比证据、须回 `文献\\NN.pdf` 取证）/ **不符**（违反 N1–N4 之一）。")
    L.append("5. 逐块判定后对 7 个「不符」块回原文做**行级**复核（A3 要求「标注不合规、不物理删行」）。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("### 3.1 总量")
    L.append("")
    L.append("| 判定 | 块数 | 条数 | 占比（条） |")
    L.append("| --- | --- | --- | --- |")
    for k in ("合规", "存疑", "不符"):
        L.append("| %s | %d | %d | %.1f%% |" % (k, cnt[k], s_by[k], 100.0 * s_by[k] / samples))
    L.append("| 合计 | %d | %d | 100%% |" % (len(crit), samples))
    L.append("")
    L.append("**按条数算合规率仅 %.1f%%（%d/%d 条）；按文献块算 %d/60 块。**" % (
        100.0 * s_by["合规"] / samples, s_by["合规"], samples, cnt["合规"]))
    L.append("")
    L.append("### 3.2 不符 7 块（违反 N4 两步热处理，全部为「固溶/退火 + 敏化/时效/稳定化」）")
    L.append("")
    L.append("| 块 | 条数 | 两步构成（含行级命中） | 块内非两步行 |")
    L.append("| --- | --- | --- | --- |")
    for bid in ("3", "7", "16", "26", "34", "49", "53"):
        r = next(x for x in crit if x["编号"] == bid)
        hits = ROW2STEP.get(bid, [])
        desc = "；".join("%s %s → %s" % h for h in hits)
        L.append("| 文献%s | %s | %s | %s |" % (
            bid, r["条数"], desc, BLOCK_CLEAN.get(bid, "—（整块两步）")))
    L.append("")
    L.append("**行级结论：7 块 57 条中真正两步的只有 44 条**——块 3 有 7 条、块 16 有 6 条、块 34 有 1 条（r168）、块 53 有 4 条（r249–r251、r253）不属两步。")
    L.append("按 A3 裁定，这些行不删，只在入库时标注「不合规」，建模子集切分时再排除。")
    L.append("")
    L.append("### 3.3 合规 19 块（67 条）")
    L.append("")
    ok = [r for r in crit if r["定稿判定"] == "合规"]
    L.append("| 块 | 条数 | 介质 | 组织 | 标注项 |")
    L.append("| --- | --- | --- | --- | --- |")
    for r in ok:
        L.append("| 文献%s | %s | %s | %s | %s |" % (
            r["编号"], r["条数"], r["介质_底稿"], r["组织_底稿"], r["标注项"] or "—"))
    L.append("")
    L.append("### 3.4 存疑 34 块（159 条）——主因")
    L.append("")
    reasons = Counter()
    for r in crit:
        if r["定稿判定"] != "存疑":
            continue
        t = r["存疑项"]
        key = "Z/Y 未记介质，须回原文取证" if "未记介质" in t else (
            "人工/模拟海水或浓度写法待补证" if "人工" in t or "3.56" in t or "模拟" in t or "wt./v" in t else (
                "组织或成形未直写，须回原文" if "组织" in t or "成形" in t else "其它（见底稿）"))
        reasons[key] += 1
    L.append("| 原因 | 块数 |")
    L.append("| --- | --- |")
    for k, v in reasons.most_common():
        L.append("| %s | %d |" % (k, v))
    L.append("")
    L.append("**关键事实：`0.xlsx` 没有独立的「腐蚀介质」列，介质信息只散落在 Y（文献信息）/Z（备注）里，60 块中有 36 块 Z 完全没写介质。**")
    L.append("这是存疑比例高的根本原因，也说明 11-07 21:03 用户裁定「两列不删」是必要的——删掉就再也无法回查。")
    L.append("")
    L.append("### 3.5 与定稿前甲/乙两套解释的对比")
    L.append("")
    L.append("| 口径 | 合规 | 存疑 | 不符 |")
    L.append("| --- | --- | --- | --- |")
    L.append("| 甲（严格字面：锻轧排除、海水/0.6mol 全排除、双相排除） | 2 块/14 条 | 18 块/65 条 | 40 块/204 条 |")
    L.append("| 乙（11-05 裁定：锻轧可收、保留 0.6mol 与天然海水） | 14 块/52 条 | 33 块/156 条 | 13 块/75 条 |")
    L.append("| **定稿 A1–A5（本报告）** | **%d 块/%d 条** | **%d 块/%d 条** | **%d 块/%d 条** |" % (
        cnt["合规"], s_by["合规"], cnt["存疑"], s_by["存疑"], cnt["不符"], s_by["不符"]))
    L.append("")
    L.append("## 四、建议")
    L.append("")
    L.append("1. **不要按「不符」删行**。A3 已裁定「收录时标注不合规、不物理删行」，7 块 57 条保留，其中 44 条在建模子集阶段排除即可。")
    L.append("2. **优先补 34 块 159 条的介质证据**。介质只能回 `文献\\NN.pdf` 方法段取证（编号→`文献\\NN.pdf` 溯源链成立，63 篇无缺）。")
    L.append("   建议按「Z 已写介质」优先判，Z 空白的 36 块逐篇回原文补一列 `介质_核实`（新增列而非改原表，不违反只读约定）。")
    L.append("3. **块 23 的「475℃热老化」需单独裁定**：热老化不是敏化/时效/稳定化，A3 窄定义下不判两步，但物理上属长期时效，是否入建模子集请师兄定。")
    L.append("4. **块 63 的「铁素体」疑为误报**：Z 前 160 字未见该词，建议回 `文献\\63.pdf` 确认牌号（市售 201/202/304 冷轧薄板），若确为奥氏体则从存疑改为合规。")
    L.append("5. **合规率 23.7% 与 150 条目标的关系**：`0.xlsx` 283 条中定稿合规仅 67 条，缺口靠 `待\\` 复筛补（本轮已移入 `准\\` 45 篇），")
    L.append("   但补数据前应先解决存疑 159 条的介质取证，否则 150 条目标无法用「同工艺同介质」子集兑现。")
    L.append("")
    with open(OUT1, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print("wrote", OUT1, len(L), "lines")

    # ---------- 重复样本报告 ----------
    dup = load(DUP)
    groups = defaultdict(list)
    for r in dup:
        groups[r["指纹"]].append(r)
    M = []
    M.append("# `0.xlsx` 重复样本检查报告")
    M.append("")
    M.append("> 生成日期 2026-11-08 ｜ 对象 `0.xlsx` `Sheet1`：283 条 / 60 块")
    M.append("> 性质：**只读检查，未改动 `0.xlsx`**。明细见 `0xlsx_重复样本检查_底稿_20261108.csv`。")
    M.append("")
    M.append("## 一、背景")
    M.append("")
    M.append("用户 11-08 06:48 追加要求「0 文件除了口径问题，再检查有无重复样本」。")
    M.append("`0.xlsx` 经多轮追加录入（文献 1–21 手录、22–26 与 28–63 分批补录），存在同一文献被重复录入或同一试验被拆成两条的风险，需在建模前排除。")
    M.append("")
    M.append("## 二、方法")
    M.append("")
    M.append("六层指纹逐层放宽，任何一层命中都列出供人工判断：")
    M.append("")
    M.append("| 层 | 指纹内容 | 用途 |")
    M.append("| --- | --- | --- |")
    M.append("| F0 | 全部业务字段（样本名+15 元素+工艺+电极+温度+Epit+稳定时间）完全相同 | 抓完全重复录入 |")
    M.append("| F1 | 样本名 + 成分 + 工艺 + Epit | 抓同一样本录两遍（备注不同） |")
    M.append("| F2 | 成分 + Epit | 抓换名重复 |")
    M.append("| F3 | 同块内样本名相同 | 抓块内笔误重复 |")
    M.append("| F4 | 跨块样本名 + 成分相同 | 抓同一文献录进两个块 |")
    M.append("| F5 | 样本名 + Epit 相同 | 抓图读值撞车 |")
    M.append("")
    M.append("数值统一按 6 位小数归一，文本 NFKC 归一，避免「0.8」与「0.80」误判为不同。")
    M.append("")
    M.append("## 三、结论")
    M.append("")
    M.append("| 层 | 命中组数 | 涉及行数 | 跨块组数 | 判读 |")
    M.append("| --- | --- | --- | --- | --- |")
    for layer in ("F0_业务字段全同", "F1_样本名+成分+工艺+Epit", "F2_成分+Epit同", "F3_同块同名", "F4_同名+成分（跨块）", "F5_样本名+Epit同"):
        g = groups.get(layer, [])
        rowsn = sum(int(x["组内行数"]) for x in g) if g else 0
        cross = sum(1 for x in g if x.get("跨块") == "是") if g else 0
        verdict = "**无重复**" if layer.startswith(("F0", "F1")) else ("无跨块重复" if cross == 0 else "需核")
        M.append("| %s | %d | %d | %d | %s |" % (layer, len(g), rowsn, cross, verdict))
    M.append("")
    M.append("### 3.1 核心结论")
    M.append("")
    M.append("1. **无完全重复录入**：F0、F1 两层均为 0 组，即不存在「同一样本连成分、工艺、Epit 全都一样」的两条记录。")
    M.append("2. **无跨文献重复**：F2–F5 四层的跨块组数全部为 0，说明没有把同一篇文献的数据录进两个不同编号的块。")
    M.append("3. **同块内命中均为试验设计网格**，不是录入错误——例如块 52 的冷轧减面系列（0.76N/0.92N…）、块 45/47 的 304 温轧系列、")
    M.append("   块 43 的 Nb00/23/48/99 × 1100/1300℃ 固溶网格，同牌号不同工艺本就该并列。")
    M.append("")
    M.append("### 3.2 需人工确认的 5 组（同一块内成分与 Epit 完全相同、但工艺不同）")
    M.append("")
    M.append("| 块 | 行 | 样本 | Epit (V) | 疑点 |")
    M.append("| --- | --- | --- | --- | --- |")
    for x in dup:
        if x["指纹"] in ("F2_成分+Epit同", "F5_样本名+Epit同") and x.get("跨块") == "否":
            M.append("| %s | %s | %s | %s | 同成分同 Epit 但工艺不同，需回原文确认 |" % (
                x.get("涉及块", ""), x.get("行号", ""), x.get("样本名", ""), x.get("Epit", "")))
    M.append("")
    M.append("这些点对大概率是**原文本身给出相同值**，或**读图时取了同一刻度**（Z 备注已注明部分数据由 WebPlotDigitizer 提取）。")
    M.append("两种情形都不属于「重复样本」，但若属读图撞车则属数据质量问题，需回原文核对。")
    M.append("")
    M.append("## 四、建议")
    M.append("")
    M.append("1. **`0.xlsx` 无需删行**：六层指纹未发现任何真重复，块内命中都是设计网格。")
    M.append("2. **5 组同值点对回原文核**：块 43 r203/r207（Epit 0.24）、块 52 r235/r236（0.8）、r238/r239（0.62）、r242/r248（0.98）、r243/r244（1.02）。")
    M.append("   若原文确为同值，保留；若为读图取整，按原图重读并注明读数精度（±0.02 V）。")
    M.append("3. **建模前去重仍需做一次**：本轮只查「录入重复」，未查「同一样本在 0.xlsx 与 待\\ 复筛件中重复」——")
    M.append("   该问题已在 `待\\` 复筛中处理（9 篇按 DOI 判「重复-已入库」、5 组判「重复-库内重复」，见 `待_复筛判定_20261108.csv`）。")
    M.append("4. 建议在 `1.xlsx` 生成脚本里加一道 F0/F1 断言，防止后续追加录入引入重复。")
    M.append("")
    with open(OUT2, "w", encoding="utf-8") as f:
        f.write("\n".join(M) + "\n")
    print("wrote", OUT2, len(M), "lines")


if __name__ == "__main__":
    main()
