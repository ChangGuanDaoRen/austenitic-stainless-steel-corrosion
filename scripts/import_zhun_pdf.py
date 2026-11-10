# -*- coding: utf-8 -*-
r"""把「准」15 篇 PDF 以新编号复制进 文献\，并生成 准文献入库对照_20261104.md。

用法：
  D:/Python/python.exe scripts/import_zhun_pdf.py            只做前置检查并打印计划
  D:/Python/python.exe scripts/import_zhun_pdf.py --commit   执行复制 + 出对照清单

为什么复制而不是移动：用户要求「重编号放入文献文件夹」，但 准\ 原件是筛选链的证据，
不能删、不能改名（USER.md 第三节）。编号 27 已被剔除的 C002 占用，故新编号从 28 起。
覆盖保护：目标 文献\<no>.pdf 若已存在立即中止，绝不覆盖已有文献。
"""
import hashlib
import json
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
DATA = os.path.join(WS, "temp", "lit", "entry", "final_rows.json")
MANIFEST = os.path.join(WS, "准文献入库对照_20261104.md")

# 口径提示：从 final_rows.json 的 notes 归纳，不新增推断
FLAGS = {
    28: "Epit 仅见于 p6 Fig.3(a)，正文无数值 → 3 条记 N/A；与 B 类清单同文（见建议）",
    29: "热轧+固溶；成分与 Epit 均为文字表",
    30: "表1/表2 是图片，200/400 dpi 渲染放大逐格读数；SCE 参比",
    31: "MnS 夹杂形貌；参比 Ag/AgCl（原文未注 KCl 浓度）",
    32: "原文只给统计分布与 Fig.7 累积概率，无逐试样 Epit → 1 条记 N/A",
    33: "超高点蚀抗性 316LN 反复冷轧退火，取 1 条",
    34: "【口径存疑·重要】固溶水淬后另有 900℃×0.5h 时效，属两步热处理；成分含 < 上限值",
    35: "310S 深冷轧+退火，2 条",
    36: "同时有 SME(3.5 wt% NaCl) 与 SBF 介质，只录 SME 中退火 1 条",
    37: "304/316 温预应变，2 条",
    38: "【口径存疑】3.56 wt% NaCl(ASTM G61)；Table 1 缺 Si/P，Mo 后 3 组归属不明记「未提及」",
    39: "【待裁决】主体为 600–1000℃ 时效系列，只取 1250℃×3h 固溶态；pH 3.0 与 3.5% 表述矛盾",
    40: "纯铸态 304 的 Mg/S 含量系列，9 条，是本轮唯一的铸造态文献",
    41: "【口径存疑】As received 两组为热轧原始态；参比仅写 Ag/AgCl",
    42: "304 深冷轧+再结晶退火，2 条",
}

# 行范围（0.xlsx 写入结果，来自 scripts\append_0xlsx.py 实测）
ROWS = {
    28: (160, 162), 29: (163, 164), 30: (165, 168), 31: (169, 170),
    32: (171, 171), 33: (172, 172), 34: (173, 179), 35: (180, 181),
    36: (182, 182), 37: (183, 184), 38: (185, 190), 39: (191, 191),
    40: (192, 200), 41: (201, 206), 42: (207, 208),
}


def md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def fmt_epit(vals):
    out = []
    for v in vals:
        out.append("N/A" if v == "N/A" else ("%g" % v))
    return " / ".join(out)


def build_manifest(data, copied):
    total = sum(len(p["rows"]) for p in data)
    na = sum(1 for p in data for r in p["rows"] if r["Epit"] == "N/A")
    nums = [r["Epit"] for p in data for r in p["rows"] if isinstance(r["Epit"], (int, float))]
    lines = []
    lines.append("## 一、背景")
    lines.append("")
    lines.append("- 用户指令：「现在先把准里的文献按照格录入0文件，然后把相应文献重编号放入文献文件夹。」"
                 "录入范围经确认是「15 篇全录」——含 11 篇确定可录、1 篇待裁决、2 篇待补、1 篇纯铸态漏检篇，"
                 "口径冲突一律写进 `0.xlsx` 的「备注」列，不因待裁决停下。")
    lines.append("- `0.xlsx` 是用户手录的原文基准表（Epit 写原文数值、参比照原文、不做 SCE 换算），"
                 "追加前为 159 行 / 157 条 / 文献编号 1–26；`文献\\27.pdf` 是已因 3.56% NaCl 整篇剔除的 C002，"
                 "文件保留但数据已删，所以新编号从 **28** 起，不复用 27。")
    lines.append("- 本轮只做「入基准表 + PDF 重编号归档」两件事；`1.xlsx` / `2.xlsx` 清洗版与建模子集不在本轮范围内。")
    lines.append("")
    lines.append("## 二、方法")
    lines.append("")
    lines.append("- 取证：`scripts\\extract_zhun.py` 把 `准\\` 15 篇 PDF 逐页抽文本到 `temp\\lit\\entry\\*.txt`；"
                 "图片表格与曲线用 `scripts\\render_pages.py` 渲染 200–600 dpi 后人工逐格核对；"
                 "三组并行核定结果 `data_g1/g2/g3.json` 经合并、补录、修正后定稿为 "
                 "`temp\\lit\\entry\\final_rows.json`（每条带 `source_page` / `source_quote`）。")
    lines.append("- 写表：`scripts\\append_0xlsx.py` 从 `备份\\0_基准_追加前_20261104.xlsx` 读原件，"
                 "沿用宋体 11 / 细边框 / 全居中 / 两行表头，A、Z、AA 按文献块合并；"
                 "写回前后各校验一次，前 159 行逐格零差异、原 80 个合并区全保留。")
    lines.append("- 归档：`scripts\\import_zhun_pdf.py` 以 `shutil.copy2` 复制 PDF，**不移动、不删除、不改名 `准\\` 原件**；"
                 "每个目标文件复制前检查不存在，复制后逐个比对 MD5。")
    lines.append("- 数据来源均为本工作区 PDF 原文与上述中间文件；`0.xlsx` 追加前后备份在 "
                 "`备份\\0_基准_追加前_20261104.xlsx`（与追加前 `0.xlsx` 字节一致）。")
    lines.append("")
    lines.append("## 三、结论")
    lines.append("")
    lines.append("- `0.xlsx` 现为 **208 行 / 206 条**（157 条 + 49 条），新增数据占第 **160–208 行**；"
                 "新增文献 15 篇、编号 **28–42**，其中 45 条有 Epit 数值、4 条记 `N/A`，"
                 "数值范围 %g ~ %g V（基准表按原文参比，未换算 SCE）。" % (min(nums), max(nums)))
    lines.append("- `文献\\` 新增 **28.pdf–42.pdf** 共 15 个文件，复制后 MD5 与 `准\\` 原件逐一相同；"
                 "原有 `01.pdf–27.pdf` 未改动。")
    lines.append("")
    lines.append("### 3.1 入库对照表")
    lines.append("")
    lines.append("| 新编号 | 新文件名 | 原文件名（准\\） | 题名（文献信息首行） | 条数 | 0.xlsx 行 | Epit(V) | 口径提示 |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- |")
    for p in data:
        no = p["no"]
        title = p["citation"].split("\n")[0].strip().replace("|", "\\|")
        s, e = ROWS[no]
        row = "| %d | `%02d.pdf` | `%s` | %s | %d | %d–%d | %s | %s |" % (
            no, no, p["file"].replace("|", "\\|"), title,
            len(p["rows"]), s, e, fmt_epit([r["Epit"] for r in p["rows"]]), FLAGS[no])
        lines.append(row)
    lines.append("")
    lines.append("### 3.2 风险与未决问题")
    lines.append("")
    lines.append("- **4 条 Epit 记 `N/A`**：文献 28 的 3 条（3.5% NaCl/50℃ 极化曲线只在 p6 Fig.3(a) 图里，"
                 "正文无数值型点蚀电位）、文献 32 的 1 条（原文只给中位数 270–290 mV SCE 与 Fig.7 累积概率分布，"
                 "不对应单一试样）。本轮按「不编造数值」保守留空，未做读图取点。")
    lines.append("- **师兄口径待裁决 5 处**（已逐条写进 `0.xlsx`「备注」列）："
                 "① 文献 34 固溶水淬后还有 900℃×0.5h 时效，属「不要两步热处理」的冲突项；"
                 "② 文献 38 用 3.56 wt% NaCl，而库内文献 27 曾因 3.56% 被整篇删除；"
                 "③ 文献 39 主体是时效系列，只取固溶态 1 条是否算单步，且方法节 pH 3.0 与摘要 3.5% NaCl 矛盾；"
                 "④ 文献 41 的 As received 两组是热轧原始态，归「冷轧+退火」是否合适；"
                 "⑤ 文献 38 的 200 系 Table 1 缺 Si/P、Mo 后三组归属不明，按「未提及」填，未猜数。")
    lines.append("- **跨工作线重复风险**：`准\\_B类可录清单_20261104.csv` 里的 "
                 "`SD_S1350630725010702_Damage_mechanism_of_316L...pdf` 文件名写 316L，"
                 "正文实为 Li 2024 Heliyon 10:e25964（Mo/N 超级奥氏体），**与本轮文献 28 是同一篇**。"
                 "A/B 类 49 条 / 11 篇目前只在 CSV 清单里、还没入 `0.xlsx`，后续合并必须按 DOI/题名去重，"
                 "不能把同一实验算两遍。")
    lines.append("- **清洗版已滞后**：`1.xlsx`（150 条）与 `2.xlsx`（85 条）都基于旧的 157 条基准表，"
                 "没有包含本轮 49 条；`1.xlsx` 的 SCE 换算、0 填充、Fe 实算等口径也不能直接套到新数据上。")
    lines.append("- 复制校验：%d / %d 个文件 MD5 一致。" % (len(copied), len(data)))
    lines.append("")
    lines.append("## 四、建议")
    lines.append("")
    lines.append("1. 先打开 `0.xlsx` 翻到第 160–208 行，重点看「备注」列里标了【口径存疑】【待裁决】的 5 篇，"
                 "请师兄定去留；定完再决定要不要回改基准表。")
    lines.append("2. 文献 28 的 3 条 Epit 可以补：`准\\_B类可录清单_20261104.csv` 已按读图给出 "
                 "0.96 / 0.90 等数值（Fig.3a，误差 ±0.02 V）。补之前先核对该清单的参比口径"
                 "（表头写 `Epit_V_SCE`，而基准表按原文参比照录），并只补 3.5% NaCl/50℃ 这一组；"
                 "文献 32 的统计分布同理，若要用需注明是中位数而非逐试样值。")
    lines.append("3. 重算清洗版：以 `0.xlsx` 的 206 条为新基线重跑 `scripts\\build_1.py` 及后续分组脚本，"
                 "把新数据的 Ag/AgCl 参比换算、`<0.02` 一类上限值、`N/A` 占位统一处理；"
                 "A/B 类 49 条入册前先按 DOI 与题名去重。")
    lines.append("4. `文献\\27.pdf` 继续保留为已剔除 C002 的原件，不占新编号；"
                 "`准\\` 15 个 PDF 原件保持不动，作为后续复核的证据源。")
    lines.append("")
    return "\n".join(lines)


def main():
    data = json.load(open(DATA, encoding="utf-8"))
    if [p["no"] for p in data] != list(range(28, 43)):
        raise SystemExit("编号不是 28-42，停止")

    plan = []
    for p in data:
        src = os.path.join(WS, "准", p["file"])
        dst = os.path.join(WS, "文献", "%02d.pdf" % p["no"])
        if not os.path.isfile(src):
            raise SystemExit("缺源 PDF：" + src)
        if os.path.exists(dst):
            raise SystemExit("目标已存在，拒绝覆盖：" + dst)
        plan.append((p["no"], src, dst))

    print("计划复制 %d 个 PDF：" % len(plan))
    for no, src, dst in plan:
        print("  %02d.pdf  <-  %s  (%d bytes)" % (no, os.path.basename(src), os.path.getsize(src)))

    if "--commit" not in sys.argv:
        print("[dry] 未复制、未写对照清单；确认后加 --commit")
        return

    copied = []
    for no, src, dst in plan:
        shutil.copy2(src, dst)
        if md5(src) != md5(dst):
            raise SystemExit("MD5 不一致，停止：" + dst)
        copied.append(no)
    print("复制完成 %d / %d，MD5 全部一致" % (len(copied), len(plan)))

    with open(MANIFEST, "w", encoding="utf-8") as f:
        f.write("# 准文献入库对照（2026-11-04）\n\n")
        f.write(build_manifest(data, copied))
    print("对照清单：" + MANIFEST)


if __name__ == "__main__":
    main()
