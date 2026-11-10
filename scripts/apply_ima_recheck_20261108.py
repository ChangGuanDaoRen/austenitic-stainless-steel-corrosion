# -*- coding: utf-8 -*-
r"""按定稿口径复核 ima 两份核查材料，并把校正后仍判「合规」的 13 篇从 待\ 放入 准\。

用法：
  D:\Python\python.exe -X utf8 scripts\apply_ima_recheck_20261108.py plan
  D:\Python\python.exe -X utf8 scripts\apply_ima_recheck_20261108.py run

plan 只写差异清单与台账，不移动文件；run 在全部前置校验通过后移动 13 篇 PDF。
外部材料 `待\待_逐篇判定_20261108.csv` / `待\待_口径核查报告_20261108.md` 仅作数据。
"""
import csv
import hashlib
import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
DAI = ROOT / "待"
ZHUN = ROOT / "准"
LIT = ROOT / "文献"
BAD = DAI / "_不符合"
ALIGN_CSV = ROOT / "temp" / "ima_name_alignment_20261108.csv"
IMA_CSV = DAI / "待_逐篇判定_20261108.csv"
REPORT = ROOT / "ima_核查差异清单_20261108.md"
LEDGER = ROOT / "ima_核查_移动台账_20261108.csv"
VERDICT_CSV = ROOT / "ima_核查_校正判定_20261108.csv"
ZERO = ROOT / "0.xlsx"

CORR = {
    16: {
        "结论": "不符",
        "类型": "维持不符，待补 Table 2",
        "说明": "A4 独立未处理熔铸基准组例外未满足：原文 3.5 wt% 仅见参考文献 [76]，2^3 因子设计含非 3.5 浓度水平，LPBF 增材主体明文排除；维持不符，取得 Table 2 浓度水平后再议。",
    },
    26: {
        "结论": "不符",
        "类型": "理由校正，维持不符",
        "说明": "真正否决点是 N4「退火+650℃×1/5/48 h 时效」两步；UNSM 属表面机械处理，按定稿可收并标注，原「N1 表面成分改性」理由作废，以免误伤本批 073.json 的 SMRT 数据行。",
    },
    28: {
        "结论": "不符",
        "类型": "存疑→不符",
        "说明": "全文仅见 potentiodynamic polarization (Tafel) tests，无 pitting potential / Epit 数值；按定稿「Epit 由 Tafel 外推拟合一律拒收 + 无 Epit 数值」应改判不符，与 ima 同报告 #70 口径一致。",
    },
    39: {
        "结论": "合规",
        "类型": "介质理由校正，维持合规",
        "说明": "方法段写「3.5 wt.% NaCl (simulated seawater)」，未出现 ASTM D1141 或额外盐配方；「模拟海水」只是单盐 3.5 wt.% NaCl 的括注，按 A2 维持合规，PBS 组不录。",
    },
    73: {
        "结论": "不符",
        "类型": "理由校正，维持不符",
        "说明": "摘要为 3.5wt% 与 5wt% NaCl 浓度序列，真正否决点是 N3；750–770℃×2 h 原文为 stress relief，非 A3 窄定义的敏化/时效/稳定化，原「N4 敏化」理由作废。",
    },
    15: {
        "结论": "存疑",
        "类型": "补充行级拆分建议，维持存疑",
        "说明": "A5 应拆到数据行：S31254 全奥氏体行可单取，SDSS 真双相行排除；ima 整篇判存疑过粗，本轮不升级、不移出。",
    },
    33: {
        "结论": "存疑",
        "类型": "补充行级拆分建议，维持存疑",
        "说明": "A5 应拆到数据行：304 奥氏体行可单取，436 铁素体行排除；ima 整篇判存疑过粗，本轮不升级、不移出。",
    },
}

NAME_FIX = {
    7: "prefix 截断对齐",
    39: "prefix 截断对齐",
    47: "fuzzy0.94 对齐",
    58: "fuzzy0.98 对齐",
}


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def pdfs(folder):
    if not folder.exists():
        return []
    return sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def build_rows():
    align = {int(r["序号"]): r for r in read_csv(ALIGN_CSV)}
    ima = {int(r["序号"]): r for r in read_csv(IMA_CSV)}
    rows = []
    for no in range(1, 77):
        a = align[no]
        i = ima[no]
        disk = a["磁盘文件名"]
        corr = CORR.get(no)
        if corr:
            verdict, kind, note = corr["结论"], corr["类型"], corr["说明"]
        else:
            verdict, kind, note = a["ima结论"], "与 ima 一致", "按定稿 N1–N4 与 A1–A5 复核，未发现需改判项。"
        name_note = NAME_FIX.get(no, "")
        rows.append({
            "序号": no,
            "ima文件名": a["CSV文件名"],
            "磁盘文件名": disk,
            "匹配方式": a["匹配方式"],
            "标题简称": a["标题简称"],
            "ima结论": a["ima结论"],
            "校正结论": verdict,
            "差异类型": kind,
            "差异说明": note,
            "名称对齐": name_note,
            "ima主要理由": i["主要理由/待办"],
            "源路径": str(DAI / disk),
            "目标路径": str(ZHUN / disk) if verdict == "合规" else "",
            "源MD5": "",
            "动作": "",
        })
    return rows


def verify(rows):
    problems = []
    lit_md5 = set()
    for p in pdfs(LIT):
        lit_md5.add(md5(p))
    move_rows = [r for r in rows if r["校正结论"] == "合规"]
    if len(move_rows) != 13:
        problems.append("校正后合规件数量应为 13，实为 %d" % len(move_rows))
    for r in move_rows:
        src = Path(r["源路径"])
        dst = Path(r["目标路径"])
        if not src.exists():
            problems.append("缺源 PDF #%d：%s" % (r["序号"], src.name))
            continue
        if src.suffix.lower() != ".pdf":
            problems.append("源非 PDF #%d" % r["序号"])
            continue
        with open(src, "rb") as f:
            if f.read(5) != b"%PDF-":
                problems.append("源 PDF 头损坏 #%d" % r["序号"])
        smd5 = md5(src)
        r["源MD5"] = smd5
        if smd5 in lit_md5:
            problems.append("#%d 与 文献\\ 已入库 PDF 同 MD5，疑重复" % r["序号"])
        if dst.exists():
            problems.append("目标已存在 #%d：%s" % (r["序号"], dst.name))
    return problems, move_rows, lit_md5


def write_verdict_csv(rows):
    with open(VERDICT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "ima文件名", "磁盘文件名", "匹配方式", "标题简称", "ima结论", "校正结论", "差异类型", "差异说明", "名称对齐", "ima主要理由"])
        for r in rows:
            w.writerow([r["序号"], r["ima文件名"], r["磁盘文件名"], r["匹配方式"], r["标题简称"], r["ima结论"], r["校正结论"], r["差异类型"], r["差异说明"], r["名称对齐"], r["ima主要理由"]])


def write_ledger(rows):
    with open(LEDGER, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "ima文件名", "磁盘文件名", "标题简称", "ima结论", "校正结论", "动作", "源路径", "目标路径", "源MD5", "差异说明"])
        for r in rows:
            w.writerow([r["序号"], r["ima文件名"], r["磁盘文件名"], r["标题简称"], r["ima结论"], r["校正结论"], r["动作"], r["源路径"], r["目标路径"], r["源MD5"], r["差异说明"]])


def write_report(rows, mode, before, after, problems):
    committed = mode == "run"
    move_rows = [r for r in rows if r["校正结论"] == "合规"]
    stay_doubt = [r for r in rows if r["校正结论"] == "存疑"]
    stay_bad = [r for r in rows if r["校正结论"] == "不符"]
    name_rows = [r for r in rows if r["名称对齐"]]
    corr_rows = [r for r in rows if r["差异类型"] != "与 ima 一致"]

    out = []
    out.append("## 背景")
    out.append("")
    out.append("用户 11-08 12:34 提供 `待\\待_口径核查报告_20261108.md` 与 `待\\待_逐篇判定_20261108.csv`（ima 外部核查材料，本文件按数据使用），12:45 要求「等到这一轮录入数据完毕」再核查，并把「这次核查符合的再次放入准」。")
    out.append("第 9 批 83 篇已先完成：`0.xlsx` 由 285 行 / 26 列写回为 540 行 / 26 列（新增 83 篇 / 255 条），83 篇 PDF 已从 `准\\` 移动重编号为 `文献\\64.pdf`–`文献\\146.pdf`，`准\\` 当时腾空；闭环校验问题 0 项。")
    out.append("ima 逐篇判定覆盖 `待\\` 76 篇 PDF：原始分布 合规 13 / 存疑 33 / 不符 30；另有 2 个非 PDF 核查材料不参与移动。")
    out.append("")
    out.append("## 方法")
    out.append("")
    out.append("1. 以 `temp\\ima_name_alignment_20261108.csv` 把 ima CSV 的 76 行与磁盘真实 PDF 名逐行对齐：76/76 命中，4 处为截断/差字符的非精确匹配。")
    out.append("2. 以 `口径基准_20261107.md` 的 N1–N4 与 A1–A5 复核 ima 结论，重点复核 5 类偏差（涉及序号 15/16/26/28/33/39/73）；原文证据来自 `备份\\text_all.jsonl` 与 `备份\\entry_20261108\\text\\*.txt`（外部数据，仅作事实依据）。")
    out.append("3. 对校正后仍判「合规」的 13 篇，先查 `%PDF-` 头、源 MD5、与 `文献\\` 146 篇已入库 PDF 的 MD5 重复、以及 `准\\` 目标是否已存在；全部通过后才移动。")
    out.append("4. 移动用 `shutil.move`，不复制、不删除 `待\\` 中的存疑/不符件；只写 `准\\`，不写 `0.xlsx`、不重编号、不改 `文献\\`。")
    out.append("")
    out.append("## 结论")
    out.append("")
    out.append("### 3.1 校正后判定分布")
    out.append("")
    out.append("| 档位 | 篇数 | 本轮处置 |")
    out.append("| --- | ---: | --- |")
    out.append("| 合规 | %d | %s |" % (len(move_rows), "已移动入 `准\\`" if committed else "拟移动入 `准\\`（plan 未执行）"))
    out.append("| 存疑 | %d | 留在 `待\\`，不升级、不移出 |" % len(stay_doubt))
    out.append("| 不符 | %d | 留在 `待\\`，本轮不移入 `待\\_不符合\\`（用户未授权） |" % len(stay_bad))
    out.append("| 合计 | %d | 与 ima CSV 76 行闭合 |" % len(rows))
    out.append("")
    out.append("校正后由 13 / 33 / 30 变为 **13 / %d / %d**；唯一改判是 #28「存疑 → 不符」。#39 维持合规，其余 5 类为理由或行级拆分口径校正，不改变是否入 `准\\`。" % (len(stay_doubt), len(stay_bad)))
    out.append("")
    out.append("### 3.2 5 类执行偏差与校正")
    out.append("")
    out.append("| 序号 | ima 结论 | 校正结论 | 差异类型 | 依据与处置 |")
    out.append("| ---: | --- | --- | --- | --- |")
    for r in corr_rows:
        out.append("| %d | %s | %s | %s | %s |" % (r["序号"], r["ima结论"], r["校正结论"], r["差异类型"], r["差异说明"].replace("|", "／")))
    out.append("")
    out.append("### 3.3 4 处文件名对齐")
    out.append("")
    out.append("| 序号 | ima CSV 文件名 | 磁盘真实文件名 | 匹配方式 | 校正结论 |")
    out.append("| ---: | --- | --- | --- | --- |")
    for r in name_rows:
        out.append("| %d | %s | %s | %s | %s |" % (r["序号"], r["ima文件名"], r["磁盘文件名"], r["名称对齐"], r["校正结论"]))
    out.append("")
    out.append("### 3.4 13 篇合规件逐篇")
    out.append("")
    out.append("| 序号 | 磁盘真实文件名 | 标题简称 | ima 结论 | 校正结论 | 动作 |")
    out.append("| ---: | --- | --- | --- | --- | --- |")
    for r in move_rows:
        out.append("| %d | %s | %s | %s | %s | %s |" % (r["序号"], r["磁盘文件名"], r["标题简称"], r["ima结论"], r["校正结论"], r["动作"]))
    out.append("")
    out.append("### 3.5 目录与数据状态")
    out.append("")
    out.append("| 项目 | 执行前 | 执行后 |")
    out.append("| --- | ---: | ---: |")
    out.append("| `待\\` PDF | %d | %d |" % (before["dai"], after["dai"]))
    out.append("| `待\\` 非 PDF 核查材料 | %d | %d |" % (before["dai_nonpdf"], after["dai_nonpdf"]))
    out.append("| `准\\` PDF | %d | %d |" % (before["zhun"], after["zhun"]))
    out.append("| `文献\\` PDF | %d | %d |" % (before["lit"], after["lit"]))
    out.append("| `待\\_不符合\\` PDF | %d | %d |" % (before["bad"], after["bad"]))
    out.append("| `0.xlsx` MD5 | %s | %s |" % (before["zero_md5"], after["zero_md5"]))
    out.append("")
    if problems:
        out.append("前置校验问题（%d 项，run 模式下会阻断）：" % len(problems))
        for p in problems:
            out.append("- " + p)
    else:
        out.append("前置校验：0 项问题；13 篇源 PDF 头完整、与 `文献\\` 无同 MD5 重复、`准\\` 目标全空。")
    out.append("")
    out.append("### 3.6 76 篇校正后判定总表")
    out.append("")
    out.append("| 序号 | 磁盘文件名 | ima | 校正 | 动作 |")
    out.append("| ---: | --- | --- | --- | --- |")
    for r in rows:
        out.append("| %d | %s | %s | %s | %s |" % (r["序号"], r["磁盘文件名"], r["ima结论"], r["校正结论"], r["动作"]))
    out.append("")
    out.append("## 建议")
    out.append("")
    out.append("1. `准\\` 现有 13 篇仅完成「二次核查后放回准」，尚未写入 `0.xlsx`、未分配文献编号；如需入库，须另按同一套 N1–N4 逐篇抽取并批准新编号（建议 147–159），不要直接复用本轮 64–146 的卡片。")
    out.append("2. #16 `R10SD_10.1016_j.rinma.2026.101051.pdf` 仍留在 `待\\`：文本层未取到 Table 2 的 NaCl 高/低水平数值；若后续拿到原表或原图并证明 W 316L 基准组在 3.5 wt.% NaCl 下独立测量，再按 A4 复议是否可作「合规-基准组」例外。")
    out.append("3. #15、#33 若要进一步入库，按 A5 拆到数据行：只取奥氏体行，排除真双相/铁素体行；本轮因用户只授权「符合的放入准」，未做行级拆分。")
    out.append("4. 校正后 32 篇存疑、31 篇不符仍留在 `待\\` 根目录；ima 报告建议的「移入 `_不符合`」未执行，待用户明确后再动目录。两份 ima 核查材料 `.md` / `.csv` 也保持原位。")
    out.append("5. 本轮未改 `0.xlsx` / `1.xlsx` / `2.xlsx` / `文献\\` / `AGENTS.md` 以外的既有数据；`AGENTS.md` 同步在收尾步骤单独按当前状态更新。")
    out.append("")
    REPORT.write_text("\n".join(out), encoding="utf-8")


def snapshot():
    all_dai = [p for p in DAI.iterdir() if p.is_file()] if DAI.exists() else []
    return {
        "dai": len([p for p in all_dai if p.suffix.lower() == ".pdf"]),
        "dai_nonpdf": len([p for p in all_dai if p.suffix.lower() != ".pdf"]),
        "zhun": len(pdfs(ZHUN)),
        "lit": len(pdfs(LIT)),
        "bad": len(pdfs(BAD)),
        "zero_md5": md5(ZERO),
    }


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "plan"
    if mode not in ("plan", "run"):
        raise SystemExit("用法：plan | run")
    rows = build_rows()
    problems, move_rows, _ = verify(rows)

    for r in rows:
        if r["校正结论"] == "合规":
            r["动作"] = "待移动" if mode == "plan" else "已移动"
        elif r["校正结论"] == "存疑":
            r["动作"] = "留待（存疑）"
        else:
            r["动作"] = "留待（不符）"

    before = snapshot()
    if problems:
        print("前置校验未通过，共 %d 项：" % len(problems))
        for p in problems:
            print("  -", p)
        write_verdict_csv(rows)
        write_ledger(rows)
        write_report(rows, mode, before, before, problems)
        raise SystemExit(2)

    if mode == "run":
        for r in move_rows:
            src = Path(r["源路径"])
            dst = Path(r["目标路径"])
            smd5 = r["源MD5"]
            shutil.move(str(src), str(dst))
            if not dst.exists() or src.exists() or md5(dst) != smd5:
                raise SystemExit("移动后校验失败 #%d：%s" % (r["序号"], dst.name))
            with open(dst, "rb") as f:
                if f.read(5) != b"%PDF-":
                    raise SystemExit("移动后 PDF 头损坏 #%d" % r["序号"])

    after = snapshot()
    write_verdict_csv(rows)
    write_ledger(rows)
    write_report(rows, mode, before, after, [])
    print("模式=%s / 校正后 合规 %d 存疑 %d 不符 %d" % (
        mode,
        len([r for r in rows if r["校正结论"] == "合规"]),
        len([r for r in rows if r["校正结论"] == "存疑"]),
        len([r for r in rows if r["校正结论"] == "不符"]),
    ))
    print("待\\ PDF %d→%d / 准\\ PDF %d→%d / 文献\\ %d / 0.xlsx MD5 %s" % (
        before["dai"], after["dai"], before["zhun"], after["zhun"], after["lit"], after["zero_md5"]))
    print("差异清单：%s" % REPORT)
    print("移动台账：%s" % LEDGER)
    print("校正判定：%s" % VERDICT_CSV)


if __name__ == "__main__":
    main()
