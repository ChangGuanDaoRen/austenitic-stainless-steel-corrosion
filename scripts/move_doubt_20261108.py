#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""按用户 2026-11-08 16:35 指令：把校正结论=存疑的 32 篇从 待\ 移动进 准\；不符 31 篇原位不动。

用法：
  D:\Python\python.exe -X utf8 scripts\move_doubt_20261108.py dry
  D:\Python\python.exe -X utf8 scripts\move_doubt_20261108.py run

只读不碰 0.xlsx / 1.xlsx / 文献\ / 待\ 的非 PDF 与不符件。
run 用 shutil.move（用户已授权移动），逐篇校验 %PDF- 头与 MD5，写台账与报告。
"""
import csv
import hashlib
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAI = os.path.join(ROOT, "待")
ZHUN = os.path.join(ROOT, "准")
VERDICT = os.path.join(ROOT, "ima_核查_校正判定_20261108.csv")
LEDGER = os.path.join(ROOT, "存疑移入准_移动台账_20261108.csv")
REPORT = os.path.join(ROOT, "存疑移入准_报告_20261108.md")
AGENTS = os.path.join(ROOT, "AGENTS.md")
ZERO = os.path.join(ROOT, "0.xlsx")


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1024 * 1024), b""):
            h.update(b)
    return h.hexdigest()


def pdf_head(path):
    with open(path, "rb") as f:
        return f.read(5) == b"%PDF-"


def pdfs(folder):
    return sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))


def read_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def build_rows():
    rows = read_csv(VERDICT)
    doubt = [r for r in rows if r["校正结论"] == "存疑"]
    bad = [r for r in rows if r["校正结论"] == "不符"]
    good = [r for r in rows if r["校正结论"] == "合规"]
    if len(doubt) != 32 or len(bad) != 31 or len(good) != 13 or len(rows) != 76:
        raise SystemExit("判定表分布异常：存疑 %d / 不符 %d / 合规 %d / 总 %d" % (len(doubt), len(bad), len(good), len(rows)))
    return rows, doubt, bad, good


def verify(doubt, bad):
    problems = []
    zhun_names = set(pdfs(ZHUN))
    for r in doubt:
        name = r["磁盘文件名"]
        src = os.path.join(DAI, name)
        if not os.path.isfile(src):
            problems.append("存疑件源缺失：" + name)
            continue
        if not pdf_head(src):
            problems.append("存疑件非 PDF：" + name)
        if name in zhun_names:
            problems.append("准\\ 已有同名：" + name)
    for r in bad:
        name = r["磁盘文件名"]
        src = os.path.join(DAI, name)
        if not os.path.isfile(src):
            problems.append("不符件不在 待\\（按指令应原位不动）：" + name)
    before_pdf = pdfs(DAI)
    uncovered = [n for n in before_pdf if n not in {r["磁盘文件名"] for r in doubt + bad}]
    if uncovered:
        problems.append("待\\ 有未被判定覆盖的 PDF：" + ", ".join(uncovered))
    return problems


def snapshot():
    return {
        "dai_pdf": len(pdfs(DAI)),
        "dai_other": len([f for f in os.listdir(DAI) if not f.lower().endswith(".pdf")]),
        "zhun_pdf": len(pdfs(ZHUN)),
        "zero_md5": md5(ZERO),
    }


def write_ledger(doubt, bad, mode, problems):
    with open(LEDGER, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "磁盘文件名", "校正结论", "动作", "源MD5", "目标路径", "移动后校验", "主要理由/待办"])
        for r in doubt:
            name = r["磁盘文件名"]
            src = os.path.join(DAI, name)
            dst = os.path.join(ZHUN, name)
            srcmd5 = md5(src) if os.path.isfile(src) else ""
            if mode == "run" and not problems:
                dstmd5 = md5(dst) if os.path.isfile(dst) else ""
                ok = dstmd5 == srcmd5 and not os.path.exists(src) and pdf_head(dst)
                w.writerow([r["序号"], name, "存疑", "已移动", srcmd5, "准\\" + name, "MD5一致/源已消失/%PDF-好" if ok else "校验失败", r["ima主要理由"]])
            else:
                w.writerow([r["序号"], name, "存疑", "[dry] 待移动" if mode == "dry" else "阻断未移动", srcmd5, "准\\" + name, "", r["ima主要理由"]])
        for r in bad:
            name = r["磁盘文件名"]
            w.writerow([r["序号"], name, "不符", "原位不动", md5(os.path.join(DAI, name)) if os.path.isfile(os.path.join(DAI, name)) else "", "待\\" + name, "", r["ima主要理由"]])


def write_report(doubt, bad, mode, before, after, problems):
    moved = 0 if mode != "run" or problems else sum(1 for r in doubt if os.path.isfile(os.path.join(ZHUN, r["磁盘文件名"])) and not os.path.exists(os.path.join(DAI, r["磁盘文件名"])))
    lines = [
        "# 存疑件移入准、不符件原位不动（2026-11-08）",
        "",
        "时间：2026-11-08 16:3x；用户 16:35 原话「先把存疑的移动到准，不符的留待不动。」",
        "",
        "## 一、背景",
        "",
        "第 10 批 13 篇合规件已录入 `0.xlsx` 并移动重编号 `文献\\147.pdf`–`159.pdf`。`待\\` 余 63 篇 PDF 经 ima 二次核查校正为 32 存疑 + 31 不符，另有 2 个非 PDF 核查材料。按用户指令把 32 存疑件移入 `准\\`，31 不符件原位不动；本次只动目录，不写 `0.xlsx`、不重编号。",
        "",
        "## 二、方法",
        "",
        "- 判定源 `ima_核查_校正判定_20261108.csv`：只取「校正结论＝存疑」行，按「磁盘文件名」定位 `待\\` 源件。",
        "- 前置校验：32 存疑件全在 `待\\`、`%PDF-` 头完整、`准\\` 无同名；31 不符件全在 `待\\`；`待\\` 63 篇 PDF 恰被 32+31 覆盖；`0.xlsx` MD5 前后不变。",
        "- `dry` 只预演；`run` 用 `shutil.move` 逐篇移动并校验 `md5(目标)==源MD5`、`源已消失`、`%PDF-` 头；台账 `存疑移入准_移动台账_20261108.csv` 记 63 行（32 已移动 + 31 原位不动）。",
        "- 同步 `AGENTS.md` 第二节目录树当前真值（只改本节）；原 `需人工裁决文献报告_20261108.md` 为移动前状态快照，保留并在本报告标注。",
        "",
        "## 三、结论",
        "",
        "| 项 | 数值 |",
        "| --- | --- |",
        "| 模式 | %s |" % mode,
        "| 判定表分布 | 合规 13（第10批已入库）/ 存疑 32 / 不符 31 / 合计 76 |",
        "| 存疑移动 | %d / 32 |" % moved,
        "| 不符原位不动 | 31 / 31 |",
        "| `待\\` PDF | %d → %d（另 2 个非 PDF 材料保持原位） |" % (before["dai_pdf"], after["dai_pdf"]),
        "| `准\\` PDF | %d → %d |" % (before["zhun_pdf"], after["zhun_pdf"]),
        "| `0.xlsx` MD5 | %s → %s（未改） |" % (before["zero_md5"], after["zero_md5"]),
        "| 前置/后置问题 | %d 项 |" % len(problems),
        "",
        "逐篇明细（序号 / 文件名 / 判定 / 动作 / 源 MD5 / 目标 / 校验 / 理由）见 `存疑移入准_移动台账_20261108.csv`。",
        "",
        "## 四、建议",
        "",
        "1. `准\\` 32 篇是二次核查后的存疑件，**尚未录入 `0.xlsx`**；是否入库、是否补证后另起一批从 160 起编号，等用户新指令，不自行扩量。",
        "2. `待\\` 31 篇不符件与 2 个非 PDF 核查材料原位不动；若后续要移入 `_不符合\\`，须先确认该目录当前不存在、是否新建与移动范围。",
        "3. 回滚：按台账把 `准\\` 中 32 篇移回 `待\\` 即可；台账有源 MD5，`0.xlsx` 未参与本动作。",
        "4. `需人工裁决文献报告_20261108.md` 的「在位核查」为移动前快照（32 存疑曾在 `待\\`）；移动后以本报告与台账为准。",
    ]
    if problems:
        lines.insert(len(lines) - 4, "")
        lines.insert(len(lines) - 4, "### 阻断项")
        lines.insert(len(lines) - 4, "")
        for p in problems:
            lines.insert(len(lines) - 4, "- " + p)
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def sync_agents(after):
    with open(AGENTS, "r", encoding="utf-8") as f:
        text = f.read()
    old = "├─ 准\\                        **0 篇 PDF**；第9批 83 篇与第10批 13 篇均已入 0 并移动重编号"
    new = "├─ 准\\                        **%d 篇 PDF**；由 `待\\` 32 篇校正结论「存疑」移入（2026-11-08 16:35 指令），未入 0、不编号；第9批 83 篇与第10批 13 篇已入 0 并移动重编号" % after["zhun_pdf"]
    if old not in text and "由 `待\\` 32 篇校正结论「存疑」移入" in text:
        print("AGENTS.md 已同步，跳过目录树重写")
        return
    if old not in text:
        raise SystemExit("AGENTS.md 未找到准\\ 定位串，停止同步")
    text = text.replace(old, new, 1)
    old2 = "├─ 待\\                        **63 篇 PDF + 2 个非 PDF 核查材料**（`待_口径核查报告_20261108.md`、`待_逐篇判定_20261108.csv`）"
    new2 = "├─ 待\\                        **%d 篇 PDF（31 篇校正结论「不符」原位不动）+ 2 个非 PDF 核查材料**（`待_口径核查报告_20261108.md`、`待_逐篇判定_20261108.csv`）" % after["dai_pdf"]
    if old2 not in text:
        raise SystemExit("AGENTS.md 未找到待\\ 定位串，停止同步")
    text = text.replace(old2, new2, 1)
    old3 = "├─ 需人工裁决文献报告_20261108.md 待\\ 63 篇需人工裁决文献报告"
    new3 = "├─ 需人工裁决文献报告_20261108.md 待\\ 63 篇需人工裁决文献报告（移动前快照）\n├─ 存疑移入准_报告_20261108.md 32 存疑移入准、31 不符原位不动\n├─ 存疑移入准_移动台账_20261108.csv 63 行逐篇动作与源 MD5"
    if old3 not in text:
        raise SystemExit("AGENTS.md 未找到报告定位串，停止同步")
    text = text.replace(old3, new3, 1)
    old4 = "- **当前待用户看**：`需人工裁决文献报告_20261108.md` 中 63 篇待裁决件；优先 6 条争议件。未获明确指令前，不升级、不移出、不写入 `0.xlsx`。"
    new4 = "- **当前待用户看**：`需人工裁决文献报告_20261108.md` 中 63 篇待裁决件（移动前快照）；优先 6 条争议件。32 篇存疑件已按 16:35 指令移入 `准\\`、31 篇不符件留 `待\\`，见 `存疑移入准_报告_20261108.md`。未获明确指令前，`准\\` 32 篇不写入 `0.xlsx`、不编号；`待\\` 31 篇不移出。"
    if old4 not in text:
        raise SystemExit("AGENTS.md 未找到待办定位串，停止同步")
    text = text.replace(old4, new4, 1)
    if len(text.encode("utf-8")) > 16384:
        raise SystemExit("AGENTS.md 超过 16 KB，停止同步")
    with open(AGENTS, "w", encoding="utf-8") as f:
        f.write(text)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    if mode not in ("dry", "run"):
        raise SystemExit("用法：move_doubt_20261108.py dry|run")
    rows, doubt, bad, good = build_rows()
    problems = verify(doubt, bad)
    print("模式=%s / 存疑 %d / 不符 %d / 合规 %d（已入库）" % (mode, len(doubt), len(bad), len(good)))
    if problems:
        print("前置阻断 %d 项：" % len(problems))
        for p in problems:
            print("  ", p)
    before = snapshot()
    print("移动前：待\\ PDF %d / 准\\ PDF %d / 0.xlsx MD5 %s" % (before["dai_pdf"], before["zhun_pdf"], before["zero_md5"]))
    if mode == "run" and not problems:
        for r in doubt:
            name = r["磁盘文件名"]
            src = os.path.join(DAI, name)
            dst = os.path.join(ZHUN, name)
            srcmd5 = md5(src)
            shutil.move(src, dst)
            if not os.path.isfile(dst) or os.path.exists(src) or md5(dst) != srcmd5 or not pdf_head(dst):
                raise SystemExit("移动后校验失败：" + name)
        print("已移动 %d 篇到 准\\" % len(doubt))
    after = snapshot()
    print("移动后：待\\ PDF %d / 待\\ 非PDF %d / 准\\ PDF %d / 0.xlsx MD5 %s" % (after["dai_pdf"], after["dai_other"], after["zhun_pdf"], after["zero_md5"]))
    if after["zero_md5"] != before["zero_md5"]:
        raise SystemExit("0.xlsx MD5 变化，停止")
    write_ledger(doubt, bad, mode, problems)
    write_report(doubt, bad, mode, before, after, problems)
    if mode == "run" and not problems:
        sync_agents(after)
        print("已同步 AGENTS.md")
    print("已写台账与报告")
    if problems:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
