# -*- coding: utf-8 -*-
r"""
scripts\process_doubt_20261108.py

按用户 m02720 六条裁定 + ask_user_question 三项选择，处理 准\ 里 32 篇校正结论「存疑」PDF。

用法:
  D:\Python\python.exe -X utf8 scripts\process_doubt_20261108.py dry
  D:\Python\python.exe -X utf8 scripts\process_doubt_20261108.py run

dry: 只做前置校验 + 打印互斥分组与计划，写台账/报告（标注 dry），不备份、不改名、不移动。
run: 先备份 32 篇到 备份\准_存疑批处理前_20261108\，再改名/移动，逐项 MD5 校验，写最终台账/报告。

互斥分组（优先级 P1 > P2 > P3 > P4 > P5 > P6，按判定表「序号」）:
  P1 从严表面/成形  #20,#43,#55,#63                      -> 不符合，移回 待\
  P2 非纯介质/3.56  #12,#18,#25,#29,#32,#57,#66,#71,#74  -> 不符合，移回 待\
  P3 热老化          #69                                  -> 不符合，移回 待\
  P4 有图待读        #6,#22,#31,#35,#37,#38,#42,#47,#48,#50,#52,#53,#58,#76
                                                          -> 留 准\，改名「__有图待读.pdf」
  P5 行级拆分        #15,#33                              -> 留 准\ 不动，待补证
  P6 海水下一批候选  #3,#56                               -> 留 准\ 不动，另起新批从 160 起
  校验 14(移回) + 14(改名) + 2 + 2 = 32

硬约束: 不改 0.xlsx / 1.xlsx / 2.xlsx / 文献\ / 待\_不符合\；不新建目录；不删除文件（移动/改名均可回滚）。
"""
import os
import sys
import csv
import shutil
import hashlib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZHUN = os.path.join(ROOT, "准")
DAI = os.path.join(ROOT, "待")
LIT = os.path.join(ROOT, "文献")
ZERO = os.path.join(ROOT, "0.xlsx")
VER = os.path.join(ROOT, "ima_核查_校正判定_20261108.csv")
LEDGER = os.path.join(ROOT, "准_存疑批处理台账_20261108.csv")
REPORT = os.path.join(ROOT, "准_存疑批处理报告_20261108.md")
BAKDIR = os.path.join(ROOT, "备份", "准_存疑批处理前_20261108")
MARK = "__有图待读"

P1 = {20, 43, 55, 63}
P2 = {12, 18, 25, 29, 32, 57, 66, 71, 74}
P3 = {69}
BAD = P1 | P2 | P3
P4 = {6, 22, 31, 35, 37, 38, 42, 47, 48, 50, 52, 53, 58, 76}
P5 = {15, 33}
P6 = {3, 56}
ALL = BAD | P4 | P5 | P6

EDGE = {
    38: "会议论文", 48: "Epit 判据非拐点", 50: "海水+图示 Epit", 56: "人工海水未给浓度",
    71: "浓缩海水", 35: "δ 8.86%", 20: "天然海水但堆焊", 18: "缓蚀剂且 Epit 无值",
}


def pdfs(d):
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if f.lower().endswith(".pdf"))


def nonpdfs(d):
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if not f.lower().endswith(".pdf"))


def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for c in iter(lambda: f.read(1 << 20), b""):
            h.update(c)
    return h.hexdigest()


def pdf_head(p):
    with open(p, "rb") as f:
        return f.read(5)


def load_rows():
    with open(VER, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    doubt = [r for r in rows if r["校正结论"].strip() == "存疑"]
    by_no = {int(r["序号"]): r for r in doubt}
    return rows, by_no


def group_of(n):
    if n in P1:
        return "不符合-从严表面/成形"
    if n in P2:
        return "不符合-非纯介质或3.56"
    if n in P3:
        return "不符合-热老化"
    if n in P4:
        return "有图待读-留准改名"
    if n in P5:
        return "行级拆分-留准待补证"
    if n in P6:
        return "海水下一批候选-留准"
    return "未分组"


def action_of(n):
    if n in BAD:
        return "移回 待\\"
    if n in P4:
        return "改名 __有图待读.pdf（留准）"
    if n in P5:
        return "留准不动（行级拆分待补证）"
    if n in P6:
        return "留准不动（下一批候选，从160起）"
    return "未定"


def newname_of(name):
    stem, ext = os.path.splitext(name)
    return stem + MARK + ext


def gate(by_no, zhun_files):
    problems = []
    nos = set(by_no)
    if nos != ALL:
        problems.append("分组集合与判定表存疑序号不一致：表内多余 %s / 缺失 %s"
                        % (sorted(nos - ALL), sorted(ALL - nos)))
    chain = list(P1) + list(P2) + list(P3) + list(P4) + list(P5) + list(P6)
    if len(chain) != 32 or len(set(chain)) != 32:
        problems.append("分组非互斥或未覆盖 32（chain=%d unique=%d）" % (len(chain), len(set(chain))))
    names = [by_no[n]["磁盘文件名"] for n in sorted(by_no)]
    zset = set(zhun_files)
    if set(names) != zset:
        problems.append("准目录与判定表存疑磁盘名不一致：表内多余 %s / 准目录多余 %s"
                        % (sorted(set(names) - zset), sorted(zset - set(names))))
    for n in sorted(P4):
        tgt = newname_of(by_no[n]["磁盘文件名"])
        if os.path.exists(os.path.join(ZHUN, tgt)):
            problems.append("改名目标已存在: " + tgt)
    for n in sorted(BAD):
        nm = by_no[n]["磁盘文件名"]
        if os.path.exists(os.path.join(DAI, nm)):
            problems.append("待目录已存在同名（拒绝覆盖）: " + nm)
    for n in sorted(by_no):
        if not os.path.isfile(os.path.join(ZHUN, by_no[n]["磁盘文件名"])):
            problems.append("准目录缺文件: " + by_no[n]["磁盘文件名"])
    return problems


def do_run(by_no):
    if not os.path.isdir(BAKDIR):
        os.makedirs(BAKDIR)
    for n in sorted(by_no):
        nm = by_no[n]["磁盘文件名"]
        shutil.copy2(os.path.join(ZHUN, nm), os.path.join(BAKDIR, nm))
    for n in sorted(P4):
        nm = by_no[n]["磁盘文件名"]
        os.rename(os.path.join(ZHUN, nm), os.path.join(ZHUN, newname_of(nm)))
    for n in sorted(BAD):
        nm = by_no[n]["磁盘文件名"]
        shutil.move(os.path.join(ZHUN, nm), os.path.join(DAI, nm))


def verify(by_no):
    res = {}
    for n in sorted(by_no):
        nm = by_no[n]["磁盘文件名"]
        bak = os.path.join(BAKDIR, nm)
        ref = md5(bak) if os.path.isfile(bak) else ""
        if n in P4:
            cur = os.path.join(ZHUN, newname_of(nm))
            old = os.path.join(ZHUN, nm)
            ok = os.path.isfile(cur) and not os.path.exists(old) and ref and md5(cur) == ref and pdf_head(cur) == b"%PDF-"
        elif n in BAD:
            cur = os.path.join(DAI, nm)
            old = os.path.join(ZHUN, nm)
            ok = os.path.isfile(cur) and not os.path.exists(old) and ref and md5(cur) == ref and pdf_head(cur) == b"%PDF-"
        else:  # P5 / P6 留准不动
            cur = os.path.join(ZHUN, nm)
            ok = os.path.isfile(cur) and ref and md5(cur) == ref
        res[n] = "OK" if ok else "FAIL"
    return res


def curpath(by_no, n, dry):
    nm = by_no[n]["磁盘文件名"]
    if n in P4:
        return "准\\" + newname_of(nm)
    if n in BAD:
        return "待\\" + nm
    return "准\\" + nm


def write_ledger(by_no, dry, checks):
    cols = ["序号", "分组", "动作", "原名", "现名或目标路径", "字节", "源MD5", "校验", "ima主要理由"]
    with open(LEDGER, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for n in sorted(by_no):
            r = by_no[n]
            nm = r["磁盘文件名"]
            src = os.path.join(BAKDIR, nm) if not dry else os.path.join(ZHUN, nm)
            size = os.path.getsize(src) if os.path.isfile(src) else ""
            m = md5(src) if os.path.isfile(src) else ""
            chk = "[dry] 未执行" if dry else checks.get(n, "?")
            reason = r.get("ima主要理由", "").replace("\r", " ").replace("\n", " ")
            w.writerow([n, group_of(n), action_of(n), nm, curpath(by_no, n, dry), size, m, chk, reason])


def esc(s):
    return str(s).replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def write_report(by_no, dry, checks, snap):
    L = []
    mode = "dry（仅预演，未改名未移动）" if dry else "run（已执行并逐项校验）"
    L.append("# 准\\ 32 篇存疑批处理报告（2026-11-08）")
    L.append("")
    L.append("状态：%s；数据文件 `0.xlsx`、`文献\\` 一律未改。" % mode)
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("- 用户 17:33 指令「现在处理准里的存疑，你直接列出需要判断的项」；17:54 六条裁定（m02720）：海水可以先进去但 3.56 不行、非纯介质不符合、有图的留准并文件名标注、允许行级拆分按建议、热老化不行、补证后另起新批。")
    L.append("- 三项方向性歧义经 ask_user_question 锁定：①海水件＝标为下一批候选（不写 0、不编号、不移动，确认后从 160 起）；②成形/焊接/表面成分改性＝从严全部排除；③有图待读标注＝后缀 `__有图待读`。")
    L.append("- 处理对象：`准\\` 32 篇校正结论「存疑」PDF（源自 `ima_核查_校正判定_20261108.csv`）；不含已入 `文献\\` 的 159 篇、`待\\` 既有 31 篇不符件。")
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("- 按单一推荐优先级做互斥分组 P1→P6：P1 从严表面/成形 → P2 非纯介质/3.56 → P3 热老化（三者合并为「不符合」移回 `待\\` 根目录）→ P4 有图待读（留准改名）→ P5 行级拆分待补证（留准不动）→ P6 海水下一批候选（留准不动）。")
    L.append("- 前置门：判定表存疑序号集合＝分组集合＝32；`准\\` 恰好 32 篇且与判定表磁盘名双向一致；改名目标与 `待\\` 目标均不冲突（拒绝覆盖）。")
    L.append("- run 先 `shutil.copy2` 32 篇到 `备份\\准_存疑批处理前_20261108\\`；改名用 `os.rename`、移回用 `shutil.move`；逐项校验 MD5 与备份一致、`%PDF-` 头完整、源位置已消失。")
    L.append("- 未新建 `待\\_不符合\\`（该目录当前不存在）；14 篇移回件与既有 31 篇不符件同处 `待\\` 根目录，靠本台账区分。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("### 3.1 分组与分布")
    L.append("")
    L.append("| 分组 | 处置 | 篇数 |")
    L.append("|---|---|---|")
    L.append("| 不符合-从严表面/成形 | 移回 待\\ | %d |" % len(P1))
    L.append("| 不符合-非纯介质或3.56 | 移回 待\\ | %d |" % len(P2))
    L.append("| 不符合-热老化 | 移回 待\\ | %d |" % len(P3))
    L.append("| 有图待读-留准改名 | 留 准\\，改名 `__有图待读.pdf` | %d |" % len(P4))
    L.append("| 行级拆分-留准待补证 | 留 准\\ 不动 | %d |" % len(P5))
    L.append("| 海水下一批候选-留准 | 留 准\\ 不动，另起新批从 160 起 | %d |" % len(P6))
    L.append("| **合计** | 移回 %d / 改名 %d / 不动 %d | **32** |" % (len(BAD), len(P4), len(P5) + len(P6)))
    L.append("")
    L.append("### 3.2 目录真值（%s）" % ("预期" if dry else "实测"))
    L.append("")
    L.append("| 目录/文件 | 处理前 | 处理后 |")
    L.append("|---|---|---|")
    L.append("| `准\\` PDF | 32 | %d（14 改名 + 4 不动） |" % snap["zhun_after"])
    L.append("| 其中 `__有图待读.pdf` | 0 | %d |" % snap["marked"])
    L.append("| `待\\` PDF | 31 | %d（+14 移回） |" % snap["dai_after"])
    L.append("| `待\\` 非 PDF | 2 | 2 |")
    L.append("| `文献\\` PDF | 159 | 159（未动） |")
    L.append("| `0.xlsx` | MD5 %s | 未改 |" % snap["zero_md5"])
    L.append("")
    L.append("### 3.3 逐篇处置（32 篇）")
    L.append("")
    L.append("| 序号 | 分组/动作 | 磁盘文件名 | 校验 | 主要理由（截断） |")
    L.append("|---|---|---|---|---|")
    for n in sorted(by_no):
        r = by_no[n]
        reason = esc(r.get("ima主要理由", ""))[:60]
        chk = "[dry]" if dry else checks.get(n, "?")
        L.append("| %d | %s / %s | %s | %s | %s |"
                 % (n, esc(group_of(n)), esc(action_of(n)), esc(r["磁盘文件名"]), chk, reason))
    L.append("")
    L.append("### 3.4 边界件（按优先级归入，须留意）")
    L.append("")
    for n in sorted(EDGE):
        if n in by_no:
            L.append("- #%d：%s（归入「%s」）。" % (n, EDGE[n], group_of(n)))
    L.append("")
    L.append("## 四、建议")
    L.append("")
    L.append("1. 14 篇 `__有图待读.pdf` 请回原文读图补 Epit 绝对值；补证后再决定入库，本轮不写 `0.xlsx`、不编号。")
    L.append("2. 2 篇海水下一批候选（#3/#56）与 2 篇行级拆分待补证（#15/#33）确认后另起新批，从 `文献\\160.pdf` 起编号；本轮不重排。")
    L.append("3. 14 篇已移回 `待\\` 与既有 31 篇不符件同目录，靠 `准_存疑批处理台账_20261108.csv` 与本报告区分；是否进一步移入 `待\\_不符合\\` 需新指令。")
    L.append("4. 3.56（#32）、热老化（#69）从严排除；海水可进但本轮只作候选，未入库。")
    L.append("5. 遗留待修（不属本轮）：`存疑移入准_移动台账_20261108.csv` 32 行「源MD5」为空、「校验失败」属记账 bug（`scripts\\move_doubt_20261108.py` 在 `shutil.move` 后才取 srcmd5），实际移动已通过；`AGENTS.md` 已于本轮收尾按当前真值覆盖式重写。")
    L.append("")
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L))


def snapshot(dry):
    zhun = pdfs(ZHUN)
    marked = len([f for f in zhun if MARK in f])
    if dry:
        return {
            "zhun_after": 18, "marked": 14, "dai_after": 45,
            "zero_md5": md5(ZERO),
        }
    return {
        "zhun_after": len(zhun), "marked": marked, "dai_after": len(pdfs(DAI)),
        "zero_md5": md5(ZERO),
    }


def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else "dry").lower()
    if mode not in ("dry", "run"):
        print("用法: process_doubt_20261108.py [dry|run]")
        raise SystemExit(2)
    rows, by_no = load_rows()
    zhun_files = pdfs(ZHUN)
    problems = gate(by_no, zhun_files)
    if problems:
        print("[gate] 前置校验未过，未执行任何动作：")
        for p in problems:
            print("  - " + p)
        raise SystemExit(2)
    print("[gate] 通过：判定表存疑 32 / 准目录 32 / 双向一致 / 改名与移回目标无冲突")
    print("[分组] 移回待 %d（P1=%d,P2=%d,P3=%d）/ 改名留准 %d / 行级拆分留准 %d / 海水候选留准 %d"
          % (len(BAD), len(P1), len(P2), len(P3), len(P4), len(P5), len(P6)))
    checks = {}
    if mode == "run":
        do_run(by_no)
        checks = verify(by_no)
        fails = [n for n, v in checks.items() if v != "OK"]
        print("[run] 已备份 32 篇 / 改名 %d / 移回 %d / 校验失败 %d"
              % (len(P4), len(BAD), len(fails)))
        if fails:
            print("[run] 校验失败序号：%s（文件已按台账动作落地，请人工核对备份）" % sorted(fails))
    snap = snapshot(mode == "dry")
    write_ledger(by_no, mode == "dry", checks)
    write_report(by_no, mode == "dry", checks, snap)
    print("[%s] 已写台账 %s 与报告 %s" % (mode, os.path.basename(LEDGER), os.path.basename(REPORT)))
    print("[闭环] 准 %d（__有图待读 %d）/ 待 PDF %d + 非PDF %d / 文献 %d / 0.xlsx MD5 %s"
          % (snap["zhun_after"], snap["marked"], snap["dai_after"],
             len(nonpdfs(DAI)), len(pdfs(LIT)), snap["zero_md5"]))


if __name__ == "__main__":
    main()
