# -*- coding: utf-8 -*-
r"""第 9 批入库收尾编排：卡片收齐后一条命令跑完下游确定性链路。

用法：`D:\Python\python.exe -X utf8 scripts\finalize_entry9_20261108.py`
先做只读前置检查（83 张卡片齐、无 flag、源 PDF 全好、目标文献全空），
任一不满足即退出且不改任何数据文件；通过后才按序执行：
  preflight → build check → build → commit → 移动重编号 → 校验报告 → 闭环校验
每步用同一个 Python 解释器、不捕获子进程输出、失败即停（check=True）。

本脚本会改 `0.xlsx`、`文献\` 与 `准\`（2026-11-08 10:32 批准入库、12:45 批准改移动腾空），
不改 `1.xlsx` / `2.xlsx` / `待\` / `AGENTS.md`；`AGENTS.md` 同步单独手工 edit。
"""
import hashlib
import json
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
CARD_DIR = os.path.join(WORK, "cards")
SRC_DIR = os.path.join(ROOT, "准")
LIT_DIR = os.path.join(ROOT, "文献")
PY = sys.executable

STEPS = [
    ("前置只读体检", "scripts\\preflight_entry9_20261108.py", []),
    ("语义合规审计（基准组/换算/长备注/日志措辞）", "scripts\\audit_cards_semantics_20261108.py", []),
    ("卡片质检 + 生成 final_rows.json", "scripts\\build_entry9_20261108.py", ["check"]),
    ("生成校验件并逐格校验（不改 0.xlsx）", "scripts\\build_entry9_20261108.py", ["build"]),
    ("写回 0.xlsx 并二次复验", "scripts\\build_entry9_20261108.py", ["commit"]),
    ("移动重编号 文献\\64.pdf–146.pdf（准\\ 腾空）+ 出对照报告", "scripts\\copy_lit_and_report_20261108.py", ["run"]),
    ("出 0.xlsx 第9批入库校验报告", "scripts\\make_entry9_verify_report_20261108.py", []),
    ("最终闭环校验（0.xlsx ↔ 文献 ↔ 台账 MD5）", "scripts\\verify_entry9_final_20261108.py", []),
]


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def gate():
    """卡片全齐、无 flag、JSON 可解析、源 PDF 全好、目标文献全空、0.xlsx 未变。"""
    with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    problems = []
    base = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
    target = os.path.join(ROOT, "0.xlsx")
    if not os.path.exists(base) or not os.path.exists(target):
        problems.append("0.xlsx 或第9批入库前底本缺失")
    elif md5(base) != md5(target):
        problems.append("0.xlsx 已与第9批入库前底本不一致（build/commit 会阻断）")
    for m in manifest:
        no = m["newno"]
        card = os.path.join(CARD_DIR, "%03d.json" % no)
        if not os.path.exists(card):
            problems.append("缺卡片 %03d" % no)
            continue
        try:
            with open(card, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            problems.append("卡片 %03d JSON 错：%s" % (no, e))
            continue
        if str(c.get("flag") or "").strip():
            problems.append("卡片 %03d 带 flag：%s" % (no, c["flag"][:60]))
        rows = c.get("rows")
        if not isinstance(rows, list) or len(rows) < 1:
            problems.append("卡片 %03d 无数据行" % no)
        src = os.path.join(SRC_DIR, m["file"])
        if not os.path.exists(src):
            problems.append("缺源 PDF %03d：%s" % (no, m["file"]))
        else:
            with open(src, "rb") as f:
                if f.read(5) != b"%PDF-":
                    problems.append("源非 PDF %03d" % no)
    for no in range(64, 147):
        dst = os.path.join(LIT_DIR, "%02d.pdf" % no)
        if os.path.exists(dst):
            problems.append("目标已存在 文献\\%02d.pdf（需人工核对是否已移动/复制过）" % no)
    return problems


def main():
    print("=" * 72)
    print("第 9 批入库收尾编排 / 工作目录：" + ROOT)
    print("=" * 72)
    problems = gate()
    if problems:
        print("\n前置检查未通过，共 %d 项；未改任何文件：\n" % len(problems))
        for x in problems[:40]:
            print("  -", x)
        if len(problems) > 40:
            print("  ... 其余 %d 项省略" % (len(problems) - 40))
        raise SystemExit(2)
    print("前置检查通过：83 张卡片齐全 / 无 flag / 源 PDF 全好 / 64–146 目标全空\n")

    for i, (title, script, args) in enumerate(STEPS, 1):
        print("-" * 72)
        print("[%d/%d] %s" % (i, len(STEPS), title))
        print("      $ %s %s" % (script, " ".join(args)))
        print("-" * 72)
        rc = subprocess.run([PY, "-X", "utf8", script] + args, cwd=ROOT).returncode
        if rc != 0:
            print("\n步骤 %d 失败（exit %d），后续步骤未执行。" % (i, rc))
            raise SystemExit(rc)
        print("")

    print("=" * 72)
    print("全部 %d 步通过。剩余手工事项：同步 `AGENTS.md`（目录树行数/文献篇数/准\\ 状态/" % len(STEPS))
    print("新增执行结果条目）、按 `备份\\entry_20261108\\异常登记_待复核_20261108.md` 留痕、四段汇报。")
    print("=" * 72)


if __name__ == "__main__":
    main()
