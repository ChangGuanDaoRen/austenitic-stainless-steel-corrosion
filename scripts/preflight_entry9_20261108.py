# -*- coding: utf-8 -*-
r"""第 9 批入库前置只读体检：源 PDF / 目标文献 / 已回收卡片的 check() 兼容性。

不写任何文件；输出会在卡片收齐前提前暴露「file 与台账不符 / 缺 23 键 / flag」
等会阻断 `build_entry9_20261108.py check` 的问题。
"""
import hashlib
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
CARD_DIR = os.path.join(WORK, "cards")
SRC_DIR = os.path.join(ROOT, "准")
LIT_DIR = os.path.join(ROOT, "文献")
BASE = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
TARGET = os.path.join(ROOT, "0.xlsx")

FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    by_no = {m["newno"]: m for m in manifest}
    print("台账：%d 篇 / 编号 %d–%d" % (len(manifest), min(by_no), max(by_no)))

    # 1) 源 PDF 与目标文献
    src_bad, md5_groups = [], {}
    for no in sorted(by_no):
        name = by_no[no]["file"]
        src = os.path.join(SRC_DIR, name)
        if not os.path.exists(src):
            src_bad.append("缺源 %d：%s" % (no, name))
            continue
        with open(src, "rb") as f:
            if f.read(5) != b"%PDF-":
                src_bad.append("非 PDF %d：%s" % (no, name))
        md5_groups.setdefault(md5(src), []).append(no)
    dups = {k: v for k, v in md5_groups.items() if len(v) > 1}
    targets_exist = [n for n in range(64, 147)
                     if os.path.exists(os.path.join(LIT_DIR, "%02d.pdf" % n))]
    lit_pdfs = [x for x in os.listdir(LIT_DIR) if x.lower().endswith(".pdf")]
    print("源 PDF：%d 篇可读；坏件/缺失 %d；完全同 MD5 的重复组 %d %s"
          % (sum(len(v) for v in md5_groups.values()), len(src_bad),
             len(dups), [v for v in dups.values()]))
    print("文献\\：%d 篇 PDF；64–146 目标已存在 %d 个 %s"
          % (len(lit_pdfs), len(targets_exist), targets_exist[:10]))
    if os.path.exists(BASE) and os.path.exists(TARGET):
        print("0.xlsx 与底本 MD5：%s / %s（必须相同，否则 build/commit 会阻断）"
              % (md5(TARGET), md5(BASE)))
    else:
        print("0.xlsx 或第9批入库前底本缺失，build/commit 会阻断")
    for x in src_bad[:20]:
        print("  -", x)

    # 2) 已回收卡片对 check() 的兼容性
    have, would_block, would_flag, warn = [], [], [], []
    for no in range(64, 147):
        p = os.path.join(CARD_DIR, "%03d.json" % no)
        if not os.path.exists(p):
            continue
        have.append(no)
        try:
            with open(p, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            would_block.append((no, "JSON: %s" % e))
            continue
        probs = []
        if c.get("newno") not in (no, str(no)):
            probs.append("newno=%r" % c.get("newno"))
        f = str(c.get("file") or "").strip()
        if f != by_no[no]["file"]:
            probs.append("file 与台账不符：%r" % f[:70])
        if not str(c.get("citation") or "").strip():
            probs.append("citation 空")
        if not str(c.get("notes") or "").strip():
            probs.append("notes 空")
        rows = c.get("rows")
        if not isinstance(rows, list) or len(rows) < 1:
            probs.append("rows 空")
            rows = []
        for i, r in enumerate(rows):
            if not isinstance(r, dict):
                probs.append("row%d 非对象" % i)
                continue
            missing = [k for k in FIELDS
                       if k not in r or r[k] is None or str(r[k]).strip() == ""]
            if missing:
                probs.append("row%d 缺 %s" % (i, ",".join(missing)))
            extra = [k for k in r if k not in FIELDS]
            if extra:
                warn.append((no, "row%d 多字段 %s（check 会忽略）" % (i, ",".join(extra))))
        flag = str(c.get("flag") or "").strip()
        if flag:
            would_flag.append((no, flag[:90]))
        if probs:
            would_block.append((no, "；".join(probs)[:200]))

    missing = [n for n in range(64, 147) if n not in have]
    print("\n卡片：已回收 %d / 缺 %d" % (len(have), len(missing)))
    print("缺号：" + ",".join(str(n) for n in missing))
    print("check() 会阻断的卡片 %d 张；带 flag 的 %d 张；多字段警告 %d 条"
          % (len(would_block), len(would_flag), len(warn)))
    for no, why in would_block[:30]:
        print("  [阻断] %03d %s" % (no, why))
    for no, why in would_flag[:30]:
        print("  [flag] %03d %s" % (no, why))
    for no, why in warn[:12]:
        print("  [警告] %03d %s" % (no, why))


if __name__ == "__main__":
    main()
