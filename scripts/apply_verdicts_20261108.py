# -*- coding: utf-8 -*-
"""汇总逐篇判定并按三分类落地（合规→准，存疑→待，不符→待\\_不符合）。

只移动、不删除；每次移动写回滚台账。
用法：
  D:\\Python\\python.exe -X utf8 scripts\\apply_verdicts_20261108.py          # 只汇总出 CSV，不动文件
  D:\\Python\\python.exe -X utf8 scripts\\apply_verdicts_20261108.py run      # 执行移动
"""
import csv
import json
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JUDGE = os.path.join(ROOT, "temp", "judge_20261108")
ZHUN = os.path.join(ROOT, "准")
DAI = os.path.join(ROOT, "待")
BUHE = os.path.join(DAI, "_不符合")

GOOD = {"合规", "合规-基准组"}
DOUBT = {"存疑-待补证"}
BAD = {"不符", "重复-已入库", "重复-库内重复"}


def locate(fn):
    for d in (ZHUN, DAI, BUHE):
        p = os.path.join(d, fn)
        if os.path.exists(p):
            return p
    return None


def main():
    run = len(sys.argv) > 1 and sys.argv[1] == "run"
    with open(os.path.join(JUDGE, "manifest.json"), encoding="utf-8") as f:
        man = json.load(f)
    ovp = os.path.join(JUDGE, "dup_override.json")
    override = {}
    if os.path.exists(ovp):
        with open(ovp, encoding="utf-8") as f:
            override = json.load(f)
    rows = []
    missing = []
    for m in man:
        vp = os.path.join(JUDGE, "verdicts", m["idx"] + ".json")
        v = {}
        if os.path.exists(vp):
            try:
                with open(vp, encoding="utf-8") as f:
                    v = json.load(f)
            except Exception:
                v = {}
        if not v.get("判定"):
            missing.append(m["idx"])
            continue
        ov = override.get(m["file"])
        if ov:
            v = dict(v)
            v["触发条款"] = "去重：" + ov["判定"]
            v["判定"] = ov["判定"]
            v["依据"] = ov.get("依据", "")
        cur = locate(m["file"])
        rows.append(
            {
                "idx": m["idx"],
                "文件": m["file"],
                "现位置": os.path.basename(os.path.dirname(cur)) if cur else "缺失",
                "判定": v.get("判定", ""),
                "触发条款": v.get("触发条款", ""),
                "介质原文": v.get("介质原文", ""),
                "成形与组织原文": v.get("成形与组织原文", ""),
                "热处理原文": v.get("热处理原文", ""),
                "Epit与参比": v.get("Epit与参比", ""),
                "依据": v.get("依据", ""),
                "原路径": cur or "",
            }
        )

    order = {"合规": 0, "合规-基准组": 1, "存疑-待补证": 2, "不符": 3, "重复-已入库": 4}
    rows.sort(key=lambda r: (order.get(r["判定"], 9), r["文件"]))
    out = os.path.join(ROOT, "逐篇判级_20261108.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=[c for c in rows[0].keys() if c != "原路径"])
        w.writeheader()
        for r in rows:
            w.writerow({k: v for k, v in r.items() if k != "原路径"})

    from collections import Counter

    c = Counter(r["判定"] for r in rows)
    print("已判 %d 篇，缺判定 %d 篇" % (len(rows), len(missing)))
    for k in ["合规", "合规-基准组", "存疑-待补证", "不符", "重复-已入库"]:
        print("  %-8s %d" % (k, c.get(k, 0)))
    if missing:
        print("  缺判定 idx:", ",".join(missing[:40]))

    if not run:
        print("（未执行移动；加 run 参数才动文件）")
        return

    os.makedirs(BUHE, exist_ok=True)
    ledger = []
    moved = 0
    for r in rows:
        src = r["原路径"]
        if not src:
            continue
        if r["判定"] in GOOD:
            dst_dir = ZHUN
        elif r["判定"] in DOUBT:
            dst_dir = DAI
        elif r["判定"] in BAD:
            dst_dir = BUHE
        else:
            continue
        dst = os.path.join(dst_dir, r["文件"])
        if os.path.normcase(os.path.dirname(src)) == os.path.normcase(dst_dir):
            ledger.append([r["文件"], r["判定"], src, dst, "原地不动", ""])
            continue
        if os.path.exists(dst):
            ledger.append([r["文件"], r["判定"], src, dst, "目标已存在，跳过", ""])
            continue
        shutil.move(src, dst)
        moved += 1
        ledger.append([r["文件"], r["判定"], src, dst, "已移动", "移回原路径即可回滚"])

    lp = os.path.join(ROOT, "逐篇判级_移动台账_20261108.csv")
    with open(lp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件", "判定", "原路径", "新路径", "动作", "回滚方式"])
        w.writerows(ledger)
    print("实际移动 %d 个文件；台账 %s" % (moved, os.path.basename(lp)))
    print("准=%d 待=%d 待\\_不符合=%d"
          % (len([x for x in os.listdir(ZHUN) if x.lower().endswith('.pdf')]),
             len([x for x in os.listdir(DAI) if x.lower().endswith('.pdf')]),
             len([x for x in os.listdir(BUHE) if x.lower().endswith('.pdf')])))


if __name__ == "__main__":
    main()
