# -*- coding: utf-8 -*-
"""262 篇去重（确定性规则，不靠模糊相似度）。

规则：
  R1 与 `文献\\` 已入库 63 篇「归一化开头 300 字完全相同」→ 重复-已入库
  R2 262 篇内部「归一化开头 300 字完全相同」或「MD5 相同」→ 保留一篇，其余 重复-库内重复
保留优先级：判定为合规的优先 → 已在 准\\ 的优先 → idx 小的优先。

输出 temp\\judge_20261108\\dup_override.json 与 逐篇去重复核_20261108.csv
用法：D:\\Python\\python.exe -X utf8 scripts\\dedupe_final_20261108.py
"""
import csv
import hashlib
import json
import os
import re
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JUDGE = os.path.join(ROOT, "temp", "judge_20261108")
ZHUN = os.path.join(ROOT, "准")
DAI = os.path.join(ROOT, "待")
SRC = os.path.join(ROOT, "temp", "rescreen_20261108", "text_all.jsonl")

GOOD = {"合规", "合规-基准组"}


def norm(s):
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", (s or "").lower())


def head_key(rec):
    k = norm(rec.get("head") or (rec.get("text") or "")[:900])
    return k[:300] if len(k) >= 60 else ""


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    recs = []
    with open(SRC, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                recs.append(json.loads(line))
    dai = [r for r in recs if r["tag"] == "DAI"]
    wen = [r for r in recs if r["tag"] == "WEN"]
    wen_key = {}
    for r in wen:
        k = head_key(r)
        if k:
            wen_key.setdefault(k, r["file"])

    with open(os.path.join(JUDGE, "manifest.json"), encoding="utf-8") as f:
        man = json.load(f)
    idx_of = {m["file"]: m["idx"] for m in man}

    verd = {}
    for m in man:
        p = os.path.join(JUDGE, "verdicts", m["idx"] + ".json")
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
                    verd[m["file"]] = json.load(f).get("判定", "")
            except Exception:
                verd[m["file"]] = ""

    def cur_dir(fn):
        for d in (ZHUN, DAI):
            if os.path.exists(os.path.join(d, fn)):
                return os.path.basename(d)
        return "缺失"

    override = {}
    # R1
    for r in dai:
        k = head_key(r)
        if k and k in wen_key:
            override[r["file"]] = {"判定": "重复-已入库",
                                   "依据": "开头与 文献\\%s 完全相同（同篇）" % wen_key[k]}

    # R2
    groups = defaultdict(list)
    for r in dai:
        if r["file"] in override:
            continue
        k = head_key(r)
        if k:
            groups["H:" + k].append(r["file"])
        p = None
        for d in (ZHUN, DAI):
            q = os.path.join(d, r["file"])
            if os.path.exists(q):
                p = q
                break
        if p:
            groups["M:" + md5(p)].append(r["file"])

    done = set()
    for key, files in groups.items():
        files = [f for f in files if f not in override and f not in done]
        if len(files) < 2:
            continue
        files.sort(key=lambda f: (0 if verd.get(f) in GOOD else 1,
                                  0 if cur_dir(f) == "准" else 1,
                                  idx_of.get(f, "")))
        keep = files[0]
        for f in files[1:]:
            override[f] = {"判定": "重复-库内重复",
                           "依据": "与 %s 同篇（%s）" % (keep, "开头相同" if key.startswith("H:") else "MD5 相同")}
            done.add(f)

    with open(os.path.join(JUDGE, "dup_override.json"), "w", encoding="utf-8") as f:
        json.dump(override, f, ensure_ascii=False, indent=1)
    with open(os.path.join(ROOT, "逐篇去重复核_20261108.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "文件", "现位置", "去重判定", "说明"])
        for m in man:
            fn = m["file"]
            ov = override.get(fn)
            w.writerow([m["idx"], fn, cur_dir(fn), ov["判定"] if ov else "", ov["依据"] if ov else ""])

    print("去重覆盖 %d 篇：%s" % (len(override), dict(Counter(v["判定"] for v in override.values()))))
    for fn, v in sorted(override.items(), key=lambda x: (x[1]["判定"], x[0])):
        print("  %-12s %s | %s" % (v["判定"], fn[:62], v["依据"]))


if __name__ == "__main__":
    main()
