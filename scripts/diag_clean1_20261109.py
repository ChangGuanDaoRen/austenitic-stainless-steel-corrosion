# -*- coding: utf-8 -*-
r"""重建 1.xlsx 前的聚焦只读诊断（扩展版）。

在原有诊断基础上，补：
- Epit 文本全部可疑写法里，哪些「无单位但 |v|>2」应判 mV；
- 元素异常形态全表；
- 块级备注长度分布与短备注里仍含工作日志措辞的块；
- Y/Z 块首值是否完整。
只输出到控制台，不写文件。
"""
import json
import os
import re
import unicodedata

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
AUDIT = os.path.join(ROOT, "备份", "清洗1_20261109", "audit_inputs.json")


def main():
    with open(AUDIT, "r", encoding="utf-8") as f:
        d = json.load(f)
    forms = d["epit_forms"]
    print("=== Epit 文本形态：无单位且 |v|>2（应判 mV）===")
    pat = re.compile(r"^([-+−‑‒–—]?\d+(?:\.\d+)?)\s*(?:±\s*[\d.]+)?\s*$")
    mV_like = []
    for k, n in forms.items():
        s = unicodedata.normalize("NFKC", str(k)).strip()
        if s.upper().startswith("N/A"):
            continue
        m = pat.match(s)
        if not m:
            continue
        try:
            v = float(m.group(1).replace("−", "-").replace("–", "-"))
        except ValueError:
            continue
        if abs(v) > 2:
            mV_like.append((s, n, v))
    print("共 %d 种：" % len(mV_like))
    for s, n, v in sorted(mV_like, key=lambda x: -abs(x[2]))[:60]:
        print("   %3d | %-22r | %s" % (n, s, v))

    print("\n=== Epit 含范围 to/~ 或括号 ± 的写法 ===")
    for k, n in sorted(forms.items(), key=lambda kv: kv[0]):
        s = str(k)
        if re.search(r"\bto\b|~|～|\(±|（±", s):
            print("   %3d | %r" % (n, s))

    print("\n=== 元素异常形态（完整）===")
    for c, forms2 in d["elem_forms"].items():
        bad = []
        for k, n in forms2.items():
            s = unicodedata.normalize("NFKC", str(k)).strip()
            if s in ("", "0", "未提及", "—", "-", "None") or re.fullmatch(r"[-+]?\d+(?:\.\d+)?", s):
                continue
            bad.append((s, n))
        if bad:
            print("   col %-3s %d 种: %s" % (c, len(bad), ", ".join("%r×%d" % (s, n) for s, n in bad)))

    print("\n=== 备注长度分布 ===")
    lens = [(int(k), v["z_len"]) for k, v in d["blocks"].items()]
    buckets = {"0": 0, "1-40": 0, "41-80": 0, "81-120": 0, ">120": 0}
    for _, L in lens:
        if L == 0:
            buckets["0"] += 1
        elif L <= 40:
            buckets["1-40"] += 1
        elif L <= 80:
            buckets["41-80"] += 1
        elif L <= 120:
            buckets["81-120"] += 1
        else:
            buckets[">120"] += 1
    print(buckets)

    print("\n=== 短备注(≤120)里仍含工作日志/口径推演措辞的块 ===")
    BAD = ["本轮入库", "原件与逐条取证", "口径存疑", "按师兄", "1.xlsx", "2.xlsx",
           "清洗时", "清洗版", "A/Z/AA", "已放入 evidence", "source_quote",
           "AA 备注", "待裁决", "待师兄", "按用户", "本核定", "台账初稿",
           "【用途】", "【Fe 记法】", "【主程复核】", "数据待补", "清单写"]
    for k, v in sorted(d["blocks"].items(), key=lambda kv: int(kv[0])):
        z = v["z"]
        if v["z_len"] > 120:
            continue
        hit = [b for b in BAD if b in z]
        if hit:
            print("   %-4s len=%-4d hit=%s | %r" % (k, v["z_len"], hit, z[:100]))

    print("\n=== Y/Z 块首值完整性 ===")
    print("   Y 空块 %d / Z 空块 %d / 总块 %d" % (
        sum(1 for v in d["blocks"].values() if not v["y"]),
        sum(1 for v in d["blocks"].values() if not v["z"]),
        len(d["blocks"])))


if __name__ == "__main__":
    main()
