# -*- coding: utf-8 -*-
r"""把 audit_inputs.json 压成可读摘要，供设计清洗规则。

只写 备份\清洗1_20261109\audit_summary.txt；不碰 temp。
"""
import json
import os
import re
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "备份", "清洗1_20261109", "audit_inputs.json")
OUT = os.path.join(ROOT, "备份", "清洗1_20261109", "audit_summary.txt")


def fmt(counter):
    return "\n".join("  %4d | %s" % (n, repr(k)) for k, n in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0])))


def main():
    with open(SRC, "r", encoding="utf-8") as f:
        d = json.load(f)
    out = []
    out.append("META dims=%s rows=%s blocks=%s merges=%s zoom=%s freeze=%s" % (
        d["meta"]["dims"], d["meta"]["data_rows"], d["meta"]["blocks"],
        d["meta"]["merges"], d["meta"]["zoom"], d["meta"]["freeze"]))
    out.append("H1 " + repr(d["meta"]["h1"]))
    out.append("H2 " + repr(d["meta"]["h2"]))
    out.append("\nEPIT_CLASS " + repr(d["epit_class"]))
    out.append("\nEPIT_FORMS count=%d / 只列非纯数值、非单值、非N/A的可疑写法" % len(d["epit_forms"]))
    sus = []
    for k, n in d["epit_forms"].items():
        s = str(k).strip()
        if s.upper().startswith("N/A"):
            continue
        if re.fullmatch(r"[-+−‑‒–—]?\d+(?:\.\d+)?", s):
            continue
        sus.append((s, n))
    out.append("\n".join("  %4d | %s" % (n, repr(s)) for s, n in sorted(sus, key=lambda kv: (-kv[1], kv[0]))))
    out.append("\nRE_FORMS count=%d\n%s" % (len(d["re_forms"]), fmt(Counter(d["re_forms"]))))
    out.append("\nTEMP_FORMS count=%d\n%s" % (len(d["temp_forms"]), fmt(Counter(d["temp_forms"]))))
    out.append("\nSTAB_FORMS count=%d\n%s" % (len(d["stab_forms"]), fmt(Counter(d["stab_forms"]))))
    out.append("\nCE_FORMS count=%d\n%s" % (len(d["ce_forms"]), fmt(Counter(d["ce_forms"]))))
    out.append("\nWE_FORMS count=%d / 前 30\n%s" % (
        len(d["we_forms"]), "\n".join("  %4d | %s" % (n, repr(k)) for k, n in Counter(d["we_forms"]).most_common(30))))
    out.append("\nFE_FORMS count=%d\n%s" % (len(d["fe_forms"]), fmt(Counter(d["fe_forms"]))))
    out.append("\nELEMENT unusual forms (非纯数字/非0/未提及/余量):")
    for c, forms in d["elem_forms"].items():
        bad = []
        for k, n in forms.items():
            s = str(k).strip()
            if s in ("", "0", "未提及", "—", "-", "None") or re.fullmatch(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", s):
                continue
            bad.append((s, n))
        if bad:
            out.append("  col %s: %d 种" % (c, len(bad)))
            out.extend("    %4d | %s" % (n, repr(s)) for s, n in sorted(bad, key=lambda kv: (-kv[1], kv[0]))[:60])
    out.append("\nLONG_NOTES %d 块" % len(d["long_notes"]))
    for k, v in sorted(d["long_notes"].items(), key=lambda kv: int(kv[0])):
        out.append("  %s len=%d | %s" % (k, v["len"], repr(v["z"][:120])))
    out.append("\nBLOCK Y/Z 空值统计:")
    noy = [k for k, v in d["blocks"].items() if not v["y"]]
    noz = [k for k, v in d["blocks"].items() if not v["z"]]
    out.append("  Y空 %d: %s" % (len(noy), noy[:30]))
    out.append("  Z空 %d: %s" % (len(noz), noz[:30]))
    txt = "\n".join(out) + "\n"
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(txt)
    print("summary ->", OUT, os.path.getsize(OUT), "B")
    print(txt[:7000])


if __name__ == "__main__":
    main()
