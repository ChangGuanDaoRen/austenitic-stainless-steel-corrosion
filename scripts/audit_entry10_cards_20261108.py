# -*- coding: utf-8 -*-
r"""第10批卡片只读语义审计（不修卡，只列待复核）。

检查：顶层禁 flag、23 键/全字符串、newno/file 与台账一致、notes 短写、
不该进入 0.xlsx 的换算/日志措辞、Epit=N/A、15 元素全「未提及」。
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry10_20261108")
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]
ELEMS = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
         "Ti", "Nb", "Co", "Sn", "Fe"]
BAD_WORDS = ["→SCE", "0.1989", "换算至 SCE", "换算成 SCE", "统一 SCE",
             "清洗时", "1.xlsx", "2.xlsx", "后续处理", "本轮入库",
             "口径存疑", "【提取方式】", "【Fe 记法】", "【用途】",
             "3.5 wt.% NaCl 溶液", "3.5% NaCl 溶液", "介质为 3.5"]


def main():
    with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    by_no = {m["newno"]: m for m in manifest}

    blocking, warns, na, all_unknown = [], [], [], []
    have, rows_n = 0, 0
    for no in range(147, 160):
        p = os.path.join(WORK, "cards", "%03d.json" % no)
        if not os.path.exists(p):
            print("[缺卡] %03d" % no)
            continue
        have += 1
        try:
            with open(p, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            blocking.append("%03d JSON 错：%s" % (no, e))
            continue
        if c.get("flag"):
            blocking.append("%03d 顶层含 flag：%s" % (no, c["flag"]))
        if c.get("newno") not in (no, str(no)):
            blocking.append("%03d newno=%r" % (no, c.get("newno")))
        if c.get("file") != by_no[no]["file"]:
            blocking.append("%03d file 与台账不符：%r" % (no, c.get("file")))
        notes = str(c.get("notes") or "").strip()
        citation = str(c.get("citation") or "").strip()
        if not notes:
            blocking.append("%03d notes 空" % no)
        if not citation:
            blocking.append("%03d citation 空" % no)
        rows = c.get("rows")
        if not isinstance(rows, list) or len(rows) < 1:
            blocking.append("%03d rows 零行" % no)
            rows = []
        rows_n += len(rows)
        for i, r in enumerate(rows):
            if not isinstance(r, dict):
                blocking.append("%03d row%d 非对象" % (no, i))
                continue
            miss = [k for k in FIELDS if k not in r or not isinstance(r[k], str) or not r[k].strip()]
            if miss:
                blocking.append("%03d row%d 缺/非字符串：%s" % (no, i, ",".join(miss)))
            extra = [k for k in r if k not in FIELDS]
            if extra:
                warns.append("%03d row%d 多字段（脚本忽略）：%s" % (no, i, ",".join(extra)))
            if str(r.get("Epit", "")).strip().upper() in ("N/A", "NA", "N.A.", "—", "-"):
                na.append((no, i, str(r.get("sample"))[:30]))
            if all(str(r.get(k, "")).strip() in ("未提及", "", "None") for k in ELEMS):
                all_unknown.append((no, i, str(r.get("sample"))[:30]))
        blob = notes + "\n" + json.dumps(rows, ensure_ascii=False)
        hits = [w for w in BAD_WORDS if w in blob]
        if hits:
            warns.append("%03d 命中不该写进 0.xlsx 的措辞 %s" % (no, hits))
        if len(notes) > 120:
            warns.append("%03d 备注 %d 字（>120）" % (no, len(notes)))
        if len(notes) > 80:
            warns.append("%03d 备注 %d 字（>80，建议压缩）" % (no, len(notes)))

    print("已审计卡片 %d / 13 / 数据行 %d" % (have, rows_n))
    print("\n[阻断] %d 项" % len(blocking))
    for x in blocking:
        print("  -", x)
    print("\n[提醒] %d 项" % len(warns))
    for x in warns:
        print("  -", x)
    print("\n[Epit=N/A] %d 行 %s" % (len(na), na))
    print("[15 元素全未提及] %d 行 %s" % (len(all_unknown), all_unknown))
    if len(blocking) or have != 13:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
