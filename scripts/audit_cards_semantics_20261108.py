# -*- coding: utf-8 -*-
r"""第 9 批卡片「语义合规」只读审计（不修卡，只列待裁决清单）。

检查已回收卡片是否触犯批准计划里的四条硬约束：
  1) 判级为「合规-基准组」的卡片，notes 是否写了「仅取基准组」；
  2) 是否出现 `→SCE` / `0.1989 V vs SHE` / `换算` 等本该留给 1.xlsx 的换算说明；
  3) 备注是否 >120 字（违反前段人工短写习惯）；
  4) 是否含工作日志抬头 / 对 1·2.xlsx 的指示 / 介质复述等不该写进 0.xlsx 的文字；
  5) Epit 为 N/A 或成分全「未提及」的行数（只登记，不删）。
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")

BAD_WORDS = ["→SCE", "0.1989", "换算至 SCE", "换算成 SCE", "统一 SCE",
             "清洗时", "1.xlsx", "2.xlsx", "后续处理", "本轮入库",
             "口径存疑", "【提取方式】", "【Fe 记法】", "【用途】",
             "3.5 wt.% NaCl 溶液", "3.5% NaCl 溶液", "介质为 3.5"]


def main():
    with open(os.path.join(WORK, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    by_no = {m["newno"]: m for m in manifest}

    base_missing, conv, long_notes, log_words, na, all_unknown = [], [], [], [], [], []
    have = 0
    for no in range(64, 147):
        p = os.path.join(WORK, "cards", "%03d.json" % no)
        if not os.path.exists(p):
            continue
        have += 1
        try:
            with open(p, encoding="utf-8") as f:
                c = json.load(f)
        except Exception as e:
            print("[JSON 错] %03d %s" % (no, e))
            continue
        notes = str(c.get("notes") or "")
        rows = c.get("rows") or []
        verdict = by_no[no]["verdict"]
        if verdict == "合规-基准组" and "仅取基准组" not in notes:
            base_missing.append((no, notes[:70]))
        blob = notes + "\n" + json.dumps(rows, ensure_ascii=False)
        hits = [w for w in BAD_WORDS if w in blob]
        if hits:
            log_words.append((no, hits, notes[:90]))
        if "→SCE" in blob or "0.1989" in blob:
            conv.append(no)
        if len(notes) > 120:
            long_notes.append((no, len(notes)))
        for i, r in enumerate(rows):
            if str(r.get("Epit", "")).strip().upper() in ("N/A", "NA", "N.A.", "—", "-"):
                na.append((no, i, str(r.get("sample"))[:30]))
            elems = ["C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
                     "Ti", "Nb", "Co", "Sn", "Fe"]
            if all(str(r.get(k, "")).strip() in ("未提及", "", "None") for k in elems):
                all_unknown.append((no, i, str(r.get("sample"))[:30]))

    print("已审计卡片 %d / 83" % have)
    print("\n[1] 合规-基准组但未写「仅取基准组」：%d 篇" % len(base_missing))
    for no, s in base_missing:
        print("  - %03d %s" % (no, s))
    print("\n[2] 出现换算说明：%d 篇 %s" % (len(conv), conv))
    print("\n[3] 备注 >120 字：%d 篇 %s" % (len(long_notes), long_notes))
    print("\n[4] 命中不该写进 0.xlsx 的措辞：%d 篇" % len(log_words))
    for no, hits, s in log_words[:30]:
        print("  - %03d %s | %s" % (no, hits, s))
    print("\n[5] Epit=N/A：%d 行 %s" % (len(na), na))
    print("[6] 15 元素全「未提及」：%d 行 %s" % (len(all_unknown), all_unknown))


if __name__ == "__main__":
    main()
