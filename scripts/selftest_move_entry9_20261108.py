# -*- coding: utf-8 -*-
r"""离线自测「移动重编号」脚本：在临时目录里造 83 篇假 PDF，验证 shutil.move 腾空 准\、冲突不覆盖。

不碰真实 0.xlsx / 文献\ / 准\ / 待\；临时目录跑完自动删除。
"""
import csv
import hashlib
import importlib.util
import json
import os
import shutil
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT_REAL = os.path.dirname(HERE)
SRC_SCRIPT = os.path.join(HERE, "copy_lit_and_report_20261108.py")


def load_module():
    spec = importlib.util.spec_from_file_location("move_entry9_under_test", SRC_SCRIPT)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def fake_tree(root, conflict=False):
    work = os.path.join(root, "备份", "entry_20261108")
    src_dir = os.path.join(root, "准")
    dst_dir = os.path.join(root, "文献")
    os.makedirs(work)
    os.makedirs(src_dir)
    os.makedirs(dst_dir)
    names = ["fake_%03d.pdf" % n for n in range(64, 147)]
    numbering = []
    data = []
    for no, name in zip(range(64, 147), names):
        body = (b"%PDF-1.4\n" + ("content of %d\n" % no).encode("utf-8") * 7)
        with open(os.path.join(src_dir, name), "wb") as f:
            f.write(body)
        numbering.append({
            "新编号": str(no), "原文件名": name,
            "判级": "合规" if no % 2 else "合规-基准组",
            "全文字数": str(len(body)), "目标文献编号": "文献\\%03d.pdf" % no,
        })
        data.append({
            "no": no, "file": name, "verdict": numbering[-1]["判级"],
            "citation": "Fake citation %d" % no,
            "notes": "仅取基准组；成分取自 Table 1。" if no % 2 == 0 else "成分取自 Table 1。",
            "flag": "", "rows": [{
                "sample": "S-%d" % no, "C": "0.03", "Si": "0.5", "N": "未提及",
                "P": "未提及", "S": "未提及", "Cu": "未提及", "Mn": "1.2",
                "Cr": "18", "Mo": "2", "Ni": "10", "Ti": "未提及", "Nb": "未提及",
                "Co": "未提及", "Sn": "未提及", "Fe": "余量",
                "process": "熔铸+固溶", "WE": "试样", "RE": "SCE", "CE": "铂片",
                "temp": "25℃", "Epit": "0.3", "stable": "30 min",
            }],
        })
    with open(os.path.join(work, "numbering.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(numbering[0].keys()))
        w.writeheader()
        w.writerows(numbering)
    with open(os.path.join(work, "final_rows.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    if conflict:
        with open(os.path.join(dst_dir, "64.pdf"), "wb") as f:
            f.write(b"%PDF-1.4\nDIFFERENT TARGET CONTENT\n")
    return work, src_dir, dst_dir, names


def run_case(m, root, conflict=False):
    work, src_dir, dst_dir, names = fake_tree(root, conflict)
    m.ROOT = root
    m.WORK = work
    m.NUMCSV = os.path.join(work, "numbering.csv")
    m.FINAL = os.path.join(work, "final_rows.json")
    m.LEDGER = os.path.join(root, "准_入0与重编号对照_20261108.csv")
    m.REPORT = os.path.join(root, "准文献入库对照_20261108.md")
    m.SRC_DIR = src_dir
    m.DST_DIR = dst_dir
    old_argv = sys.argv
    sys.argv = [SRC_SCRIPT, "run"]
    try:
        m.main()
        rc = 0
    except SystemExit as e:
        rc = e.code or 0
    finally:
        sys.argv = old_argv
    got = sorted(x for x in os.listdir(dst_dir) if x.lower().endswith(".pdf"))
    remain = sorted(x for x in os.listdir(src_dir) if x.lower().endswith(".pdf"))
    with open(m.LEDGER, encoding="utf-8-sig") as f:
        ledger = list(csv.DictReader(f))
    return rc, got, remain, ledger


def main():
    m = load_module()
    results = []

    root1 = tempfile.mkdtemp(prefix="entry9_move_ok_", dir=os.path.join(ROOT_REAL, "temp"))
    try:
        rc, got, remain, ledger = run_case(m, root1, conflict=False)
        want = sorted("%02d.pdf" % n for n in range(64, 147))
        ok = (rc == 0 and got == want and remain == []
              and len(ledger) == 83
              and all(x["动作"] == "已移动" for x in ledger))
        results.append(("happy path 移动腾空", ok,
                        "rc=%s 文献=%d 准\\剩余=%d 台账=%d 全已移动=%s"
                        % (rc, len(got), len(remain), len(ledger),
                           all(x["动作"] == "已移动" for x in ledger))))
    finally:
        shutil.rmtree(root1, ignore_errors=True)

    root2 = tempfile.mkdtemp(prefix="entry9_move_conf_", dir=os.path.join(ROOT_REAL, "temp"))
    try:
        rc, got, remain, ledger = run_case(m, root2, conflict=True)
        l64 = next(x for x in ledger if int(x["新编号"]) == 64)
        ok = (rc == 1 and len(remain) == 1 and remain == ["fake_064.pdf"]
              and l64["动作"] == "冲突，未覆盖"
              and len(got) == 83)
        results.append(("冲突不覆盖不删源", ok,
                        "rc=%s 准\\剩余=%s 64动作=%s 文献=%d"
                        % (rc, remain, l64["动作"], len(got))))
    finally:
        shutil.rmtree(root2, ignore_errors=True)

    bad = 0
    for name, ok, detail in results:
        print(("PASS " if ok else "FAIL ") + name + " -> " + detail)
        if not ok:
            bad += 1
    print("selftest_move: %d/%d passed" % (len(results) - bad, len(results)))
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
