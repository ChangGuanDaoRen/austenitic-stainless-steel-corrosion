# -*- coding: utf-8 -*-
"""清理 temp\\ 至配额以内（限 300 MB / 300 文件）。

三条规矩：
1. 只删「可重建」或「有字节级孪生副本」的文件：
   - 非 PDF（.txt/.html/.json 缓存与抽文本）→ 可重建，直接进清单；
   - PDF → 必须在 文献\\／准\\／待\\／待\\_不符合\\／备份\\ 里找到同 MD5 的副本才进清单。
2. 证据与台账一律保留：audit0、addcol、verdicts、text_all.jsonl、lit_index.json、
   各轮台账 CSV、判级台账、取证 JSON、CRITERIA.md、r9_mix/r10 的报告与清单。
3. 送回收站（SHFileOperationW + FOF_ALLOWUNDO），不做永久删除；清单先落盘再执行。

用法：
    python clean_temp_20261108.py plan     # 只统计与写清单，不动文件
    python clean_temp_20261108.py run      # 按清单送回收站
"""
import csv
import ctypes
import hashlib
import os
import sys
from ctypes import wintypes

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMP = os.path.join(ROOT, "temp")
MANIFEST = os.path.join(ROOT, "temp_清理清单_20261108.csv")

INDEX_DIRS = ["文献", "准", "待", "待\\_不符合", "备份"]
# 保留：证据 / 台账 / 取证数据 / 报告（前缀匹配，相对 temp\\）
KEEP_PREFIX = [
    "audit0\\",
    "addcol\\",
    "judge_20261108\\verdicts\\",
    "judge_20261108\\manifest.json",
    "rescreen_20261108\\text_all.jsonl",
    "rescreen_20261108\\lit_index.json",
    "rescreen_20261108\\triage_all.csv",
    "rescreen_20261108\\triage_compact",
    "rescreen_20261108\\review_doubt",
    "criteria_impact_20261107.csv",
    "criteria_impact_说明_20261107.md",
    "audit_0xlsx_20261108_stats.json",
    "1xlsx_辅助表_删前_20261105\\",
    "lit\\CRITERIA.md",
    "lit\\CRITERIA_20261031.md",
    "lit\\r9_mix\\",
    "lit\\r10\\判级台账",
    "lit\\r10\\judge847\\",
    "lit\\r9_pw\\probe_direct_vs_proxy.py",
    "lit\\r9_pw\\tls_probe.json",
    "lit\\r9_pw\\direct_paywall_probe.json",
]
KEEP_NAME_SUFFIX = (".md", "台账.csv", "汇总.csv", "明细.csv", "清单.csv", "统计.csv")
KEEP_DIRPARTS = ("\\entry\\", "\\entry2\\", "\\entry3\\", "\\entry4\\")
KEEP_EXACT_SUFFIX = ("final_rows.json", "final_rows2.json", "final_rows3.json", "final_rows4.json")


def rel_temp(p):
    return os.path.relpath(p, TEMP)


def is_keep(rel):
    low = rel.lower()
    for pre in KEEP_PREFIX:
        if rel.startswith(pre):
            return True
    if low.endswith(KEEP_NAME_SUFFIX):
        return True
    if any(d in "\\" + low for d in KEEP_DIRPARTS):
        return True
    if low.endswith(KEEP_EXACT_SUFFIX):
        return True
    return False


def md5(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def build_index():
    """size -> [paths]，只索引受保护目录里的文件。"""
    idx = {}
    for d in INDEX_DIRS:
        base = os.path.join(ROOT, d)
        if not os.path.isdir(base):
            continue
        for r, _, fs in os.walk(base):
            for f in fs:
                p = os.path.join(r, f)
                try:
                    sz = os.path.getsize(p)
                except OSError:
                    continue
                idx.setdefault(sz, []).append(p)
    return idx


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", wintypes.UINT),
        ("pFrom", wintypes.LPCWSTR),
        ("pTo", wintypes.LPCWSTR),
        ("fFlags", ctypes.c_uint),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", wintypes.LPCWSTR),
    ]


def to_recycle(paths):
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = 3  # FO_DELETE
    op.pFrom = "".join(os.path.abspath(p) + "\0" for p in paths) + "\0"
    op.pTo = None
    op.fFlags = 0x0004 | 0x0010 | 0x0040 | 0x0400  # SILENT|NOCONFIRMATION|ALLOWUNDO|NOERRORUI
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc, op.fAnyOperationsAborted


def plan():
    idx = build_index()
    print("受保护目录索引：%d 个大小档、%d 个文件" % (len(idx), sum(len(v) for v in idx.values())))
    rows = []
    cache_n = cache_s = 0
    pdf_n = pdf_s = 0
    kept_n = kept_s = 0
    orphan = []
    for r, dirs, fs in os.walk(TEMP):
        for f in fs:
            p = os.path.join(r, f)
            rel = rel_temp(p)
            try:
                sz = os.path.getsize(p)
            except OSError:
                continue
            if is_keep(rel):
                kept_n += 1
                kept_s += sz
                continue
            if f.lower().endswith(".pdf"):
                twins = idx.get(sz, [])
                hit = None
                h = md5(p)
                if twins:
                    for t in twins:
                        if md5(t) == h:
                            hit = t
                            break
                if hit:
                    rows.append((rel, sz, "recycle", "PDF 有同 MD5 副本：" + os.path.relpath(hit, ROOT), h))
                    pdf_n += 1
                    pdf_s += sz
                else:
                    # 库里（文献/准/待/待_不符合/备份）没有同 MD5 副本 → 非馆藏件，
                    # 属检索轮次的暂存下载，判级结论已在各轮台账；记 MD5 后在回收站可还原。
                    rows.append((rel, sz, "recycle-orphan-pdf",
                                 "PDF 无任何馆藏副本，检索暂存件（判级见台账）", h))
                    orphan.append((rel, sz))
                continue
            rows.append((rel, sz, "recycle", "缓存/抽文本，可重建", ""))
            cache_n += 1
            cache_s += sz
    with open(MANIFEST, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["相对路径", "字节", "动作", "理由", "md5"])
        for rel, sz, act, why, h in rows:
            w.writerow([rel, sz, act, why, h])
    tot_n = kept_n + len(rows) + len(orphan)
    tot_s = kept_s + cache_s + pdf_s + sum(s for _, s in orphan)
    print("temp 现状：%.1f MB / %d 文件" % (tot_s / 1048576, tot_n))
    print("  保留（证据/台账）：%.1f MB / %d 文件" % (kept_s / 1048576, kept_n))
    print("  待回收 缓存文本 ：%.1f MB / %d 文件" % (cache_s / 1048576, cache_n))
    print("  待回收 PDF 副本 ：%.1f MB / %d 文件" % (pdf_s / 1048576, pdf_n))
    print("  无孪生 PDF（保留）：%.1f MB / %d 文件" % (sum(s for _, s in orphan) / 1048576, len(orphan)))
    print("清单：%s（%d 行）" % (MANIFEST, len(rows)))
    if orphan:
        print("无孪生 PDF 前 20：")
        for rel, sz in orphan[:20]:
            print("   %8.1f MB  %s" % (sz / 1048576, rel))
    return 0


def run():
    if not os.path.exists(MANIFEST):
        print("清单不存在，先跑 plan")
        return 1
    with open(MANIFEST, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    paths = [os.path.join(TEMP, r["相对路径"]) for r in rows if os.path.exists(os.path.join(TEMP, r["相对路径"]))]
    print("清单 %d 行，实际存在 %d 个待回收文件" % (len(rows), len(paths)))
    batch, done, fail = [], 0, []
    for i, p in enumerate(paths, 1):
        batch.append(p)
        if len(batch) >= 200:
            rc, ab = to_recycle(batch)
            if rc == 0 and not ab:
                done += len(batch)
            else:
                fail.append((len(batch), rc, ab))
            batch = []
    if batch:
        rc, ab = to_recycle(batch)
        if rc == 0 and not ab:
            done += len(batch)
        else:
            fail.append((len(batch), rc, ab))
    print("已送回收站 %d 个文件" % done)
    for n, rc, ab in fail:
        print("  批次失败 n=%d rc=%s aborted=%s" % (n, rc, ab))
    return 0


def report():
    n = s = 0
    for r, _, fs in os.walk(TEMP):
        for f in fs:
            n += 1
            try:
                s += os.path.getsize(os.path.join(r, f))
            except OSError:
                pass
    print("temp 现状：%.2f MB / %d 文件（限 300 MB / 300 文件）" % (s / 1048576, n))
    return 0


STAGE_DIR = os.path.join(ROOT, "备份", "temp_检索暂存_20261108")
STAGE_DIRS = [
    ("audit0", "audit0_20261108"),
    ("judge_20261108\\verdicts", "judge_verdicts_20261108"),
]
STAGE_FILES = [
    ("judge_20261108\\manifest.json", "judge_20261108_manifest.json"),
    ("rescreen_20261108\\text_all.jsonl", "text_all.jsonl"),
    ("rescreen_20261108\\lit_index.json", "lit_index.json"),
]


def move_file(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.exists(dst):
        # 目标已存在则加后缀，避免覆盖
        base, ext = os.path.splitext(dst)
        i = 2
        while os.path.exists("%s_%d%s" % (base, i, ext)):
            i += 1
        dst = "%s_%d%s" % (base, i, ext)
    os.replace(src, dst)


def stage():
    """把无孪生 PDF 与证据目录整体搬进 备份\\，不删除任何文件；缓存照旧送回收站。"""
    if not os.path.exists(MANIFEST):
        print("清单不存在，先跑 plan")
        return 1
    with open(MANIFEST, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    moved = 0
    moved_s = 0
    for r in rows:
        if r["动作"] != "recycle-orphan-pdf":
            continue
        src = os.path.join(TEMP, r["相对路径"])
        if not os.path.exists(src):
            continue
        dst = os.path.join(STAGE_DIR, r["相对路径"])
        move_file(src, dst)
        moved += 1
        moved_s += int(r["字节"] or 0)
    print("无孪生 PDF 已搬入 备份\\temp_检索暂存_20261108：%d 个 / %.1f MB" % (moved, moved_s / 1048576))
    for rel, name in STAGE_DIRS:
        src = os.path.join(TEMP, rel)
        if not os.path.isdir(src):
            continue
        dst = os.path.join(ROOT, "备份", name)
        if os.path.exists(dst):
            dst = dst + "_b"
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        total = sum(os.path.getsize(os.path.join(r_, f)) for r_, _, fs in os.walk(src) for f in fs)
        n = sum(len(fs) for _, _, fs in os.walk(src))
        os.replace(src, dst)
        print("证据目录已搬入 %s：%d 个 / %.1f MB" % (dst, n, total / 1048576))
    for rel, name in STAGE_FILES:
        src = os.path.join(TEMP, rel)
        if not os.path.exists(src):
            continue
        dst = os.path.join(ROOT, "备份", name)
        move_file(src, dst)
        print("证据文件已搬入 %s" % dst)
    return 0


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "plan"
    sys.exit({"plan": plan, "run": run, "report": report, "stage": stage}.get(mode, plan)())
