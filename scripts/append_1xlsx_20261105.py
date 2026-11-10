# -*- coding: utf-8 -*-
r"""把第 3 批（0.xlsx 行 261-285 / 文献 56-63）清洗后追加到 1.xlsx 的 Sheet1。

严格沿用 build_1_20261105.py 的既有规则，只作用于新增 25 行：
  元素 未提及/空→0；<x/≤x→x；a-b→中值；ppm→/1e4
  Fe「余量」→ 100 − Σ14 元素（2 位）
  Epit 统一 SCE：SCE Δ=0；Ag/AgCl(sat.KCl) Δ=−0.042
  温度 室温/未提及→25；25±1℃→25；298 K→25；30℃/60℃→原值
  稳定时间 统一分钟：未提及→30；3600 s OCP→60；1 h OCP→60；阴极极化段取末段
  对电极 铂族→铂；文献编号逐行填
两处新增特例（都在报告里留痕）：
  文献61 的 Epit「<-0.35」是上限值 → 按「上限取该数值」清洗为 −0.35
  文献60 的参比已核定为饱和 KCl Ag/AgCl → −0.042 V，Epit 0.343 → 0.301

1.xlsx 按用户 11-05 要求**只保留 Sheet1**，本脚本不重建任何辅助表。
用法：
  D:/Python/python.exe scripts/append_1xlsx_20261105.py            dry（写校验件）
  D:/Python/python.exe scripts/append_1xlsx_20261105.py --commit    写回 1.xlsx
"""
import hashlib
import os
import re
import shutil
import sys
from copy import copy

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

WS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC0 = os.path.join(WS, "0.xlsx")
BACKUP = os.path.join(WS, "备份", "1_基准_第3批清洗前_20261105.xlsx")
TARGET = os.path.join(WS, "1.xlsx")
STAGE = os.path.join(WS, "temp", "lit", "entry3", "1_追加校验3.xlsx")
REPORT = os.path.join(WS, "temp", "lit", "entry3", "清洗明细_第3批.csv")

NCOL = 24
OLD_MAX_ROW = 252
NEW_ROWS = 25
NEW_MAX_ROW = OLD_MAX_ROW + NEW_ROWS        # 277
SRC_FIRST_ROW = 261                          # 0.xlsx 中第 3 批起始行

C_ID, C_NAME = 1, 2
C_ELEM = list(range(3, 17))                  # C..Sn
C_FE = 17
C_PROC, C_WE, C_RE, C_CE = 18, 19, 20, 21
C_TEMP, C_EPIT, C_STAB = 22, 23, 24

RE_MAP = {
    "SCE": (0.0, "SCE"),
    "SCE(饱和KCl甘汞)": (0.0, "SCE"),
    "SCE(饱和甘汞)": (0.0, "SCE"),
    "Ag/AgCl": (-0.031, "Ag/AgCl(未注明->按3M)"),
    "Ag/AgCl(3M KCl)": (-0.031, "Ag/AgCl(3M KCl)"),
    "Ag/AgCl(sat.KCl)": (-0.042, "Ag/AgCl(饱和KCl)"),
    "Ag/AgCl(饱和 KCl)": (-0.042, "Ag/AgCl(饱和KCl)"),
    "Ag/AgCl(饱和 KCl, 0.1989 V vs SHE)": (-0.0421, "Ag/AgCl(饱和,0.1989V vs SHE)"),
    "Ag/AgCl(饱和KCl)→SCE": (0.0, "Ag/AgCl(饱和) 原文已换算"),
}

NUM_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*$")
CAP_RE = re.compile(r"^\s*[<≤]\s*(-?[0-9]+(?:\.[0-9]+)?)\s*$")
RANGE_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*[-~–]\s*([0-9]+(?:\.[0-9]+)?)\s*$")
PPM_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:ppm|PPM)\s*$")


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def r2(x, n=2):
    v = round(float(x), n)
    return int(v) if v == int(v) else v


def parse_elem(v):
    if v is None:
        return 0.0, "0(空)"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v), "num"
    s = str(v).strip()
    if s in ("", "未提及", "—", "-", "None"):
        return 0.0, "0(未提及)"
    for pat, tag in ((CAP_RE, "cap"), (RANGE_RE, "mid"), (PPM_RE, "ppm"), (NUM_RE, "num")):
        m = pat.match(s)
        if m:
            if tag == "cap":
                return float(m.group(1)), "cap"
            if tag == "mid":
                return round((float(m.group(1)) + float(m.group(2))) / 2, 4), "mid"
            if tag == "ppm":
                return round(float(m.group(1)) / 1e4, 6), "ppm"
            return float(m.group(1)), "num"
    return 0.0, "unknown:" + s


def parse_temp(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "室温", "None"):
        return 25, "默认25"
    m = re.fullmatch(r"约?\s*([0-9.]+)\s*±\s*[0-9.]+\s*℃", s)
    if m:
        return r2(m.group(1)), "原文"
    m = re.fullmatch(r"([0-9.]+)\s*[-~–]\s*([0-9.]+)\s*℃", s)
    if m:
        return r2((float(m.group(1)) + float(m.group(2))) / 2), "原文"
    m = re.fullmatch(r"约?\s*([0-9.]+)\s*℃", s)
    if m:
        return r2(m.group(1)), "原文"
    m = re.fullmatch(r"([0-9.]+)\s*K", s)
    if m:
        return int(round(float(m.group(1)) - 273.15)), "原文(K)"
    return None, "未识别:" + s


def _mins(val, unit):
    if unit.startswith("h"):
        return val * 60
    if unit.startswith("s"):
        return val / 60
    return val


def parse_stab(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "None"):
        return 30, "默认30"
    if "预浸" in s and "OCP" in s:
        tot = sum(_mins(float(a), b) for a, b in re.findall(r"([0-9.]+)\s*(h|min|s)", s))
        return r2(tot), "原文(预浸+OCP相加)"
    if s.startswith("阴极极化"):
        ms = re.findall(r"([0-9.]+)\s*(h|min|s)", s)
        if ms:
            return r2(_mins(float(ms[-1][0]), ms[-1][1])), "原文(阴极极化段)"
    ms = re.findall(r"([0-9.]+)\s*(h|min|s)", s)
    if ms:
        a, b = ms[-1]
        return r2(_mins(float(a), b)), "原文"
    return None, "未识别:" + s


def parse_ce(v):
    s = "" if v is None else str(v).strip()
    if s in ("", "未提及", "None"):
        return "铂", "默认铂"
    if "石墨" in s:
        return "石墨", "原文"
    if "铂" in s or s.upper().startswith("PT"):
        return "铂", "原文"
    return None, "未识别:" + s


def clean_rows(ws0):
    """清洗 0.xlsx 的 261-285 行，返回 (records, flags)。"""
    recs, flags = [], []
    cur = None
    for r in range(SRC_FIRST_ROW, ws0.max_row + 1):
        raw_id = ws0.cell(r, C_ID).value
        if raw_id not in (None, ""):
            cur = int(float(raw_id))
        rec = {C_ID: cur, C_NAME: str(ws0.cell(r, C_NAME).value)}
        zero_cols, unk = [], []
        for c in C_ELEM:
            val, tag = parse_elem(ws0.cell(r, c).value)
            rec[c] = r2(val, 6) if tag in ("cap", "mid", "ppm") else val
            if tag.startswith("0"):
                zero_cols.append(ws0.cell(2, c).value)
            elif tag.startswith("unknown"):
                unk.append(tag)
        fe_raw = ws0.cell(r, C_FE).value
        if isinstance(fe_raw, (int, float)) and not isinstance(fe_raw, bool):
            rec[C_FE] = r2(fe_raw, 2)
            fe_src = "实测"
        else:
            rec[C_FE] = r2(100 - sum(rec[c] for c in C_ELEM), 2)
            fe_src = "反算"

        re_raw = str(ws0.cell(r, C_RE).value or "").strip()
        if re_raw not in RE_MAP:
            raise SystemExit("参比未覆盖：0.xlsx r%d %r" % (r, re_raw))
        delta, re_tag = RE_MAP[re_raw]
        epit_raw = ws0.cell(r, C_EPIT).value
        if isinstance(epit_raw, str) and CAP_RE.match(epit_raw):
            epit_val, epit_note = float(CAP_RE.match(epit_raw).group(1)), "Epit按上限取值"
        else:
            epit_val, epit_note = float(epit_raw), ""
        rec[C_EPIT] = r2(epit_val + delta, 4)
        rec[C_RE] = "SCE"

        temp, temp_src = parse_temp(ws0.cell(r, C_TEMP).value)
        stab, stab_src = parse_stab(ws0.cell(r, C_STAB).value)
        ce, ce_src = parse_ce(ws0.cell(r, C_CE).value)
        if None in (temp, stab, ce):
            raise SystemExit("解析失败 0.xlsx r%d: %r %r %r" % (r, temp_src, stab_src, ce_src))
        rec[C_TEMP], rec[C_STAB], rec[C_CE] = temp, stab, ce
        rec[C_PROC] = ws0.cell(r, C_PROC).value
        rec[C_WE] = ws0.cell(r, C_WE).value
        recs.append(rec)
        tot = sum(rec[c] for c in C_ELEM) + rec[C_FE]
        notes = list(unk)
        if epit_note:
            notes.append(epit_note)
        if fe_src == "反算" and any(x in zero_cols for x in ("C", "Si", "Mn", "Cr", "Ni")):
            notes.append("Fe反算受未报主元素影响")
        if abs(tot - 100) > 1:
            notes.append("成分和=%.2f" % tot)
        flags.append([cur, rec[C_NAME], r, ",".join(zero_cols) if zero_cols else "—",
                      len(zero_cols), fe_src, re_tag, delta, temp_src, stab_src, ce_src,
                      "；".join(notes) if notes else ""])
    if len(recs) != NEW_ROWS:
        raise SystemExit("清洗条数 %d != 25" % len(recs))
    return recs, flags


def style_sig(cell):
    f, a, b = cell.font, cell.alignment, cell.border
    return (f.name, f.sz, f.b, f.i,
            a.horizontal, a.vertical, a.wrap_text, cell.number_format,
            b.left.style, b.right.style, b.top.style, b.bottom.style)


def build(recs):
    wb = openpyxl.load_workbook(BACKUP)
    ws = wb["Sheet1"]
    if ws.max_row != OLD_MAX_ROW or ws.max_column != NCOL:
        raise SystemExit("1.xlsx 原件结构变了：%d 行 / %d 列（应为 252 / 24）" %
                         (ws.max_row, ws.max_column))
    tpl = {c: copy(ws.cell(OLD_MAX_ROW, c)._style) for c in range(1, NCOL + 1)}
    for i, rec in enumerate(recs, start=OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            cell = ws.cell(i, c)
            cell._style = copy(tpl[c])
            cell.value = rec.get(c)
    wb.save(STAGE)


def verify(recs, path):
    problems = []
    old = openpyxl.load_workbook(BACKUP)
    new = openpyxl.load_workbook(path)
    if old.sheetnames != ["Sheet1"] or new.sheetnames != ["Sheet1"]:
        problems.append("sheet 列表异常：%r -> %r" % (old.sheetnames, new.sheetnames))
    o, n = old["Sheet1"], new["Sheet1"]
    if n.max_row != NEW_MAX_ROW or n.max_column != NCOL:
        problems.append("成品尺寸不对：%d 行 / %d 列" % (n.max_row, n.max_column))

    changed = 0
    for r in range(1, OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            a, b = o.cell(r, c), n.cell(r, c)
            if a.value != b.value or style_sig(a) != style_sig(b):
                changed += 1
                if changed <= 5:
                    problems.append("原有格被改：r%d c%d" % (r, c))
    if changed:
        problems.append("前 %d 行共有 %d 处差异" % (OLD_MAX_ROW, changed))

    om = set(str(x) for x in o.merged_cells.ranges)
    nm = set(str(x) for x in n.merged_cells.ranges)
    if om != nm:
        problems.append("合并区变了：%s / %s" % (sorted(om - nm), sorted(nm - om)))
    ow = {k: v.width for k, v in o.column_dimensions.items()}
    nw = {k: v.width for k, v in n.column_dimensions.items()}
    if ow != nw:
        problems.append("列宽变了")
    if o.freeze_panes != n.freeze_panes:
        problems.append("冻结窗格变了：%r -> %r" % (o.freeze_panes, n.freeze_panes))
    if o.sheet_view.zoomScale != n.sheet_view.zoomScale:
        problems.append("缩放比例变了")

    for i, rec in enumerate(recs, start=OLD_MAX_ROW + 1):
        for c in range(1, NCOL + 1):
            if n.cell(i, c).value != rec.get(c):
                problems.append("新增行 r%d c%d 写入不符：%r != %r" %
                                (i, c, n.cell(i, c).value, rec.get(c)))
                break
        if style_sig(n.cell(i, 2)) != style_sig(o.cell(OLD_MAX_ROW, 2)):
            problems.append("新增行 r%d 字体/对齐与既有行不一致" % i)
        if not isinstance(n.cell(i, C_EPIT).value, (int, float)):
            problems.append("新增行 r%d Epit 不是数值" % i)
    return problems


def main():
    if not os.path.exists(BACKUP):
        shutil.copy2(TARGET, BACKUP)
        print("已创建基准备份：" + BACKUP)
    if md5(TARGET) not in (md5(BACKUP), md5(STAGE) if os.path.exists(STAGE) else ""):
        raise SystemExit("1.xlsx 已被其它进程改动，停止")

    ws0 = openpyxl.load_workbook(SRC0)["Sheet1"]
    if ws0.max_row != 285:
        raise SystemExit("0.xlsx 末行不是 285，先完成第 3 批入库")
    recs, flags = clean_rows(ws0)

    print("清洗 25 条（0.xlsx r261-285 -> 1.xlsx r253-277）")
    for f, rec in zip(flags, recs):
        print("  文献%-2d %-26s Epit %-8s→%-8s T=%-4s 稳定=%-6s Fe=%-7s %s" %
              (f[0], str(f[1])[:26], str(ws0.cell(f[2], C_EPIT).value), rec[C_EPIT],
               rec[C_TEMP], rec[C_STAB], rec[C_FE],
               ("[" + f[11] + "]") if f[11] else ""))

    build(recs)
    problems = verify(recs, STAGE)
    if problems:
        print("校验未通过，共 %d 项：" % len(problems))
        for x in problems[:40]:
            print("  -", x)
        raise SystemExit(2)

    import csv
    with open(REPORT, "w", encoding="utf-8-sig", newline="") as fo:
        w = csv.writer(fo)
        w.writerow(["文献编号", "样本", "0.xlsx行号", "0填充元素", "0填充格数",
                    "Fe来源", "Epit原参比", "换算Δ(V)", "温度", "稳定时间", "对电极", "异常/说明"])
        w.writerows(flags)
    print("\n校验通过：前 %d 行零差异 / 只有 Sheet1 / 10 合并区、冻结 C3、列宽、缩放全保留 / 25 行逐格回读一致"
          % OLD_MAX_ROW)
    print("清洗明细 -> " + REPORT)

    if "--commit" in sys.argv:
        if md5(TARGET) not in (md5(BACKUP), md5(STAGE)):
            raise SystemExit("提交前发现 1.xlsx 被改动，停止")
        shutil.copy2(STAGE, TARGET)
        problems2 = verify(recs, TARGET)
        if problems2:
            print("写回后校验失败（原件仍在 备份\\ 可恢复）：")
            for x in problems2[:20]:
                print("  -", x)
            raise SystemExit(3)
        print("[commit] 已写回 1.xlsx，MD5=%s；写回后复验通过" % md5(TARGET))
    else:
        print("[dry] 未改 1.xlsx；确认无误后加 --commit 提交")


if __name__ == "__main__":
    main()
