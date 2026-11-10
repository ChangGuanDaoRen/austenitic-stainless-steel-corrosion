# -*- coding: utf-8 -*-
r"""第 10 批入库（准\ 当前 13 篇 -> 0.xlsx，文献编号 147-159）确定性流水线。

子命令：
  prep      备份 0.xlsx + 生成 13 篇取证输入（text/digest/RULES/manifest/numbering）；不改 xlsx
  check     质检 13 张卡片 -> final_rows.json + 质检 CSV；不改 xlsx
  build     以 备份\0_基准_第10批入库前_20261108.xlsx 为底生成校验件并逐格校验
  commit    校验通过且 0.xlsx 未改时写回 0.xlsx 并二次复验
  move      13 篇从 准\ 移动重编号到 文献\147.pdf-159.pdf + 台账/对照报告
  verify    闭环校验 0.xlsx <-> 文献\01-159 <-> 台账源 MD5
  report    出「需要你看的文献」报告（待\ 63 篇：32 存疑 + 31 不符）
  finalize  gate 通过后按序 check/build/commit/move/verify/report

复用 scripts\build_entry9_20261108.py 的 load_cards/build/verify/cellval/FIELDS/md5
（importlib 载入后 monkeypatch 常量，不改历史脚本）。
"""
import csv
import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry10_20261108")
TEXT_DIR = os.path.join(WORK, "text")
DIG_DIR = os.path.join(WORK, "digest")
CARD_DIR = os.path.join(WORK, "cards")
BACKUP = os.path.join(ROOT, "备份", "0_基准_第10批入库前_20261108.xlsx")
TARGET = os.path.join(ROOT, "0.xlsx")
STAGE = os.path.join(WORK, "0_追加校验10.xlsx")
FINAL = os.path.join(WORK, "final_rows.json")
QC = os.path.join(WORK, "cards_质检_第10批_20261108.csv")
MANIFEST = os.path.join(WORK, "manifest.json")
NUMBERING = os.path.join(WORK, "numbering.csv")
RULES10 = os.path.join(WORK, "RULES.md")
EXAMPLES = os.path.join(WORK, "examples.json")

ZHUN = os.path.join(ROOT, "准")
LIT = os.path.join(ROOT, "文献")
DAI = os.path.join(ROOT, "待")
LEDGER = os.path.join(ROOT, "准_入0与重编号对照_第10批_20261108.csv")
ENTRY_REPORT = os.path.join(ROOT, "准文献入库对照_第10批_20261108.md")
VERIFY_REPORT = os.path.join(ROOT, "0_第10批入库校验_20261108.md")
NEEDS_REPORT = os.path.join(ROOT, "需人工裁决文献报告_20261108.md")
IMA_CSV = os.path.join(ROOT, "ima_核查_校正判定_20261108.csv")
TEXT_ALL = os.path.join(ROOT, "备份", "text_all.jsonl")

FIRST_NO, LAST_NO = 147, 159
N_PAPERS = LAST_NO - FIRST_NO + 1
NCOL = 26
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]

MARK = "================ 论文全文 ================\n"


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def pdf_ok(path):
    with open(path, "rb") as f:
        return f.read(5) == b"%PDF-"


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------- prep
def cmd_prep():
    os.makedirs(TEXT_DIR, exist_ok=True)
    os.makedirs(DIG_DIR, exist_ok=True)
    os.makedirs(CARD_DIR, exist_ok=True)

    # 1) 备份现表（只在不存在时留底本，绝不覆盖）
    if not os.path.exists(BACKUP):
        shutil.copy2(TARGET, BACKUP)
        print("[prep] 已备份 0.xlsx ->", BACKUP, "MD5=", md5(BACKUP))
    else:
        print("[prep] 底本已存在，校验一致性：", md5(TARGET) == md5(BACKUP))
        if md5(TARGET) != md5(BACKUP):
            raise SystemExit("0.xlsx 与既有备份不一致，停止（先人工核对）")

    # 2) 全文索引
    by_file = {}
    with open(TEXT_ALL, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            if rec.get("tag") == "DAI":
                by_file[rec.get("file")] = rec

    # 3) 准\ 13 篇
    files = sorted(x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf"))
    if len(files) != N_PAPERS:
        raise SystemExit("准\\ PDF 数应为 %d，实测 %d：%r" % (N_PAPERS, len(files), files))
    missing = [x for x in files if x not in by_file]
    if missing:
        raise SystemExit("全文索引缺 %d 篇：%r" % (len(missing), missing))

    # 4) 前 3 块人工录入样例（延续第9批风格）
    wb = openpyxl.load_workbook(BACKUP, data_only=True)
    ws = wb["Sheet1"]
    ex = []
    for rr in range(3, 12):
        if isinstance(ws.cell(rr, 1).value, (int, float)) and len(ex) < 3:
            ex.append({"文献编号": ws.cell(rr, 1).value,
                       "样本": ws.cell(rr, 2).value,
                       "文献信息": ws.cell(rr, 25).value,
                       "备注": ws.cell(rr, 26).value})
    with open(EXAMPLES, "w", encoding="utf-8") as f:
        json.dump(ex, f, ensure_ascii=False, indent=1)

    # 5) RULES：复用第9批，改路径标题 + 追加「禁止 flag」补充（优先于上文）
    with open(os.path.join(ROOT, "备份", "entry_20261108", "RULES.md"), encoding="utf-8") as f:
        rules = f.read()
    rules = rules.replace("第9批入库", "第10批入库")
    rules = rules.replace(r"备份\entry_20261108", r"备份\entry10_20261108")
    rules = rules.replace('，顶层加 `"flag": "原因"`。', "，**不要加 flag**（见文末第10批补充）。")
    rules = rules.replace("完成后**只回一行** `ok:<rows数>` 或 `flag:<原因>`；",
                          "完成后**只回一行** `ok:<rows数>`；")
    rules = rules.replace('，可选 `flag`(str)。', "；**禁止** `flag` 字段。")
    rules += (
        "\n\n## 第10批补充（优先于上文）\n"
        "- **禁止输出 `flag`**：无论能否找到数值，顶层都不要出现 `flag`；找不到可录数值就出 "
        "1 行占位（元素/工艺等写「未提及」、Epit 写 `N/A`），并在 `notes` 说明原因。加 `flag` 会被质检阻断。\n"
        "- 新编号范围 **147–159**；顶层 `newno` 必须等于任务给出的编号，`file` 必须等于 digest 题头【原文件】的完整文件名。\n"
        "- 严禁渲染页面 / 图像数字化 / 写临时 png 或 py / 联网；数值只在图里、文本层没有时一律 Epit 记 `N/A`。\n"
    )
    with open(RULES10, "w", encoding="utf-8") as f:
        f.write(rules)

    # 6) text + digest + manifest + numbering
    KW = re.compile(
        r"3\.5\s*(wt|mass|%|g)|0\.6\s*(mol|M|kmol)|35\s*g/L|600\s*mM"
        r"|NaCl|氯化钠|chloride|海水|seawater|simulated seawater|D1141|模拟海水"
        r"|点蚀电位|孔蚀电位|击穿电位|pitting|Epit|Eb\b|breakdown|polarization|极化曲线|阳极"
        r"|SCE|Ag/AgCl|calomel|甘汞|参比电极|reference electrode|饱和"
        r"|固溶|退火|时效|敏化|稳定化|均匀化|淬火|水淬|铸|cast|锻|热轧|冷轧|轧制|as[- ]received"
        r"|austenit|奥氏体|duplex|双相|铁素体|马氏体|δ|delta"
        r"|成分|composition|chemical|wt\.%|wt%|mass%|Table\s*\d|表\s*\d|Fig\.?\s*\d|图\s*\d"
        r"|Cr|Ni|Mo|Mn|Si|Cu|Nb|Ti|Co|Sn|Fe|balance|Bal\.|余量"
        r"|OCP|稳定时间|浸泡|immersion|scan rate|扫描速率|mV|V\b|SHE"
        r"|铸造|熔炼|感应炉|连铸|铸锭|砂型|离心|精密铸造"
    )
    manifest, numbering = [], []
    dig_sizes = []
    for i, fn in enumerate(files):
        no = FIRST_NO + i
        idx = "%03d" % no
        rec = by_file[fn]
        text = rec.get("text") or ""
        doi = str(rec.get("doi") or "")
        hdr = ("【新编号】%d\n【原文件】%s\n【判级】合规（ima 二次核查校正后）\n【DOI】%s\n"
               "【全文字数】%d\n" % (no, fn, doi, len(text)))
        if len(text) > 90000:
            body = text[:60000] + "\n……（中段略，如需请再读 text/%s.txt）……\n" % idx + text[-28000:]
        else:
            body = text
        with open(os.path.join(TEXT_DIR, idx + ".txt"), "w", encoding="utf-8") as f:
            f.write(hdr + MARK + body)

        head, tail = text[:4500], text[-2500:]
        mid, used, seen = [], 0, set()
        for p in re.split(r"\n\s*\n|\n", text):
            p = p.strip()
            if used >= 26000:
                break
            if len(p) < 8 or not KW.search(p):
                continue
            if p[:100] in seen:
                continue
            seen.add(p[:100])
            seg = p[:1600]
            mid.append(seg)
            used += len(seg)
        dg = (hdr + "【紧凑 digest：开头 4500 + 关键词命中段 ≤26000 + 结尾 2500 字符】\n"
              "【如需完整表格或数值，请再读 text/%s.txt（read 的 offset/limit 分段）】\n\n"
              "===== 标题与摘要（开头） =====\n%s\n\n===== 关键证据段 =====\n%s\n\n"
              "===== 原文结尾 =====\n%s\n" % (idx, head, "\n---\n".join(mid), tail))
        with open(os.path.join(DIG_DIR, idx + ".txt"), "w", encoding="utf-8") as f:
            f.write(dg)
        dig_sizes.append(len(dg))

        manifest.append({"idx": idx, "newno": no, "file": fn, "verdict": "合规",
                         "chars": len(text), "doi": doi,
                         "text": os.path.join("text", idx + ".txt")})
        numbering.append({"新编号": no, "原文件名": fn, "判级": "合规",
                          "全文字数": len(text), "目标文献编号": r"文献\%02d.pdf" % no})

    with open(MANIFEST, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    with open(NUMBERING, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["新编号", "原文件名", "判级", "全文字数", "目标文献编号"])
        w.writeheader()
        w.writerows(numbering)

    print("[prep] 13 篇 text/digest 就绪，digest min/median/max=%d/%d/%d 字符，合计 %.2f MB"
          % (min(dig_sizes), sorted(dig_sizes)[len(dig_sizes)//2], max(dig_sizes),
             sum(dig_sizes)/1024/1024))
    print("[prep] manifest / numbering / RULES.md 已写；底本 dims=%d 行 / %d 列"
          % (ws.max_row, ws.max_column))
    print("[prep] 编号 %d-%d：" % (FIRST_NO, LAST_NO))
    for m in manifest:
        print("   ", m["idx"], m["file"])


# ---------------------------------------------------------------- build 模块（monkeypatch）
def patched():
    spec = importlib.util.spec_from_file_location(
        "build9", os.path.join(ROOT, "scripts", "build_entry9_20261108.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    base = openpyxl.load_workbook(BACKUP)["Sheet1"]
    base_rows = base.max_row
    # 与第9批一致：普通行样式取最后一行；块首 A/Y/Z 样式取最后一个文献块的块首行。
    block_starts = []
    for rr in range(3, base_rows + 1):
        v = base.cell(rr, 1).value
        if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
            block_starts.append(rr)
    if not block_starts:
        raise SystemExit("底本 A 列没有文献编号，停止")
    anchor = block_starts[-1]
    m.WORK, m.CARD_DIR = WORK, CARD_DIR
    m.BACKUP, m.TARGET, m.STAGE = BACKUP, TARGET, STAGE
    m.FINAL, m.QC = FINAL, QC
    m.FIRST_NO, m.LAST_NO = FIRST_NO, LAST_NO
    m.OLD_MAX_ROW = base_rows
    m.TPL_NORMAL_ROW = base_rows
    m.TPL_ANCHOR_ROW = anchor
    m.TPL_WRAP_ROW = 3
    m.NCOL = NCOL
    return m, base_rows, anchor


def load_final(m):
    if not os.path.exists(FINAL):
        raise SystemExit("缺 final_rows.json，先跑 check")
    with open(FINAL, encoding="utf-8") as f:
        data = json.load(f)
    if [p["no"] for p in data] != list(range(FIRST_NO, LAST_NO + 1)):
        raise SystemExit("final_rows.json 编号不连续（疑似质检未过留下的旧文件），重跑 check")
    bad = [p["no"] for p in data
           if p.get("flag") or not p.get("citation") or not p.get("notes")
           or not isinstance(p.get("rows"), list) or len(p["rows"]) < 1
           or any(any(k not in r or r[k] is None or str(r[k]).strip() == "" for k in FIELDS)
                  for r in p["rows"])]
    if bad:
        raise SystemExit("final_rows.json 有 %d 篇不合格（%r），重跑 check 并补齐卡片" % (len(bad), bad))
    return data


# ---------------------------------------------------------------- check
def cmd_check():
    m, base_rows, anchor = patched()
    data, qc = m.load_cards()
    status = {}
    for x in qc:
        status[x["状态"]] = status.get(x["状态"], 0) + 1
    total = sum(len(p["rows"]) for p in data)
    print("[check] 卡片 %d 篇 / 数据 %d 条 / 编号 %d-%d / 底本 %d 行 / 模板行 %d"
          % (len(data), total, FIRST_NO, LAST_NO, base_rows, anchor))
    print("[check] 状态分布：", status)
    with open(FINAL, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    blocking = [x for x in qc if not x["状态"].startswith("ok")] + [x for x in qc if x["行数"] < 1]
    print("[check] 质检表：", QC)
    if blocking:
        print("[check] 质检未通过，%d 篇需修/缺卡/flag/零行：" % len(blocking))
        for x in blocking[:15]:
            print("   -", x["新编号"], x["原文件名"][:45], "|", x["状态"], "|", x["问题"][:140])
        raise SystemExit(2)
    print("[check] 质检通过：%d 篇 / %d 条数据，全部可入库；已写 final_rows.json" % (len(data), total))


# ---------------------------------------------------------------- build / commit
def cmd_build():
    m, base_rows, anchor = patched()
    data = load_final(m)
    if md5(TARGET) != md5(BACKUP):
        raise SystemExit("0.xlsx 与第10批入库前备份不一致，停止（先人工核对）")
    total = sum(len(p["rows"]) for p in data)
    blocks = m.build(data)
    problems = m.verify(data, blocks, STAGE)
    print("[build] 数据 %d 篇 / %d 条，写入行 %d-%d；校验件 %s"
          % (len(data), total, base_rows + 1, base_rows + total, STAGE))
    if problems:
        print("[build] 校验未通过，共 %d 项：" % len(problems))
        for x in problems[:40]:
            print("   -", x)
        raise SystemExit(2)
    _nm = len(openpyxl.load_workbook(STAGE)["Sheet1"].merged_cells.ranges)
    _om = len(openpyxl.load_workbook(BACKUP)["Sheet1"].merged_cells.ranges)
    print("[build] 校验通过：前 %d 行零差异 / 原 %d 合并区全保留 / 新增 %d / 列宽缩放不变 / %d 行逐格回读一致"
          % (base_rows, _om, _nm - _om, total))
    print("[build] [dry] 未改 0.xlsx；确认无误后跑 commit")


def cmd_commit():
    m, base_rows, anchor = patched()
    data = load_final(m)
    if md5(TARGET) != md5(BACKUP):
        raise SystemExit("提交前发现 0.xlsx 被改动，停止")
    if not os.path.exists(STAGE):
        raise SystemExit("缺校验件，先跑 build")
    total = sum(len(p["rows"]) for p in data)
    blocks = m.build(data)
    problems = m.verify(data, blocks, STAGE)
    if problems:
        print("[commit] 校验件复验失败：")
        for x in problems[:20]:
            print("   -", x)
        raise SystemExit(2)
    shutil.copy2(STAGE, TARGET)
    problems2 = m.verify(data, blocks, TARGET)
    if problems2:
        print("[commit] 写回后校验失败（原件仍在 备份\\ 可恢复）：")
        for x in problems2[:20]:
            print("   -", x)
        raise SystemExit(3)
    print("[commit] 已写回 0.xlsx，MD5=%s；%d 篇 / %d 条；写回后复验通过"
          % (md5(TARGET), len(data), total))


# ---------------------------------------------------------------- move
def cmd_move():
    m, base_rows, anchor = patched()
    data = load_final(m)
    before_lit = len([x for x in os.listdir(LIT) if x.lower().endswith(".pdf")])
    items, problems = [], []
    for p in data:
        no = p["no"]
        src = os.path.join(ZHUN, p["file"])
        dst = os.path.join(LIT, "%02d.pdf" % no)
        rec = {"新编号": no, "原文件名": p["file"], "题录": p["citation"], "备注": p["notes"],
               "数据行数": len(p["rows"]), "样本名": "; ".join(str(r["sample"]) for r in p["rows"]),
               "源字节": os.path.getsize(src) if os.path.exists(src) else 0,
               "源MD5": md5(src) if os.path.exists(src) else "", "动作": "", "目标": "%02d.pdf" % no}
        if not os.path.exists(src):
            problems.append("%d 缺源：%s" % (no, p["file"]))
            rec["动作"] = "缺源，未移动"
        elif not pdf_ok(src):
            problems.append("%d 坏件：%s" % (no, p["file"]))
            rec["动作"] = "坏件，未移动"
        elif os.path.exists(dst):
            if md5(src) == md5(dst):
                problems.append("%d 目标已存在同 MD5（用户未授权删源/覆盖）：%s" % (no, p["file"]))
                rec["动作"] = "冲突保留源（目标已存在同MD5）"
            else:
                problems.append("%d 目标已存在且内容不同：%s" % (no, p["file"]))
                rec["动作"] = "冲突，未覆盖"
        else:
            src_md5 = md5(src)
            shutil.move(src, dst)
            if not (os.path.exists(dst) and not os.path.exists(src)
                    and md5(dst) == src_md5 and pdf_ok(dst)):
                problems.append("%d 移动后校验失败：%s" % (no, p["file"]))
                rec["动作"] = "移动后校验失败"
            else:
                rec["动作"] = "已移动"
        items.append(rec)

    with open(LEDGER, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(items[0].keys()))
        w.writeheader()
        w.writerows(items)
    moved = sum(1 for r in items if r["动作"] == "已移动")
    remain = len([x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf")])
    after_lit = len([x for x in os.listdir(LIT) if x.lower().endswith(".pdf")])

    L = []
    L.append("# 准文献入库对照（第 10 批，2026-11-08）")
    L.append("")
    L.append("状态：%s；用户 m02025「准中再入0」、m02064「执行」；按 m01578 沿用「移动」重编号。" % now())
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("`准\\` 中 13 篇为 ima 二次核查校正后判「合规」的文献（`ima_核查差异清单_20261108.md`）。"
             "本批把它们录入 `0.xlsx`（编号 147–159），并把 PDF 从 `准\\` **移动**重编号为 "
             "`文献\\147.pdf`–`159.pdf`，使「编号 → `文献\\NN.pdf`」溯源链继续成立。")
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("- 逐篇并行抽取 23 键卡片，只记原文初始信息（不做 SCE 换算、备注短写、26 列无 `AA`）。")
    L.append("- 以 `备份\\0_基准_第10批入库前_20261108.xlsx` 为底生成校验件，"
             "前 %d 行值与样式零差异、147–159 连续、合并区/列宽一致全过才写回。" % base_rows)
    L.append("- `shutil.move` 逐篇校验 `%PDF-` 头与 MD5；目标已存在则保留源、不覆盖、不删除。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("| 项 | 数值 |")
    L.append("| --- | --- |")
    L.append("| 文献篇数 | %d |" % len(items))
    L.append("| 数据条数 | %d |" % sum(r["数据行数"] for r in items))
    L.append("| 成功移动 | %d |" % moved)
    L.append("| 冲突/异常 | %d |" % len(problems))
    L.append("| `文献\\` PDF | %d → %d |" % (before_lit, after_lit))
    L.append("| `准\\` 剩余 PDF | %d |" % remain)
    L.append("| `0.xlsx` | %d 行 / 26 列（写回后） |" % (base_rows + sum(r["数据行数"] for r in items)))
    L.append("")
    L.append("逐篇明细（题录 / 备注 / 数据行数 / 源 MD5 / 动作）见 `准_入0与重编号对照_第10批_20261108.csv`。")
    L.append("")
    L.append("| 编号 | 原文件名 | 行数 | 动作 |")
    L.append("| --- | --- | --- | --- |")
    for r in items:
        L.append("| %d | %s | %d | %s |" % (r["新编号"], r["原文件名"], r["数据行数"], r["动作"]))
    L.append("")
    L.append("## 四、建议")
    L.append("")
    L.append("- `准\\` 腾空后，如需继续扩量，再从 `待\\` 的 63 篇需人工裁决件里按用户裁定挑取，另起一批编号。")
    L.append("- 本批不改 `1.xlsx` / `2.xlsx`；如需把 147–159 纳入清洗/建模，属独立任务。")
    L.append("- 回滚：底本 `备份\\0_基准_第10批入库前_20261108.xlsx` 与移动台账均可还原。")
    with open(ENTRY_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")

    print("[move] 文献 %d→%d / 准剩余 %d / 已移动 %d / 冲突异常 %d"
          % (before_lit, after_lit, remain, moved, len(problems)))
    print("[move] 台账 %s / 报告 %s" % (LEDGER, ENTRY_REPORT))
    for x in problems:
        print("   -", x)
    if problems:
        raise SystemExit(1)


# ---------------------------------------------------------------- verify（闭环）
def cmd_verify():
    m, base_rows, anchor = patched()
    data = load_final(m)
    problems = []
    old = openpyxl.load_workbook(BACKUP)["Sheet1"]
    cur = openpyxl.load_workbook(TARGET)["Sheet1"]
    total = sum(len(p["rows"]) for p in data)
    if cur.max_row != base_rows + total or cur.max_column != NCOL:
        problems.append("0.xlsx 结构不对：%d 行 / %d 列" % (cur.max_row, cur.max_column))
    changed = 0
    for rr in range(1, base_rows + 1):
        for c in range(1, NCOL + 1):
            if old.cell(rr, c).value != cur.cell(rr, c).value:
                changed += 1
    if changed:
        problems.append("前 %d 行有 %d 格被改" % (base_rows, changed))
    r = base_rows
    for p in data:
        start = r + 1
        if cur.cell(start, 1).value != p["no"]:
            problems.append("编号 %d 块首 A 列不符" % p["no"])
        if cur.cell(start, 25).value != p["citation"] or cur.cell(start, 26).value != p["notes"]:
            problems.append("编号 %d 题录/备注不符" % p["no"])
        for i, row in enumerate(p["rows"]):
            rr = start + i
            for j, key in enumerate(FIELDS):
                if cur.cell(rr, 2 + j).value != row[key]:
                    problems.append("编号 %d %s 字段 r%d 不符" % (p["no"], key, rr))
                    break
        r += len(p["rows"])
    # 文献目录 01..159 连续 + 新增 147..159 MD5 与台账一致
    lit_pdfs = set(x for x in os.listdir(LIT) if x.lower().endswith(".pdf"))
    want = set("%02d.pdf" % i for i in range(1, LAST_NO + 1))
    if lit_pdfs != want:
        problems.append("文献目录不连续：缺 %r / 多 %r" % (sorted(want - lit_pdfs)[:5], sorted(lit_pdfs - want)[:5]))
    led = read_csv(LEDGER)
    for row in led:
        no = int(row["新编号"])
        dst = os.path.join(LIT, "%02d.pdf" % no)
        if not os.path.exists(dst) or md5(dst) != row["源MD5"] or not pdf_ok(dst):
            problems.append("文献 %d 与台账源 MD5 不一致" % no)
    # 说明/校验报告
    L = ["# 0.xlsx 第 10 批入库校验（2026-11-08）", "", "时间：%s。" % now(), "",
         "## 一、背景", "", "第 10 批把 `准\\` 13 篇录入 `0.xlsx`（编号 147–159）并移动重编号 PDF。", "",
         "## 二、方法", "", "只读比对：底本 vs 现表前 %d 行逐格、新增块逐字段回读、文献目录 01–%d 连续、"
         "新增 PDF 与移动台账源 MD5 一致。" % (base_rows, LAST_NO), "",
         "## 三、结论", "", "| 项 | 数值 |", "| --- | --- |",
         "| 现表 | %d 行 / %d 列 |" % (cur.max_row, cur.max_column),
         "| 新增 | %d 篇 / %d 条 / 编号 %d–%d |" % (len(data), total, FIRST_NO, LAST_NO),
         "| 文献目录 | %d 篇 |" % len(lit_pdfs),
         "| 问题项 | %d |" % len(problems), "",
         "## 四、建议", "", "问题 0 项即可交付；如需回滚，底本在 `备份\\0_基准_第10批入库前_20261108.xlsx`。"]
    if problems:
        L.append("")
        L.append("问题明细：")
        for x in problems[:30]:
            L.append("- " + x)
    with open(VERIFY_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print("[verify] 0.xlsx %d 行/%d 列 / 新增 %d 篇 %d 条 / 文献 %d 篇 / 问题 %d 项 / 报告 %s"
          % (cur.max_row, cur.max_column, len(data), total, len(lit_pdfs), len(problems), VERIFY_REPORT))
    for x in problems[:20]:
        print("   -", x)
    if problems:
        raise SystemExit(2)


# ---------------------------------------------------------------- 需要你看的文献报告
DISPUTED = {
    "15": ("是否允许 A5 行级拆分", "回原文只取 S31254 全奥氏体行、排除 SDSS 真双相行；若同意则另起一批入库"),
    "16": ("补 Table 2 浓度水平", "原文 3.5 wt% 仅见参考文献、2^3 因子含非 3.5 组、LPBF 增材主体排除；"
           "取得 Table 2 后确认是否有纯 3.5% 组且 W 316L 为独立未处理熔铸基准组，再按 A4 例外裁决"),
    "26": ("确认理由校正、结论不变", "UNSM 属表面机械处理不违 N1；真否决点是 N4「退火+650℃时效」两步；"
           "避免误伤本批已入库的表面机械滚压数据行"),
    "28": ("确认改判不符", "全文仅 Tafel 外推、无 pitting potential/Epit 数值；已由存疑改不符，无异议可移出"),
    "33": ("是否允许 A5 行级拆分", "回原文只取 304 奥氏体行、排除 436 铁素体行；若同意则另起一批入库"),
    "73": ("确认理由校正、结论不变", "真否决点是 N3 浓度序列（3.5%+5% NaCl）；750–770℃×2h 是去应力非敏化"),
}


def cmd_report():
    rows = read_csv(IMA_CSV)
    review = [r for r in rows if r["校正结论"] in ("存疑", "不符")]
    dai_pdfs = set(x for x in os.listdir(DAI) if x.lower().endswith(".pdf"))
    n_doubt = sum(1 for r in review if r["校正结论"] == "存疑")
    n_bad = sum(1 for r in review if r["校正结论"] == "不符")
    in_place = sum(1 for r in review if r["磁盘文件名"] in dai_pdfs)
    in_doubt = sum(1 for r in review
                   if r["校正结论"] == "存疑" and r["磁盘文件名"] in dai_pdfs)
    in_bad = sum(1 for r in review
                 if r["校正结论"] == "不符" and r["磁盘文件名"] in dai_pdfs)

    L = []
    L.append("# 需要你看的文献（待人工裁决，2026-11-08）")
    L.append("")
    L.append("时间：%s；依据 `ima_核查_校正判定_20261108.csv`、`待\\待_口径核查报告_20261108.md` "
             "与定稿 `口径基准_20261107.md` A1–A5。" % now())
    L.append("")
    L.append("## 一、背景")
    L.append("")
    L.append("第 10 批把 `准\\` 13 篇合规件录入 `0.xlsx` 后，剩下 `待\\` 中 **%d 篇 PDF**（%d 存疑 + %d 不符）"
             "都卡在「证据不足或口径边界」上，需要你拍板是否回原文补证、行级拆分或最终移出。"
             "另有 2 个非 PDF 核查材料（`待_口径核查报告_20261108.md`、`待_逐篇判定_20261108.csv`）保持原位。"
             % (len(review), n_doubt, n_bad))
    L.append("本报告范围＝这 %d 篇需人工裁决件；不含已入库的 `文献\\` 与 `0.xlsx` 数据行。" % len(review))
    L.append("")
    L.append("## 二、方法")
    L.append("")
    L.append("- 以 ima 两份逐篇证据为底，按定稿 A1（熔铸路线）、A2（介质）、A3（两步定义）、"
             "A4（增材/焊接主体排除但独立熔铸基准组可例外）、A5（奥氏体+少量δ 收、真双相硬排除）复核。")
    L.append("- 校正后分布：合规 13（已入第 10 批）/ 存疑 32 / 不符 31；唯一改判为 #28 存疑→不符。")
    L.append("- 4 处文件名截断已按磁盘真名对齐（#7/#39/#47/#58）。")
    L.append("")
    L.append("## 三、结论")
    L.append("")
    L.append("### 3.1 分布与在位核查")
    L.append("")
    L.append("| 校正结论 | 篇数 | 在 `待\\` 在位 |")
    L.append("| --- | --- | --- |")
    L.append("| 存疑 | %d | %d |" % (n_doubt, in_doubt))
    L.append("| 不符 | %d | %d |" % (n_bad, in_bad))
    L.append("| 合计 | %d | %d / %d |" % (len(review), in_place, len(review)))
    L.append("")
    L.append("### 3.2 六条最需要你先看（争议/差异件）")
    L.append("")
    L.append("| # | 标题简称 | 校正结论 | 要你定什么 | 需补的证据 / 说明 |")
    L.append("| --- | --- | --- | --- | --- |")
    for r in review:
        no = str(r["序号"])
        if no in DISPUTED:
            q, how = DISPUTED[no]
            L.append("| %s | %s | %s | %s | %s |"
                     % (no, r["标题简称"], r["校正结论"], q, how.replace("|", "/")))
    L.append("")
    L.append("### 3.3 %d 篇逐篇（要你看什么 / 主要理由）" % len(review))
    L.append("")
    L.append("| # | 校正结论 | 标题简称 | 主要理由 / 待办 | 差异说明 |")
    L.append("| --- | --- | --- | --- | --- |")
    for r in review:
        reason = (r.get("ima主要理由") or "").replace("|", "/").replace("\n", " ")
        diff = (r.get("差异说明") or "").replace("|", "/").replace("\n", " ")
        if len(diff) > 90:
            diff = diff[:90] + "…"
        L.append("| %s | %s | %s | %s | %s |"
                 % (r["序号"], r["校正结论"], r["标题简称"], reason, diff))
    L.append("")
    L.append("## 四、建议")
    L.append("")
    L.append("1. **优先裁决 6 条争议件**：#16 补 Table 2、#15/#33 定是否允许行级拆分；"
             "#26/#73/#28 只需确认理由校正，不必改结论。")
    L.append("2. **31 篇不符**：若你确认，按既有惯例移入 `待\\_不符合\\`（当前目录不存在，需先建）；"
             "本轮未授权，未动。")
    L.append("3. **32 篇存疑**：多为 Epit 仅图示、人工海水/浓度序列、缺参比、缓蚀剂、含双相行；"
             "补证后能转合规的再另起一批编号入库（第 10 批已占 147–159，后续从 160 起）。")
    L.append("4. `1.xlsx` / `2.xlsx`、两维工艺分组不在本轮范围；`0.xlsx` 的 Y/Z 两列按你此前指示保留。")
    with open(NEEDS_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print("[report] 需看 %d 篇（存疑 %d / 不符 %d / 在位 %d）/ 6 条争议 / 报告 %s"
          % (len(review), n_doubt, n_bad, in_place, NEEDS_REPORT))


# ---------------------------------------------------------------- finalize
def cmd_finalize():
    # gate
    problems = []
    if not os.path.exists(BACKUP):
        problems.append("缺底本")
    else:
        if md5(TARGET) != md5(BACKUP):
            problems.append("0.xlsx 已被改过（与底本 MD5 不同），停止")
    cards = ["%03d.json" % n for n in range(FIRST_NO, LAST_NO + 1)]
    have = [x for x in cards if os.path.exists(os.path.join(CARD_DIR, x))]
    if len(have) != N_PAPERS:
        problems.append("卡片 %d/%d 不齐，缺 %r" % (len(have), N_PAPERS,
                        [x for x in cards if x not in have]))
    else:
        for x in cards:
            try:
                with open(os.path.join(CARD_DIR, x), encoding="utf-8") as f:
                    c = json.load(f)
                if c.get("flag") or not isinstance(c.get("rows"), list) or len(c["rows"]) < 1:
                    problems.append("%s flag/零行" % x)
            except Exception as e:
                problems.append("%s JSON 错误 %s" % (x, e))
    src = [x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf")]
    if len(src) != N_PAPERS:
        problems.append("准\\ PDF %d 篇（应为 %d）" % (len(src), N_PAPERS))
    elif any(not pdf_ok(os.path.join(ZHUN, x)) for x in src):
        problems.append("准\\ 有坏件")
    exist = ["%02d.pdf" % n for n in range(FIRST_NO, LAST_NO + 1)
             if os.path.exists(os.path.join(LIT, "%02d.pdf" % n))]
    if exist:
        problems.append("文献\\147-159 目标已存在：%r" % exist)
    if problems:
        print("[finalize] 前置门未过：")
        for x in problems:
            print("   -", x)
        raise SystemExit(2)
    print("[finalize] 前置门通过：%d 卡齐 / 13 源 PDF 全好 / 147-159 目标全空 / 0.xlsx 与底本一致" % N_PAPERS)
    steps = ["check", "build", "commit", "move", "verify", "report"]
    for s in steps:
        print("\n===== finalize -> %s =====" % s, flush=True)
        subprocess.run([sys.executable, "-X", "utf8", os.path.abspath(__file__), s], check=True)
    print("\n[finalize] 全部 %d 步完成" % len(steps))


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "prep"
    cmds = {"prep": cmd_prep, "check": cmd_check, "build": cmd_build, "commit": cmd_commit,
            "move": cmd_move, "verify": cmd_verify, "report": cmd_report, "finalize": cmd_finalize}
    if mode not in cmds:
        raise SystemExit("未知子命令：%s（应为 %s）" % (mode, "/".join(cmds)))
    cmds[mode]()


if __name__ == "__main__":
    main()
