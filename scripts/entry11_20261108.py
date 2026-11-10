# -*- coding: utf-8 -*-
r"""第 11 批入库（准\ 当前 18 篇 -> 0.xlsx，文献编号 160-177）确定性流水线。

子命令：
  prep       备份 0.xlsx + 生成 18 篇取证输入（text/digest/RULES/manifest/numbering）
  figs       为 14 篇需读图文献预渲染候选图页（temp\entry11_20261108\figs，≤28 PNG）
  check      质检 18 张卡片 -> final_rows.json + 质检 CSV；不改 xlsx
  build      以 备份\0_基准_第11批入库前_20261108.xlsx 为底生成校验件并逐格校验
  commit     校验通过且 0.xlsx 未改时写回 0.xlsx 并二次复验
  move       18 篇从 准\ 移动重编号到 文献\160.pdf-177.pdf + 台账/对照报告
  verify     闭环校验 0.xlsx <-> 文献\01-177 <-> 台账源 MD5
  reading    汇总视觉试读结果，出「第 11 批读图说明」
  finalize   gate 通过后按序 check/build/commit/move/verify/reading

复用 scripts\build_entry9_20261108.py 的 load_cards/build/verify/cellval/FIELDS/md5
（importlib 载入后 monkeypatch 常量，不改历史脚本）。
"""
import csv
import glob
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
WORK = os.path.join(ROOT, "备份", "entry11_20261108")
TEXT_DIR = os.path.join(WORK, "text")
DIG_DIR = os.path.join(WORK, "digest")
CARD_DIR = os.path.join(WORK, "cards")
VISION_PROMPT_DIR = os.path.join(WORK, "vision_prompts")
VISION_DIR = os.path.join(WORK, "vision")
BACKUP = os.path.join(ROOT, "备份", "0_基准_第11批入库前_20261108.xlsx")
TARGET = os.path.join(ROOT, "0.xlsx")
STAGE = os.path.join(WORK, "0_追加校验11.xlsx")
FINAL = os.path.join(WORK, "final_rows.json")
QC = os.path.join(WORK, "cards_质检_第11批_20261108.csv")
MANIFEST = os.path.join(WORK, "manifest.json")
NUMBERING = os.path.join(WORK, "numbering.csv")
RULES11 = os.path.join(WORK, "RULES.md")
EXAMPLES = os.path.join(WORK, "examples.json")
FIGURES_JSON = os.path.join(WORK, "figures.json")
VISION_RESULTS = os.path.join(WORK, "vision_results.json")
DOUBT_LEDGER = os.path.join(ROOT, "准_存疑批处理台账_20261108.csv")

ZHUN = os.path.join(ROOT, "准")
LIT = os.path.join(ROOT, "文献")
DAI = os.path.join(ROOT, "待")
LEDGER = os.path.join(ROOT, "准_入0与重编号对照_第11批_20261108.csv")
ENTRY_REPORT = os.path.join(ROOT, "准文献入库对照_第11批_20261108.md")
VERIFY_REPORT = os.path.join(ROOT, "0_第11批入库校验_20261108.md")
READING_REPORT = os.path.join(ROOT, "第11批读图说明_20261108.md")
TEXT_ALL = os.path.join(ROOT, "备份", "text_all.jsonl")
FIG_DIR = os.path.join(ROOT, "temp", "entry11_20261108", "figs")
TEMP_DIR = os.path.join(ROOT, "temp")

FIRST_NO, LAST_NO = 160, 177
N_PAPERS = LAST_NO - FIRST_NO + 1
NCOL = 26
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]
FIGURE_SUFFIX = "__有图待读.pdf"
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


def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def strip_marker(name):
    if name.endswith(FIGURE_SUFFIX):
        return name[:-len(FIGURE_SUFFIX)] + ".pdf"
    return name


def count_temp():
    files = 0
    total = 0
    for base, _dirs, names in os.walk(TEMP_DIR):
        for n in names:
            p = os.path.join(base, n)
            if os.path.isfile(p):
                files += 1
                total += os.path.getsize(p)
    return files, total


def index_text_all():
    by_exact, by_norm = {}, {}
    with open(TEXT_ALL, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            if rec.get("tag") != "DAI":
                continue
            fn = str(rec.get("file") or "")
            if not fn:
                continue
            by_exact.setdefault(fn, rec)
            by_norm.setdefault(strip_marker(fn), rec)
    return by_exact, by_norm


def get_rec(filename, by_exact, by_norm):
    base = strip_marker(filename)
    return (by_exact.get(filename) or by_norm.get(filename)
            or by_exact.get(base) or by_norm.get(base))


def reason_by_file():
    out = {}
    if not os.path.exists(DOUBT_LEDGER):
        return out
    for r in read_csv(DOUBT_LEDGER):
        path = str(r.get("现名或目标路径") or "")
        if path.startswith(r"准\\") or path.startswith("准\\") or path.startswith("准/"):
            name = os.path.basename(path.replace("\\", os.sep).replace("/", os.sep))
            out[name] = "；".join(str(r.get(k) or "") for k in ("分组", "动作", "ima主要理由")).strip("；")
    return out


# ---------------------------------------------------------------- prep
def cmd_prep():
    for d in (WORK, TEXT_DIR, DIG_DIR, CARD_DIR, VISION_PROMPT_DIR, VISION_DIR):
        os.makedirs(d, exist_ok=True)

    if not os.path.exists(BACKUP):
        shutil.copy2(TARGET, BACKUP)
        print("[prep] 已备份 0.xlsx ->", BACKUP, "MD5=", md5(BACKUP))
    else:
        same = md5(TARGET) == md5(BACKUP)
        print("[prep] 底本已存在，校验一致性：", same)
        if not same:
            raise SystemExit("0.xlsx 与既有备份不一致，停止（先人工核对）")

    by_exact, by_norm = index_text_all()
    reasons = reason_by_file()
    files = sorted(x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf"))
    if len(files) != N_PAPERS:
        raise SystemExit(f"准\\ PDF 数应为 {N_PAPERS}，实测 {len(files)}：{files}")
    targets = [os.path.join(LIT, f"{n:02d}.pdf") for n in range(FIRST_NO, LAST_NO + 1)]
    if any(os.path.exists(p) for p in targets):
        raise SystemExit("文献\\160.pdf-177.pdf 有目标已存在，停止")
    recs = [(fn, get_rec(fn, by_exact, by_norm)) for fn in files]
    missing = [fn for fn, rec in recs if rec is None]
    if missing:
        raise SystemExit(f"全文索引缺 {len(missing)} 篇：{missing}")
    bad = [fn for fn in files if not pdf_ok(os.path.join(ZHUN, fn))]
    if bad:
        raise SystemExit(f"准\\ 有坏件：{bad}")

    wb = openpyxl.load_workbook(BACKUP, data_only=True)
    ws = wb["Sheet1"]
    ex = []
    for rr in range(3, min(ws.max_row, 30) + 1):
        if isinstance(ws.cell(rr, 1).value, (int, float)) and len(ex) < 3:
            ex.append({"文献编号": ws.cell(rr, 1).value,
                       "样本": ws.cell(rr, 2).value,
                       "文献信息": ws.cell(rr, 25).value,
                       "备注": ws.cell(rr, 26).value})
    save_json(EXAMPLES, ex)

    # 规则：复用第9批，改路径/标题 + 追加第11批补充（优先于上文）
    base_rules = os.path.join(ROOT, "备份", "entry_20261108", "RULES.md")
    with open(base_rules, encoding="utf-8") as f:
        rules = f.read()
    rules = rules.replace("第9批入库", "第11批入库")
    rules = rules.replace(r"备份\entry_20261108", r"备份\entry11_20261108")
    rules = re.sub(r"完成后\*\*只回一行\*\* `ok:<rows数>` 或 `flag:<原因>`；",
                   "完成后**只回一行** `ok:<rows数>`；", rules)
    rules = rules.replace('，可选 `flag`(str)。', "；**禁止** `flag` 字段。")
    rules = rules.replace('，顶层加 `"flag": "原因"`。', "，**不要加 flag**（见文末第11批补充）。")
    rules = rules.replace("`AA`（口径标注）：不要出现在卡片里，脚本统一留空",
                          "本表为 **26 列（A–Z）**，没有 `AA` 列")
    rules += (
        "\n\n## 第11批补充（优先于上文）\n"
        "- 新编号范围 **160–177**；顶层 `newno` 必须等于任务给出的编号，`file` 必须等于 digest 题头【原文件】完整文件名。\n"
        "- 顶层 **禁止 `flag`**；无论能否找到数值，都要出至少 1 行卡片；找不到可录数值写「未提及」或 `N/A（待读图）`，在 `notes` 说明原因。\n"
        "- 14 篇文件名含 `__有图待读.pdf`：只录「不依赖读图也能确定」的原文初始信息；Epit 若只有曲线/图、文本层无绝对值，一律 `N/A（待读图）`。要从图里才能确定的样本名、数值、曲线条件也不要猜；写 `未提及` 或 `N/A（待读图）`。\n"
        "- 4 篇非读图件：#3/#56 海水候选按「海水可先进去、3.56 不行」收录，notes 标海水特殊性；#15/#33 按行级拆分只取奥氏体数据行，不取真双相/铁素体行。\n"
        "- 严禁渲染页面 / 图像数字化 / 写临时 png 或 py / 联网；读图尝试由主流程单独完成，结果只进说明，不回填 `0.xlsx`。\n"
        "- `0.xlsx` 只记原文初始信息，不做 SCE 换算、不补 0、不反算 Fe、不写对 1/2.xlsx 的指示；RE 不写 `→SCE` / `0.1989 V vs SHE`。\n"
        "- `notes` 短写（尽量 <80 字），只写数据来源与特殊情况；需读图件必须写明「Epit 待读图：Fig.X」或「Epit 仅图示」。\n"
        "- 完成后**只回一行** `ok:<rows数>`；卡片必须用 `write` 工具写到 `备份\\entry11_20261108\\cards\\NNN.json`。\n"
    )
    with open(RULES11, "w", encoding="utf-8") as f:
        f.write(rules)

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
        r"|铸造|熔炼|感应炉|连铸|铸锭|砂型|离心|精密铸造|待读图|图示"
    )
    manifest, numbering = [], []
    dig_sizes = []
    for i, fn in enumerate(files):
        no = FIRST_NO + i
        idx = f"{no:03d}"
        rec = recs[i][1]
        text = str(rec.get("text") or "")
        doi = str(rec.get("doi") or "")
        idx_name = str(rec.get("file") or "")
        needs_fig = fn.endswith(FIGURE_SUFFIX)
        reason = reasons.get(fn, "第11批：从准\\ 入库" + ("；需读图" if needs_fig else "；非读图"))
        hdr = (
            f"【新编号】{no}\n【原文件】{fn}\n【全文索引名】{idx_name}\n"
            f"【需读图】{'是' if needs_fig else '否'}\n【批次来源】{reason}\n"
            f"【DOI】{doi}\n【全文字数】{len(text)}\n"
        )
        if len(text) > 90000:
            body = text[:60000] + f"\n……（中段略，如需请再读 text/{idx}.txt）……\n" + text[-28000:]
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
        dg = (
            hdr + "【紧凑 digest：开头 4500 + 关键词命中段 ≤26000 + 结尾 2500 字符】\n"
            f"【如需完整表格或数值，请再读 text/{idx}.txt（read 的 offset/limit 分段）】\n\n"
            f"===== 标题与摘要（开头） =====\n{head}\n\n===== 关键证据段 =====\n"
            f"{chr(10).join(['---'] + mid) if mid else ''}\n\n===== 原文结尾 =====\n{tail}\n"
        )
        with open(os.path.join(DIG_DIR, idx + ".txt"), "w", encoding="utf-8") as f:
            f.write(dg)
        dig_sizes.append(len(dg))

        manifest.append({"idx": idx, "newno": no, "file": fn,
                         "norm_file": idx_name, "needs_fig": needs_fig,
                         "verdict": "需读图" if needs_fig else "非读图",
                         "chars": len(text), "doi": doi,
                         "text": os.path.join("text", idx + ".txt"),
                         "reason": reason})
        numbering.append({"新编号": no, "原文件名": fn, "是否需读图": "是" if needs_fig else "否",
                          "全文索引名": idx_name, "全文字数": len(text),
                          "目标文献编号": f"文献\\{no:02d}.pdf", "读图原因": reason})
    save_json(MANIFEST, manifest)
    with open(NUMBERING, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(numbering[0].keys()))
        w.writeheader()
        w.writerows(numbering)

    print(f"[prep] {N_PAPERS} 篇 text/digest 就绪，digest min/median/max="
          f"{min(dig_sizes)}/{sorted(dig_sizes)[len(dig_sizes)//2]}/{max(dig_sizes)} 字符，"
          f"合计 {sum(dig_sizes)/1024/1024:.2f} MB")
    print(f"[prep] manifest / numbering / RULES.md 已写；底本 dims={ws.max_row} 行 / {ws.max_column} 列")
    print(f"[prep] 编号 {FIRST_NO}-{LAST_NO}：")
    for m in manifest:
        print("   ", m["idx"], "需读图" if m["needs_fig"] else "非读图", m["file"])


# ---------------------------------------------------------------- figs
def page_score(page):
    try:
        txt = page.get_text("text") or ""
    except Exception:
        txt = ""
    try:
        imgs = len(page.get_images(full=True))
    except Exception:
        imgs = 0
    try:
        draws = len(page.get_drawings())
    except Exception:
        draws = 0
    score = 0
    if re.search(r"Fig(?:ure)?\.?\s*\d|图\s*\d", txt, re.I):
        score += 3
    if re.search(r"polarization|potentiodynamic|pitting|breakdown|Epit|Eb\b|current density|OCP"
                 r"|点蚀|孔蚀|击穿|极化|阳极", txt, re.I):
        score += 3
    if re.search(r"SCE|Ag/AgCl|calomel|甘汞|参比", txt, re.I):
        score += 1
    if re.search(r"3\.5\s*(wt|mass|%)|0\.6\s*(mol|M)|NaCl|氯化钠|chloride", txt, re.I):
        score += 1
    if imgs:
        score += 2
    if draws > 20:
        score += 1
    return score, txt, imgs, draws


def cmd_figs():
    os.makedirs(FIG_DIR, exist_ok=True)
    os.makedirs(VISION_PROMPT_DIR, exist_ok=True)
    manifest = load_json(MANIFEST, [])
    needs = [m for m in manifest if m.get("needs_fig")]
    if len(needs) != 14:
        raise SystemExit(f"需读图篇数应为 14，manifest 实测 {len(needs)}")
    existing = glob.glob(os.path.join(FIG_DIR, "*.png"))
    if len(existing) > 28:
        raise SystemExit(f"图页缓存已有 {len(existing)} 张 PNG（>28），停止以免 temp 超配额")

    try:
        import pymupdf
    except Exception as e:
        raise SystemExit(f"缺 pymupdf：{e}")
    try:
        from rapidocr_onnxruntime import RapidOCR
        ocr = RapidOCR()
    except Exception as e:
        print("[figs] RapidOCR 不可用，OCR 文本留空：", e)
        ocr = None

    fig_items, prompts = [], []
    rendered = 0
    for m in needs:
        no = int(m["newno"])
        idx = m["idx"]
        src = os.path.join(ZHUN, m["file"])
        if not os.path.exists(src):
            raise SystemExit(f"缺源 PDF（应在 move 前跑 figs）：{src}")
        doc = pymupdf.open(src)
        scores = []
        for pno in range(1, len(doc) + 1):
            page = doc[pno - 1]
            s, txt, imgs, draws = page_score(page)
            scores.append({"page": pno, "score": s, "caption": txt[:600].replace("\n", " "),
                           "imgs": imgs, "draws": draws})
        if not scores:
            raise SystemExit(f"{idx} PDF 无页：{src}")
        picked = sorted([x for x in scores if x["score"] > 0],
                        key=lambda x: (-x["score"], x["page"]))[:2]
        if not picked:
            picked = sorted(scores, key=lambda x: (-x["imgs"], -x["draws"], x["page"]))[:2]
        picked = sorted(picked, key=lambda x: x["page"])
        pages = []
        prompt_lines = [
            f"新编号：{no}",
            f"原文件：{m['file']}",
            f"目标文献：文献\\{no:02d}.pdf（移动后）",
            f"需读图原因：{m.get('reason','')}",
            "任务：尝试从下列预渲染图页读取点蚀/击穿电位 Epit 或 Eb 的可读线索。先用 read_image 工具打开每张 PNG；"
            "若图像工具不可用，只能用下面 OCR/题注文字判断，并在 uncertainty 写明未见原图。",
            "严禁猜测；试读值只进说明，不回填 0.xlsx；不改任何文件、不渲染页面、不写 png/py、不联网。",
            "输出结构化结果：newno, readable, epit_estimates, evidence, uncertainty, reason, manual_action。",
        ]
        for p in picked:
            pno = p["page"]
            png = os.path.abspath(os.path.join(FIG_DIR, f"{idx}_p{pno:02d}.png"))
            if not os.path.exists(png):
                pix = doc[pno - 1].get_pixmap(dpi=130, alpha=False)
                pix.save(png)
                rendered += 1
            ocr_text = ""
            if ocr is not None:
                try:
                    res, _ = ocr(png)
                    ocr_text = "\n".join(str(r[1]) for r in (res or []))[:1200]
                except Exception as e:
                    ocr_text = f"[OCR失败：{e}]"
            item = {"page": pno, "png": png, "score": p["score"],
                    "caption": p["caption"], "ocr_text": ocr_text}
            pages.append(item)
            prompt_lines.append(
                f"\n--- 图页 p{pno} ---\nPNG：{png}\nscore={p['score']} imgs={p['imgs']} draws={p['draws']}\n"
                f"题注/页文本：{p['caption']}\nOCR：{ocr_text[:800]}"
            )
        doc.close()
        fig_items.append({"newno": no, "file": m["file"], "page_count": len(scores), "pages": pages})
        with open(os.path.join(VISION_PROMPT_DIR, idx + ".txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(prompt_lines) + "\n")
        prompts.append({"idx": idx, "newno": no, "prompt": os.path.join(VISION_PROMPT_DIR, idx + ".txt")})

    save_json(FIGURES_JSON, {"generated": now(), "dpi": 130, "items": fig_items,
                             "vision_prompts": prompts})
    pngs = glob.glob(os.path.join(FIG_DIR, "*.png"))
    temp_files, temp_bytes = count_temp()
    print(f"[figs] 已渲染 {rendered} 张新 PNG / 共 {len(pngs)} 张；14 篇 prompts 就绪；"
          f"temp {temp_files} 文件 / {temp_bytes/1024/1024:.2f} MB")
    if len(pngs) > 28 or temp_files > 300 or temp_bytes > 300 * 1024 * 1024:
        raise SystemExit("图页缓存超出 ≤28 PNG / temp≤300文件/300MB 约束，停止")
    print("[figs] figures.json", FIGURES_JSON)


# ---------------------------------------------------------------- build 模块（monkeypatch）
def patched():
    spec = importlib.util.spec_from_file_location(
        "build9", os.path.join(ROOT, "scripts", "build_entry9_20261108.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    base = openpyxl.load_workbook(BACKUP)["Sheet1"]
    base_rows = base.max_row
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


def load_final(_m=None):
    if not os.path.exists(FINAL):
        raise SystemExit("缺 final_rows.json，先跑 check")
    data = load_json(FINAL, [])
    if [p["no"] for p in data] != list(range(FIRST_NO, LAST_NO + 1)):
        raise SystemExit("final_rows.json 编号不连续（疑似质检未通过时留下的旧文件），重跑 check")
    bad = [p["no"] for p in data
           if p.get("flag") or not p.get("citation") or not p.get("notes")
           or not isinstance(p.get("rows"), list) or len(p["rows"]) < 1
           or any(any(k not in r or r[k] is None or str(r[k]).strip() == "" for k in FIELDS)
                  for r in p["rows"])]
    if bad:
        raise SystemExit(f"final_rows.json 有 {len(bad)} 篇不合格（{bad}），重跑 check 并补齐卡片")
    return data


# ---------------------------------------------------------------- check
def cmd_check():
    m, base_rows, anchor = patched()
    data, qc = m.load_cards()
    status = {}
    for x in qc:
        status[x["状态"]] = status.get(x["状态"], 0) + 1
    total = sum(len(p["rows"]) for p in data)
    print(f"[check] 卡片 {len(data)} 篇 / 数据 {total} 条 / 编号 {FIRST_NO}-{LAST_NO} / "
          f"底本 {base_rows} 行 / 模板行 {anchor}")
    print("[check] 状态分布：", status)
    # 第11批语义闸：禁换算说明、禁工作日志、短备注；读图件的 N/A（待读图）合法。
    forbidden = ["→SCE", "0.1989", "换算至 SCE", "清洗时", "1.xlsx", "2.xlsx",
                 "【本轮入库", "【提取方式", "口径推演", "介质复述"]
    semantic = []
    for p in data:
        blob = "\n".join([str(p.get("citation") or ""), str(p.get("notes") or ""),
                          *[str(v) for r in p.get("rows", []) for v in r.values()]])
        hits = [w for w in forbidden if w in blob]
        if hits:
            semantic.append((p["no"], "含不该进 0.xlsx 的措辞：" + ",".join(hits)))
        if len(str(p.get("notes") or "")) > 120:
            semantic.append((p["no"], "备注>120字"))
    for no, msg in semantic:
        print("[check] 语义提醒：", no, msg)
    save_json(FINAL, data)
    blocking = [x for x in qc if not x["状态"].startswith("ok")] + [x for x in qc if x["行数"] < 1]
    print("[check] 质检表：", QC)
    if blocking or semantic:
        print(f"[check] 质检未通过：{len(blocking)} 篇卡片问题 / {len(semantic)} 篇语义问题")
        for x in blocking[:15]:
            print("   -", x["新编号"], str(x["原文件名"])[:45], "|", x["状态"], "|", x["问题"][:140])
        raise SystemExit(2)
    print(f"[check] 质检通过：{len(data)} 篇 / {total} 条数据，全部可入库；已写 final_rows.json")


# ---------------------------------------------------------------- build / commit
def cmd_build():
    m, base_rows, anchor = patched()
    data = load_final(m)
    if md5(TARGET) != md5(BACKUP):
        raise SystemExit("0.xlsx 与第11批入库前备份不一致，停止（先人工核对）")
    total = sum(len(p["rows"]) for p in data)
    blocks = m.build(data)
    problems = m.verify(data, blocks, STAGE)
    print(f"[build] 数据 {len(data)} 篇 / {total} 条，写入行 {base_rows+1}-{base_rows+total}；校验件 {STAGE}")
    if problems:
        print(f"[build] 校验未通过，共 {len(problems)} 项：")
        for x in problems[:40]:
            print("   -", x)
        raise SystemExit(2)
    _nm = len(openpyxl.load_workbook(STAGE)["Sheet1"].merged_cells.ranges)
    _om = len(openpyxl.load_workbook(BACKUP)["Sheet1"].merged_cells.ranges)
    print(f"[build] 校验通过：前 {base_rows} 行零差异 / 原 {_om} 合并区全保留 / 新增 {_nm-_om} / "
          f"列宽缩放不变 / {total} 行逐格回读一致")
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
    print(f"[commit] 已写回 0.xlsx，MD5={md5(TARGET)}；{len(data)} 篇 / {total} 条；写回后复验通过")


# ---------------------------------------------------------------- move
def cmd_move():
    m, base_rows, anchor = patched()
    data = load_final(m)
    manifest = {int(x["newno"]): x for x in load_json(MANIFEST, [])}
    before_lit = len([x for x in os.listdir(LIT) if x.lower().endswith(".pdf")])
    items, problems = [], []
    for p in data:
        no = p["no"]
        mf = manifest.get(no, {})
        src = os.path.join(ZHUN, p["file"])
        dst = os.path.join(LIT, f"{no:02d}.pdf")
        rec = {"新编号": no, "原文件名": p["file"],
               "需读图": "是" if mf.get("needs_fig") else "否",
               "读图原因": str(mf.get("reason") or ""),
               "题录": p["citation"], "备注": p["notes"],
               "数据行数": len(p["rows"]),
               "样本名": "; ".join(str(r["sample"]) for r in p["rows"]),
               "源字节": os.path.getsize(src) if os.path.exists(src) else 0,
               "源MD5": md5(src) if os.path.exists(src) else "",
               "动作": "", "目标": f"{no:02d}.pdf"}
        if not os.path.exists(src):
            problems.append(f"{no} 缺源：{p['file']}")
            rec["动作"] = "缺源，未移动"
        elif not pdf_ok(src):
            problems.append(f"{no} 坏件：{p['file']}")
            rec["动作"] = "坏件，未移动"
        elif os.path.exists(dst):
            if md5(src) == md5(dst):
                problems.append(f"{no} 目标已存在同 MD5（用户未授权删源/覆盖）：{p['file']}")
                rec["动作"] = "冲突保留源（目标已存在同MD5）"
            else:
                problems.append(f"{no} 目标已存在且内容不同：{p['file']}")
                rec["动作"] = "冲突，未覆盖"
        else:
            src_md5 = md5(src)
            shutil.move(src, dst)
            if not (os.path.exists(dst) and not os.path.exists(src)
                    and md5(dst) == src_md5 and pdf_ok(dst)):
                problems.append(f"{no} 移动后校验失败：{p['file']}")
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
    need_nos = [int(x["newno"]) for x in load_json(MANIFEST, []) if x.get("needs_fig")]

    L = [
        "# 准文献入库对照（第 11 批，2026-11-08）", "",
        f"状态：{now()}；用户 m02877「再次启动准中的录入工作」、m02892「多启动一些子智能体」、m02902「可以，执行。」；按前批惯例「移动」重编号。", "",
        "## 一、背景", "",
        "`准\\` 中 18 篇为上一轮存疑批处理后留下的文献：14 篇文件名带 `__有图待读.pdf`，4 篇为非读图件（#3/#56 海水候选、#15/#33 行级拆分）。"
        "本批把 18 篇的原文初始信息录入 `0.xlsx`（编号 160–177），并把 PDF 从 `准\\` **移动**重编号为 "
        "`文献\\160.pdf`–`177.pdf`，使「编号 → `文献\\NN.pdf`」溯源链继续成立。", "",
        "## 二、方法", "",
        "- 4 篇非读图件：抽合规数据行；14 篇读图件：抽「不依赖读图也能确定」的 23 字段，Epit 仅图示时记 `N/A（待读图）`。",
        "- 逐篇并行抽取 23 键卡片，只记原文初始信息（不做 SCE 换算、备注短写、26 列无 `AA`）。",
        f"- 以 `备份\\0_基准_第11批入库前_20261108.xlsx` 为底生成校验件，前 {base_rows} 行值与样式零差异、160–177 连续、合并区/列宽一致全过才写回。",
        "- `shutil.move` 逐篇校验 `%PDF-` 头与 MD5；目标已存在则保留源、不覆盖、不删除。", "",
        "## 三、结论", "",
        "| 项 | 数值 |", "| --- | --- |",
        f"| 文献篇数 | {len(items)} |",
        f"| 数据条数 | {sum(r['数据行数'] for r in items)} |",
        f"| 成功移动 | {moved} |",
        f"| 冲突/异常 | {len(problems)} |",
        f"| `文献\\` PDF | {before_lit} → {after_lit} |",
        f"| `准\\` 剩余 PDF | {remain} |",
        f"| `0.xlsx` | {base_rows + sum(r['数据行数'] for r in items)} 行 / 26 列（写回后） |",
        f"| 需读图编号 | {', '.join(str(x) for x in need_nos)} |", "",
        "逐篇明细（题录 / 备注 / 数据行数 / 源 MD5 / 动作 / 是否需读图）见 "
        "`准_入0与重编号对照_第11批_20261108.csv`；读图试读结果见 `第11批读图说明_20261108.md`。", "",
        "| 编号 | 原文件名 | 需读图 | 行数 | 动作 |", "| --- | --- | --- | --- | --- |",
    ]
    for r in items:
        L.append(f"| {r['新编号']} | {r['原文件名']} | {r['需读图']} | {r['数据行数']} | {r['动作']} |")
    L += ["", "## 四、建议", "",
          "- 14 篇 `__有图待读` 文件已按原名移动，未去掉后缀；读图结果只写说明，未回填 `0.xlsx`。",
          "- 人工读图后若要回填 Epit，属新指令；建议另起一小批修 0.xlsx 并留备份。",
          "- 本批不改 `1.xlsx` / `2.xlsx`；如需把 160–177 纳入清洗/建模，属独立任务。",
          "- 回滚：底本 `备份\\0_基准_第11批入库前_20261108.xlsx` 与移动台账均可还原。"]
    with open(ENTRY_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")

    print(f"[move] 文献 {before_lit}→{after_lit} / 准剩余 {remain} / 已移动 {moved} / 冲突异常 {len(problems)}")
    print("[move] 台账", LEDGER, "/ 报告", ENTRY_REPORT)
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
        problems.append(f"0.xlsx 结构不对：{cur.max_row} 行 / {cur.max_column} 列")
    changed = 0
    for rr in range(1, base_rows + 1):
        for c in range(1, NCOL + 1):
            if old.cell(rr, c).value != cur.cell(rr, c).value:
                changed += 1
    if changed:
        problems.append(f"前 {base_rows} 行有 {changed} 格被改")
    r = base_rows
    for p in data:
        start = r + 1
        if cur.cell(start, 1).value != p["no"]:
            problems.append(f"编号 {p['no']} 块首 A 列不符")
        if cur.cell(start, 25).value != p["citation"] or cur.cell(start, 26).value != p["notes"]:
            problems.append(f"编号 {p['no']} 题录/备注不符")
        for i, row in enumerate(p["rows"]):
            rr = start + i
            for j, key in enumerate(FIELDS):
                if cur.cell(rr, 2 + j).value != row[key]:
                    problems.append(f"编号 {p['no']} {key} 字段 r{rr} 不符")
                    break
        r += len(p["rows"])
    lit_pdfs = set(x for x in os.listdir(LIT) if x.lower().endswith(".pdf"))
    want = set(f"{i:02d}.pdf" for i in range(1, LAST_NO + 1))
    if lit_pdfs != want:
        problems.append(f"文献目录不连续：缺 {sorted(want-lit_pdfs)[:5]} / 多 {sorted(lit_pdfs-want)[:5]}")
    led = read_csv(LEDGER)
    for row in led:
        no = int(row["新编号"])
        dst = os.path.join(LIT, f"{no:02d}.pdf")
        if not os.path.exists(dst) or md5(dst) != row["源MD5"] or not pdf_ok(dst):
            problems.append(f"文献 {no} 与台账源 MD5 不一致")
    remain = len([x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf")])
    if remain:
        problems.append(f"准\\ 剩余 {remain} 篇 PDF（应为 0）")
    fig = load_json(FIGURES_JSON, {})
    vision = load_vision()
    n_fig = len((fig or {}).get("items", []))
    n_vision = len(flatten_vision(vision))
    if n_fig != 14:
        problems.append(f"figures.json 需读图篇数 {n_fig}（应为 14）")
    if n_vision < 14:
        problems.append(f"vision_results.json 返回 {n_vision} 条（应≥14）")
    L = [
        "# 0.xlsx 第 11 批入库校验（2026-11-08）", "", f"时间：{now()}。", "",
        "## 一、背景", "", "第 11 批把 `准\\` 18 篇录入 `0.xlsx`（编号 160–177）并移动重编号 PDF，另对 14 篇需读图文献做图页试读。", "",
        "## 二、方法", "", f"只读比对：底本 vs 现表前 {base_rows} 行逐格、新增块逐字段回读、文献目录 01–{LAST_NO} 连续、"
        "新增 PDF 与移动台账源 MD5 一致、图页/视觉结果齐备。", "",
        "## 三、结论", "", "| 项 | 数值 |", "| --- | --- |",
        f"| 现表 | {cur.max_row} 行 / {cur.max_column} 列 |",
        f"| 新增 | {len(data)} 篇 / {total} 条 / 编号 {FIRST_NO}–{LAST_NO} |",
        f"| 文献目录 | {len(lit_pdfs)} 篇 |",
        f"| 图页候选 | {n_fig} 篇 |",
        f"| 视觉试读结果 | {n_vision} 条 |",
        f"| `准\\` 剩余 | {remain} 篇 |",
        f"| 问题项 | {len(problems)} |", "",
        "## 四、建议", "", "问题 0 项即可交付；如需回滚，底本在 `备份\\0_基准_第11批入库前_20261108.xlsx`。",
    ]
    if problems:
        L += ["", "问题明细：", *[f"- {x}" for x in problems[:30]]]
    with open(VERIFY_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print(f"[verify] 0.xlsx {cur.max_row} 行/{cur.max_column} 列 / 新增 {len(data)} 篇 {total} 条 / "
          f"文献 {len(lit_pdfs)} 篇 / 图页 {n_fig} 篇 / 视觉 {n_vision} 条 / 问题 {len(problems)} 项 / 报告 {VERIFY_REPORT}")
    for x in problems[:20]:
        print("   -", x)
    if problems:
        raise SystemExit(2)


# ---------------------------------------------------------------- 读图说明
def load_vision():
    out = []
    if os.path.isdir(VISION_DIR):
        for name in sorted(os.listdir(VISION_DIR)):
            if not name.lower().endswith(".json"):
                continue
            try:
                v = load_json(os.path.join(VISION_DIR, name), {})
            except Exception as e:
                print("[vision] 跳过坏 JSON", name, e)
                continue
            if isinstance(v, dict):
                if "newno" not in v and "no" not in v and "编号" not in v:
                    m = re.search(r"(\d{3})", name)
                    if m:
                        v["newno"] = int(m.group(1))
                v["_source"] = name
            out.append(v)
    if not out:
        out = load_json(VISION_RESULTS, []) or []
    return flatten_vision(out)


def flatten_vision(obj):
    if obj is None:
        return []
    if isinstance(obj, list):
        out = []
        for x in obj:
            out.extend(flatten_vision(x))
        return out
    if isinstance(obj, dict):
        if "items" in obj and isinstance(obj["items"], list):
            return flatten_vision(obj["items"])
        if any(k in obj for k in ("newno", "no", "编号", "epit_estimates", "readable")):
            return [obj]
        out = []
        for v in obj.values():
            if isinstance(v, (list, dict)):
                out.extend(flatten_vision(v))
        return out
    return []


def cmd_reading_report():
    manifest = load_json(MANIFEST, [])
    fig = load_json(FIGURES_JSON, {}) or {}
    fig_items = {int(x["newno"]): x for x in fig.get("items", [])}
    vision = load_vision()
    vmap = {}
    for v in vision:
        raw = v.get("newno", v.get("no", v.get("编号")))
        try:
            no = int(raw)
        except Exception:
            continue
        vmap.setdefault(no, v)
    needs = [m for m in manifest if m.get("needs_fig")]
    L = [
        "# 第 11 批需读图文献与图页试读说明（2026-11-08）", "",
        f"时间：{now()}；依据用户 m02877「需要读图的注明是哪些编号的文献，出一个说明，还有你先尝试读图，结果也在这个说明上」、"
        "m02892「多启动一些子智能体」、m02902「可以，执行。」。", "",
        "## 一、背景", "",
        f"第 11 批把 `准\\` 18 篇录入 `0.xlsx`（新编号 160–177）并移动为 `文献\\160.pdf`–`177.pdf`。"
        f"其中 **{len(needs)} 篇**此前列为「有图待读」：它们在 `0.xlsx` 中只录不依赖读图的原文初始信息，"
        "Epit 若只出现在曲线/图中，记 `N/A（待读图）`。本说明列明需读图编号，并记录先做的图页试读结果。", "",
        "## 二、方法", "",
        "- PyMuPDF 按页文本/图对象评分，从每篇挑最多 2 个候选图页，130 dpi 预渲染为 PNG，缓存 `temp\\entry11_20261108\\figs\\`（≤28 张）。",
        "- 预渲染图页同时做 RapidOCR；14 个视觉判读子智能体分 7 段×2 篇无跨段 barrier 并发，基于 PNG、题注与 OCR 文字尝试读阳极拐点/点蚀电位。",
        "- 试读值只写入本说明，**不回填 `0.xlsx`**；不可读的写「不可读 + 原因」。图页试读属辅助，估读值不得当原文数值使用。", "",
        "## 三、结论", "",
        "### 3.1 需读图编号", "",
        "| 新编号 | 目标文献 | 原文件名 | 读图原因 | 候选图页 | 试读结果 | 依据 / 不确定性 | 人工补证建议 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for m in needs:
        no = int(m["newno"])
        fi = fig_items.get(no, {})
        pages = fi.get("pages", [])
        page_txt = ", ".join(f"p{p['page']}" for p in pages) or "未生成"
        v = vmap.get(no, {})
        readable = v.get("readable")
        epit = str(v.get("epit_estimates") or "").strip()
        reason = str(v.get("reason") or "").strip()
        evidence = str(v.get("evidence") or "").strip()
        uncertainty = str(v.get("uncertainty") or "").strip()
        manual = str(v.get("manual_action") or "打开 `文献\\%02d.pdf`，核对候选图页，按阳极拐点读 Epit 与参比，标 ±0.02 V 估读误差。" % no).strip()
        if not v:
            result = "未回收视觉结果"
            ev = "本轮子智能体未返回结构化结果；仍可人工读图。"
        elif readable is False or str(readable).lower() in ("false", "no", "不可读"):
            result = "不可读（已尝试）"
            ev = (reason or evidence or "图页/曲线刻度不足或参比不清") + ("；" + uncertainty if uncertainty else "")
        else:
            result = "试读值（未回填）：" + (epit or "未给出数值")
            ev = (evidence or "见 PNG 曲线与 OCR") + ("；" + uncertainty if uncertainty else "")
        L.append(f"| {no} | `文献\\{no:02d}.pdf` | `{m['file']}` | {str(m.get('reason','')).replace('|','/')} | "
                 f"{page_txt} | {result.replace('|','/')} | {ev.replace('|','/')[:180]} | {manual.replace('|','/')[:120]} |")
    L += ["", "### 3.2 试读值使用边界", "",
          "- 上表「试读值」仅为图页辅助判读，含 ±0.02 V 以上估读不确定性；没有写回 `0.xlsx`，也没有改 `1.xlsx`/`2.xlsx`。",
          "- 若人工确认某值可用，建议另起新指令：给出「编号 → 图号 → Epit/参比」对照，备份后只改 `0.xlsx` 对应 Epit 格并留痕。", "",
          "## 四、建议", "",
          "1. 优先人工复核 14 篇中「不可读 / 未回收视觉结果」的编号；可读编号也建议对照原 PDF 图页核一次。",
          "2. `0.xlsx` 中这些行的 Epit 仍以 `N/A（待读图）` 或原文文本数值为准；不要直接把本说明试读值当原文数据。",
          "3. 图页 PNG 仍在 `temp\\entry11_20261108\\figs\\`，temp 配额内；如需长期留存可另存 `备份\\`。",
          "4. 回滚：底本 `备份\\0_基准_第11批入库前_20261108.xlsx` 与 `准_入0与重编号对照_第11批_20261108.csv` 均可还原本批动作。"]
    with open(READING_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print(f"[reading] 需读图 {len(needs)} 篇 / 视觉结果 {len(vision)} 条 / 报告 {READING_REPORT}")


# ---------------------------------------------------------------- finalize
def cmd_finalize():
    problems = []
    if not os.path.exists(BACKUP):
        problems.append("缺底本")
    elif md5(TARGET) != md5(BACKUP):
        problems.append("0.xlsx 已被改过（与底本 MD5 不同），停止")
    cards = [f"{n:03d}.json" for n in range(FIRST_NO, LAST_NO + 1)]
    have = [x for x in cards if os.path.exists(os.path.join(CARD_DIR, x))]
    if len(have) != N_PAPERS:
        problems.append(f"卡片 {len(have)}/{N_PAPERS} 不齐，缺 {[x for x in cards if x not in have]}")
    else:
        for x in cards:
            try:
                c = load_json(os.path.join(CARD_DIR, x), {})
                if c.get("flag") or not isinstance(c.get("rows"), list) or len(c["rows"]) < 1:
                    problems.append(f"{x} flag/零行")
            except Exception as e:
                problems.append(f"{x} JSON 错误 {e}")
    src = [x for x in os.listdir(ZHUN) if x.lower().endswith(".pdf")]
    if len(src) != N_PAPERS:
        problems.append(f"准\\ PDF {len(src)} 篇（应为 {N_PAPERS}）")
    elif any(not pdf_ok(os.path.join(ZHUN, x)) for x in src):
        problems.append("准\\ 有坏件")
    exist = [f"{n:02d}.pdf" for n in range(FIRST_NO, LAST_NO + 1)
             if os.path.exists(os.path.join(LIT, f"{n:02d}.pdf"))]
    if exist:
        problems.append(f"文献\\160-177 目标已存在：{exist}")
    fig = load_json(FIGURES_JSON, {}) or {}
    if len(fig.get("items", [])) != 14:
        problems.append("figures.json 缺 14 篇图页候选")
    vision = load_vision()
    if len(vision) < 14:
        problems.append(f"vision_results.json 视觉结果 {len(vision)} 条（应≥14）")
    if problems:
        print("[finalize] 前置门未过：")
        for x in problems:
            print("   -", x)
        raise SystemExit(2)
    print(f"[finalize] 前置门通过：{N_PAPERS} 卡齐 / {N_PAPERS} 源 PDF 全好 / 160-177 目标全空 / "
          f"图页 14 篇 / 视觉结果 {len(vision)} 条 / 0.xlsx 与底本一致")
    steps = ["check", "build", "commit", "move", "verify", "reading"]
    for s in steps:
        print("\n===== finalize ->", s, "=====", flush=True)
        subprocess.run([sys.executable, "-X", "utf8", os.path.abspath(__file__), s], check=True)
    print(f"\n[finalize] 全部 {len(steps)} 步完成")


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "prep"
    cmds = {"prep": cmd_prep, "figs": cmd_figs, "check": cmd_check,
            "build": cmd_build, "commit": cmd_commit, "move": cmd_move,
            "verify": cmd_verify, "reading": cmd_reading_report,
            "finalize": cmd_finalize}
    if mode not in cmds:
        raise SystemExit("未知子命令：" + mode + "（应为 " + "/".join(cmds) + "）")
    cmds[mode]()


if __name__ == "__main__":
    main()
