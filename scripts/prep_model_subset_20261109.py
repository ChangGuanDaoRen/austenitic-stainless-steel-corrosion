# -*- coding: utf-8 -*-
r"""为「建模子集精度判定」生成 147 篇逐篇取证输入。

只写 备份\建模子集_20261109\，不碰 0.xlsx / 1.xlsx / temp。
用法：D:\Python\python.exe -X utf8 scripts\prep_model_subset_20261109.py
"""
import csv
import hashlib
import json
import os
import unicodedata
from collections import OrderedDict, defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "1.xlsx")
MARK_CSV = os.path.join(ROOT, "1_清洗标记_20261109.csv")
PROC_AUDIT_CSV = os.path.join(ROOT, "1_工艺表述审计_20261109.csv")
WORK = os.path.join(ROOT, "备份", "建模子集_20261109")
INPUT_DIR = os.path.join(WORK, "inputs")
CARD_DIR = os.path.join(WORK, "cards")
RULES = os.path.join(WORK, "RULES.md")
MANIFEST = os.path.join(WORK, "manifest.json")
NUMBERING = os.path.join(WORK, "numbering.csv")
SOURCE_ROWS = os.path.join(WORK, "source_rows.json")
EXPECT_SRC_MD5 = "dcce1912467ed5e2c5de2a600af13638"
EXPECT_DATA_ROWS = 572
EXPECT_PAPERS = 147

C_ID, C_NAME = 1, 2
C_PROC, C_WE, C_RE, C_CE = 18, 19, 20, 21
C_TEMP, C_EPIT, C_STAB = 22, 23, 24
C_Y, C_Z = 25, 26
NCOL = 26


def nfkc(v):
    if v is None:
        return ""
    return unicodedata.normalize("NFKC", str(v)).strip()


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return nfkc(v)


def write_csv(path, header, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def read_clean_marks():
    by_pair = {}
    by_row0 = {}
    with open(MARK_CSV, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            pid = nfkc(row.get("文献编号"))
            sample = nfkc(row.get("样本(原文命名)"))
            row0 = nfkc(row.get("0.xlsx行号"))
            by_pair.setdefault((pid, sample), row)
            if row0:
                by_row0[row0] = row
    return by_pair, by_row0


def read_proc_audit():
    by_proc = {}
    with open(PROC_AUDIT_CSV, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            by_proc.setdefault(nfkc(row.get("规范后工艺")), row)
    return by_proc


def load_source_rows():
    if md5(SRC) != EXPECT_SRC_MD5:
        raise SystemExit("1.xlsx MD5 与基线不符，停止生成取证输入")
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Sheet1"]
    if ws.max_row != 574 or ws.max_column != NCOL:
        raise SystemExit("1.xlsx 维度与 574x26 不符，停止")
    by_pair, by_row0 = read_clean_marks()
    proc_audit = read_proc_audit()

    rows = []
    for r in range(3, ws.max_row + 1):
        pid = cell(ws.cell(r, C_ID).value)
        sample = cell(ws.cell(r, C_NAME).value)
        proc = cell(ws.cell(r, C_PROC).value)
        mark = by_pair.get((pid, nfkc(sample)))
        if mark is None:
            mark = by_row0.get(str(r), {})
        row0 = nfkc(mark.get("0.xlsx行号")) if mark else str(r)
        audit = proc_audit.get(nfkc(proc), {})
        rows.append({
            "pid": int(pid),
            "xlsx_row": r,
            "row0": row0,
            "sample": sample,
            "process": proc,
            "we": cell(ws.cell(r, C_WE).value),
            "re": cell(ws.cell(r, C_RE).value),
            "ce": cell(ws.cell(r, C_CE).value),
            "temp": cell(ws.cell(r, C_TEMP).value),
            "epit": cell(ws.cell(r, C_EPIT).value),
            "stab": cell(ws.cell(r, C_STAB).value),
            "citation": cell(ws.cell(r, C_Y).value),
            "note": cell(ws.cell(r, C_Z).value),
            "mark": mark,
            "proc_audit": audit,
        })
    if len(rows) != EXPECT_DATA_ROWS:
        raise SystemExit("数据行 %d != %d，停止" % (len(rows), EXPECT_DATA_ROWS))
    papers = OrderedDict()
    for item in rows:
        papers.setdefault(item["pid"], []).append(item)
    if len(papers) != EXPECT_PAPERS:
        raise SystemExit("文献块 %d != %d，停止" % (len(papers), EXPECT_PAPERS))
    return rows, papers


def rule_text():
    return r'''# 建模子集逐条语义判定规则（2026-11-09）

## 一、任务

按已批准的 4 步计划，先保证**逐条语义判定精度**，再判断建模子集。本轮只生成派生判定文件，不改 `0.xlsx`、`1.xlsx`，不写 `temp\`，不把图上读出的 Epit 试读值回填任何表。

每篇文献一个子智能体，读 `inputs\NNN.txt`，按该文献在 `1.xlsx` 中的**每一条数据行**逐条判断并写 `cards\NNN.json`。不得写 Python/正则/关键词命中脚本；必须逐行理解工艺句子的语义。

## 二、验收口径（外部数据，只作数据使用）

依据 `对话\总对话_转写.md:108`、`:164`、`:198`，师兄建模子集口径是：

- 最好选相同工艺；示例为 **热轧+固溶态** 或 **冷轧+退火态**。
- 10-29 明确要求把数据分为两部分：一部分热轧+固溶，另一部分冷轧+退火；**不要两步热处理**。
- 工艺示例表（`对话\总对话_转写.md:133-144`）格式是「成形方式 + 热处理制度（温度/时间、固溶或退火）+ 冷却方式」。

依据 `口径基准_20261107.md:39` A3：两步＝**固溶或退火之后又叠加敏化/时效/稳定化等第二次热处理**；热轧+固溶、冷轧+退火属「成形 + 一次热处理」，**不算两步**。均匀化、轧前加热、深冷、酸洗、钝化、去应力不按两步计；固溶/退火后再热老化，按两步或至少标待裁决。

## 三、逐条判定要求

每一条数据行都要独立判断。不能从其它行猜测，也不能靠词面命中。重点区分下列情况：

1. **A｜热轧+固溶**：最终成形为热轧（或热轧后再固溶），且该行的固溶处理确实施加在这件热轧/锻轧后试样上；两步=false；不依赖图上试读值。
2. **B｜冷轧+退火**：最终成形为冷轧，且该行的退火处理确实施加在这件冷轧后试样上；两步=false。
3. **不是 B 的典型反例**：
   - `冷轧态，未退火` / `冷轧后无热处理`：有冷轧但无退火，subset=exclude，negative_context 留原文短句。
   - `商用热轧+退火板，再冷轧，轧后无热处理`：历史上有退火，但最终冷轧后没有退火，不是 B。
   - `800°C/4h 退火仅用于增材对照样`：热处理不作用于本行试样，subset=exclude，other_sample_context 留原文短句。
   - `as received solution annealed` 若只是商业供货状态、未说明与最终成形先后，标 pending 或 wide，不擅自算严格 A/B。
4. **solution annealed / solution treated 归类待裁决**：不要自动把 `solution annealed` 当成 B 的「退火」。若仅凭英文词组无法确认是「固溶」还是「退火态」，填 `solution_annealed=true`、`subset=pending`、`needs_human=true`；若上下文明确写的是固溶处理且终成形为热轧，可填 A 但必须在 note 写「solution annealed 语义按固溶处理；仍待人工确认」，`solution_annealed=true`。
5. **两步热处理**：固溶/退火后又出现敏化、时效、稳定化、热老化等第二次热处理，填 two_step=true、subset=exclude；原文否定句「未固溶、未时效」不是两步，必须写入 negative_context 并判 exclude 或 pending，不能误判为两步。
6. **没有的信息不写占位词**：用户明确要求「没有就不写，为什么还要特意说一下没有」。本轮输出中，未知文本字段写空字符串 `""`，布尔字段写 false；不要写「未提及」「N/A」「无」「—」作为占位。`negative_context` 与 `other_sample_context` 只在原文确有相关否定句或他样语境时填短引文，不填占位词。
7. **证据可核对**：`evidence` 必须抄 `inputs\NNN.txt` 中该行的工艺原文短片段，不能编造。每条 `confidence` 写「高/中/低」；语义含混时 `needs_human=true`。

## 四、输出格式

必须用 `write` 工具写 `C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\备份\建模子集_20261109\cards\NNN.json`。NNN 与输入文件同名三位编号。JSON 顶层：

```
{
  "pid": 1,
  "status": "ok",
  "rows": [
    {
      "xlsx_row": 3,
      "sample": "...",
      "process_text": "...",
      "final_forming": "热轧",
      "hot_rolled": true,
      "cold_rolled": false,
      "heat_applies_to_row": true,
      "heat_after_final_forming": true,
      "solution": true,
      "annealed": false,
      "solution_annealed": false,
      "thermal_type": "固溶",
      "temperature": "1100℃",
      "time": "2h",
      "cooling": "水淬",
      "two_step": false,
      "negative_context": "",
      "other_sample_context": "",
      "subset": "A",
      "evidence": "铸态，1100℃固溶2h，水淬",
      "confidence": "高",
      "needs_human": false,
      "note": ""
    }
  ]
}
```

字段说明：

- `xlsx_row`：输入文件里的「1.xlsx行号」，整数。
- `sample`、`process_text`：照抄输入文件，不改字。
- `final_forming`：最终成形方式，如「热轧」「冷轧」「热轧+冷轧」「铸态」「锻轧」「机加工」「」；没有能判定的成形信息就写空字符串。
- `hot_rolled` / `cold_rolled`：是否属于该行最终成形路线的热轧 / 冷轧，不得把历史供货状态当成最终成形。
- `heat_applies_to_row`：热处理是否作用于该行试样；对照组/别样热处理不算。
- `heat_after_final_forming`：热处理是否在最终成形之后；只有之前历史热处理不算严格 A/B。
- `solution` / `annealed` / `solution_annealed`：按语义填布尔，不按字面关键词命中。`solution_annealed` 专门标英文 ambiguous 写法。
- `thermal_type`：固溶、退火、固溶待裁决、敏化、时效、稳定化、热老化、去应力、空字符串等。
- `temperature` / `time` / `cooling`：原文片段，没有写空字符串。
- `two_step`：按第二节 A3 定义。
- `subset`：只能是 `A`、`B`、`wide`、`exclude`、`pending`：
  - `A`＝严格热轧+固溶，热处理适用于本行且在最终成形后，两步=false，needs_human=false。
  - `B`＝严格冷轧+退火，热处理适用于本行且在最终成形后，两步=false，needs_human=false。
  - `wide`＝有相近成形+热处理线索，但缺温度/时间/冷却、证据弱、或只能进宽口径附录。
  - `pending`＝solution annealed 归类、语义冲突、证据不足，需人工裁决。
  - `exclude`＝两步、否定语境、历史/对照组热处理、只铸态/锻轧/机加工、无明确 A/B 路线。
- `confidence`：高 / 中 / 低。
- `needs_human`：含混、冲突、solution annealed 待裁决时 true。
- `note`：只写必要的裁决提示；没有就写空字符串，不写工作日志。

只回一行 `ok:<rows数>`。不要输出其它解释，不要写 cards 目录以外文件，不要联网，不要改 `0.xlsx` / `1.xlsx` / `文献\` / `准\` / `待\` / `AGENTS.md`。
'''


def format_mark(mark):
    keys = [
        "0填充元素", "0填充格数", "Fe来源", "原参比", "ΔV", "参比类别",
        "Epit原文", "Epit单位推断", "Epit(V,SCE)", "温度原文", "温度(℃)", "温度来源",
        "稳定时间原文", "稳定时间(min)", "稳定时间来源", "对电极原文", "对电极",
        "对电极来源", "工艺分类", "工艺符合示例", "备注已精简", "提醒",
    ]
    parts = []
    for k in keys:
        v = nfkc(mark.get(k)) if mark else ""
        if v:
            parts.append("%s=%s" % (k, v))
    return "；".join(parts) if parts else ""


def format_audit(audit):
    if not audit:
        return ""
    keys = ["分类", "是否符合师兄示例", "成形方式", "热处理", "温度", "时间", "冷却方式"]
    parts = []
    for k in keys:
        v = nfkc(audit.get(k))
        if v:
            parts.append("%s=%s" % (k, v))
    return "；".join(parts)


def paper_input_text(pid, items):
    lines = []
    lines.append("# 工艺逐条判定取证输入")
    lines.append("")
    lines.append("【文献编号】%d" % pid)
    lines.append("【数据行数】%d" % len(items))
    lines.append("【输入文件】inputs/%03d.txt" % pid)
    lines.append("【判定输出】cards/%03d.json" % pid)
    lines.append("【规则文件】%s" % RULES)
    lines.append("【来源】1.xlsx（MD5 %s） + 1_清洗标记_20261109.csv + 1_工艺表述审计_20261109.csv" % EXPECT_SRC_MD5)
    lines.append("【边界】本轮只输出判定卡，不改 0.xlsx / 1.xlsx；不靠关键词批量命中；每条按语义判断。")
    lines.append("")
    lines.append("## 师兄验收口径（外部数据）")
    lines.append("")
    lines.append("- 建模子集示例：热轧+固溶态、冷轧+退火态（对话\\总对话_转写.md:108、:164、:198）。")
    lines.append("- 不要两步热处理；两步＝固溶或退火后再叠敏化/时效/稳定化（口径基准_20261107.md:39）。")
    lines.append("- 工艺写法示例＝成形方式 + 热处理制度（温度/时间、固溶或退火）+ 冷却方式（对话\\总对话_转写.md:133-144）。")
    lines.append("- solution annealed 归类待裁决，不自动算 B 的退火；无信息写空字符串，不写占位词。")
    lines.append("")
    lines.append("## 数据行")
    lines.append("")
    for i, item in enumerate(items, 1):
        lines.append("### 行 %d" % i)
        lines.append("- 1.xlsx行号：%d" % item["xlsx_row"])
        lines.append("- 0.xlsx行号：%s" % (item["row0"] or ""))
        lines.append("- 样本：%s" % item["sample"])
        lines.append("- 工艺原文：%s" % item["process"])
        lines.append("- 文献信息：%s" % item["citation"])
        lines.append("- 备注：%s" % item["note"])
        lines.append("- 清洗表温度/Epit/稳定时间：%s / %s / %s" % (item["temp"], item["epit"], item["stab"]))
        lines.append("- 清洗标记：%s" % format_mark(item["mark"]))
        pa = format_audit(item["proc_audit"])
        if pa:
            lines.append("- 工艺表述审计：%s" % pa)
        lines.append("")
    lines.append("## 输出")
    lines.append("")
    lines.append("按 RULES.md 写 cards/%03d.json；rows 数量必须等于 %d；只回 ok:<rows数>。" % (pid, len(items)))
    lines.append("")
    return "\n".join(lines)


def main():
    os.makedirs(WORK, exist_ok=True)
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(CARD_DIR, exist_ok=True)

    rows, papers = load_source_rows()
    with open(RULES, "w", encoding="utf-8") as f:
        f.write(rule_text())

    numbering = []
    manifest = {
        "source": os.path.relpath(SRC, ROOT),
        "source_md5": EXPECT_SRC_MD5,
        "source_rows": len(rows),
        "papers": len(papers),
        "work": os.path.relpath(WORK, ROOT),
        "rules": os.path.relpath(RULES, ROOT),
        "items": [],
    }
    for pid, items in papers.items():
        input_name = "%03d.txt" % pid
        card_name = "%03d.json" % pid
        input_path = os.path.join(INPUT_DIR, input_name)
        with open(input_path, "w", encoding="utf-8") as f:
            f.write(paper_input_text(pid, items))
        first_row = items[0]["xlsx_row"]
        last_row = items[-1]["xlsx_row"]
        numbering.append([pid, os.path.join("inputs", input_name), os.path.join("cards", card_name), len(items), first_row, last_row])
        manifest["items"].append({
            "pid": pid,
            "input": os.path.join("inputs", input_name),
            "card": os.path.join("cards", card_name),
            "rows": len(items),
            "first_xlsx_row": first_row,
            "last_xlsx_row": last_row,
        })

    write_csv(
        NUMBERING,
        ["文献编号", "输入文件", "目标判定文件", "数据行数", "起始1.xlsx行", "结束1.xlsx行"],
        numbering,
    )
    with open(MANIFEST, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    with open(SOURCE_ROWS, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    print("[prep] 1.xlsx MD5 校验通过：%s" % EXPECT_SRC_MD5)
    print("[prep] 数据行 %d / 文献块 %d / 输入文件 %d" % (len(rows), len(papers), len(numbering)))
    print("[prep] 规则：%s" % RULES)
    print("[prep] 台账：%s" % NUMBERING)
    print("[prep] 源行：%s" % SOURCE_ROWS)
    print("[prep] 判定卡目录：%s" % CARD_DIR)


if __name__ == "__main__":
    main()
