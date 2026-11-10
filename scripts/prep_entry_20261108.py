# -*- coding: utf-8 -*-
"""第9批（准\\ 83 篇）入库前准备：
1) 备份 0.xlsx -> 备份\\0_基准_第9批入库前_20261108.xlsx
2) 编号 64-146，写编号台账
3) 从 备份\\text_all.jsonl 取全文，合并逐篇判级依据，生成每篇抽取输入
4) 生成抽取规则 RULES.md（含前段人工录入样例）
只读 0.xlsx / text_all.jsonl / 判级 JSON / 准\\；新增文件只写 备份\\entry_20261108\\，不占 temp 配额。
"""
import csv
import json
import os
import shutil
from collections import Counter

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
TEXT_DIR = os.path.join(WORK, "text")
CARD_DIR = os.path.join(WORK, "cards")
for d in (WORK, TEXT_DIR, CARD_DIR):
    os.makedirs(d, exist_ok=True)

# ---------- 1) 备份 ----------
src0 = os.path.join(ROOT, "0.xlsx")
bak0 = os.path.join(ROOT, "备份", "0_基准_第9批入库前_20261108.xlsx")
if not os.path.exists(bak0):
    shutil.copy2(src0, bak0)
print("backup0 bytes=", os.path.getsize(bak0), flush=True)

# ---------- 2) 全文 ----------
by_file = {}
with open(os.path.join(ROOT, "备份", "text_all.jsonl"), encoding="utf-8") as f:
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

# ---------- 3) 判级依据 ----------
idx2file = {}
mp = os.path.join(ROOT, "备份", "judge_20261108_manifest.json")
if os.path.exists(mp):
    with open(mp, encoding="utf-8") as f:
        for m in json.load(f):
            idx2file[m["idx"]] = m["file"]
file2v = {}
vd = os.path.join(ROOT, "备份", "judge_verdicts_20261108")
if os.path.isdir(vd):
    for fn in os.listdir(vd):
        if not fn.endswith(".json"):
            continue
        with open(os.path.join(vd, fn), encoding="utf-8") as f:
            v = json.load(f)
        fn2 = idx2file.get(v.get("idx"))
        if fn2:
            file2v[fn2] = v

files = sorted([f for f in os.listdir(os.path.join(ROOT, "准")) if f.lower().endswith(".pdf")])
print("zhun pdf=", len(files), flush=True)

# ---------- 4) 前段人工录入样例 ----------
wb = openpyxl.load_workbook(src0, data_only=True)
ws = wb["Sheet1"]
FIELDS = ["文献编号", "样本", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "加工工艺", "工作电极", "参比电极", "对电极",
          "腐蚀温度", "Epit(V)", "测试前稳定时间", "文献信息", "备注"]
examples = []
for r in range(3, ws.max_row + 1):
    if ws.cell(r, 1).value not in (None, ""):
        examples.append({FIELDS[c - 1]: ws.cell(r, c).value for c in range(1, 27)})
    if len(examples) >= 3:
        break
with open(os.path.join(WORK, "examples.json"), "w", encoding="utf-8") as f:
    json.dump(examples, f, ensure_ascii=False, indent=2)
print("examples:", flush=True)
for e in examples:
    print(json.dumps(e, ensure_ascii=False)[:700], flush=True)

RULES = """# 第9批入库：逐篇抽取规则（本轮唯一生效）

你是「数据抽取员」。任务：把指定**一篇**论文里符合定稿收录口径的实验数据行，抽成结构化 JSON 卡片，供写入 `0.xlsx`。`0.xlsx` 的定位是**只记从文献提取的初始信息**，后续归一/换算全部交给 `1.xlsx`、`2.xlsx`，本卡片一律不做换算。

## 输入与输出
- 规则：本文件 `备份\\entry_20261108\\RULES.md`
- 论文全文：`备份\\entry_20261108\\text\\NNN.txt`（NNN 由任务给出，三位数字，等于「新编号」）
- 输出：`备份\\entry_20261108\\cards\\NNN.json`（UTF-8，`ensure_ascii=False`）
- 完成后**只回一行** `ok:<rows数>` 或 `flag:<原因>`；不要把全文或长表格贴回回复里。

## 定稿口径（依据 `口径基准_20261107.md`）
- **N1 成形**：熔炼—凝固路线（铸造；铸锭后再锻、热轧、冷轧均可）；排除增材制造(SLM/LPBF/DED/WAAM)、粉末冶金、焊接/堆焊、涂层/镀层/渗氮渗碳等表面成分改性。
- **N2 组织**：奥氏体不锈钢为主；奥氏体+少量 δ 铁素体视为奥氏体（在 notes 标注）；真双相(2205 类)、马氏体/铁素体为主 → 不录。
- **N3 介质**：3.5 wt.% NaCl 及其等价写法（0.6 mol/L、0.61 mol/L、0.6 M、0.6 kmol/m³、35 g/L、600 mM、3.5 mass% 等）；天然海水可收（notes 标注「天然海水」）；人工海水/模拟海水(ASTM D1141、SME)/3.56%/3.5% wt./v/浓度序列中的非 3.5 组/酸碱调 pH/薄液膜/模拟一回路液 → 不录。
- **N4 热处理**：不要两步热处理（固溶或退火之后再叠敏化/时效/稳定化）；热轧+固溶、冷轧+退火不算两步；均匀化、轧前加热、熔炼、深冷、酸洗、钝化不计步。
- **Epit**：取阳极区拐点，不做 Tafel 外推；须有可识别参比电极。

## 抽取范围
- 只抽**符合上述口径**的数据行。同一论文里其它介质、其它工艺、其它牌号的行一律不抽。
- 浓度序列只取 3.5% 组；判级为「合规-基准组」的论文**只取独立未处理熔铸基准组**的数据行，并在 `notes` 写「仅取基准组」。
- **至少输出 1 行**。若确实找不到可录数值，仍输出 1 行占位（各字段写「未提及」/「N/A」），顶层加 `"flag": "原因"`。

## 字段（键名严格照此，值全部为字符串；顶层 newno 为整数）
每行 `rows[]` 含：
`sample, C, Si, N, P, S, Cu, Mn, Cr, Mo, Ni, Ti, Nb, Co, Sn, Fe, process, WE, RE, CE, temp, Epit, stable`
顶层：`newno`(int), `file`(str), `citation`(str), `notes`(str), `rows`(array)，可选 `flag`(str)。

### 写法（照前段人工录入习惯，见文末样例）
- **15 个元素**：按原文**逐元素照录字面值**。原文未给该元素写「未提及」；Fe 按原文（多为「余量」/Bal./balance）；比较符（`<0.03`、`≤0.02`）与区间（`18–20`）照录；`11 ppm` 等带单位照录。**不要**换算、补 0、反算 Fe、统一小数位。
- `sample`：原文样本名（含牌号/编号/处理标识）。
- `process`：原文成形+热处理路线的中文简述，如「真空感应熔炼铸锭→热锻→冷轧37.5%」「铸态，1120℃固溶5h，水淬」。
- `WE` 工作电极 / `RE` 参比电极 / `CE` 对电极：记原文写法（`SCE`、`Ag/AgCl(饱和KCl)`、`Pt`、`石墨`等）。**不要**把 Epit 换算到 SCE；RE 里不要写「→SCE」「0.1989 V vs SHE」这类换算说明。
- `temp`：原文写法（`25℃`、`298 K`、`25±1℃`、`室温`）。
- `Epit`：原文**数值字符串**（如 `0.412`、`-0.116`），不做 SCE 换算；只有曲线时读阳极拐点并在 notes 写「Epit 由图读取」；原文无数值写 `N/A` 并在 notes 说明。
- `stable`：原文写法（`30 min`、`1800 s`、`1 h OCP`、`阴极极化−1 V/SCE 10 min`）；原文未给写「未提及」。
- `citation`（Y 列「文献信息」）：作者. 题名[J]. 期刊, 年, 卷(期): 页. DOI/来源。尽量给出题名+期刊+年份；可从全文首页/DOI/文件名推断，缺项只写能确定的部分，**不要编造**。
- `notes`（Z 列「备注」）：**短**，只写 ①数据来源（如「成分取自 Table 1；Epit 取自 Fig.3」）②特殊情况（「天然海水」「仅取基准组」「Epit 由图读取」「原文未给固溶温度」「无稳定钝化区，Epit 记 N/A」）。**不写**：普通介质复述（能入库即默认介质合规）、口径推演、工作日志抬头、对 1/2.xlsx 的指示、实验条件复述（这些已在各列）。尽量 <80 字。
- `AA`（口径标注）：**不要**出现在卡片里，脚本统一留空。

## 前段人工录入样例（照此风格，注意 citation 与 notes 的写法）
{{EXAMPLES}}
"""
RULES = RULES.replace("{{EXAMPLES}}", json.dumps(examples, ensure_ascii=False, indent=2))
with open(os.path.join(WORK, "RULES.md"), "w", encoding="utf-8") as f:
    f.write(RULES)

# ---------- 5) 逐篇输入文件 ----------
MARK = "================ 论文全文 ================\n"
manifest = []
ocr_needed = []
missing = []
for i, fn in enumerate(files):
    newno = 64 + i
    idx = "%03d" % newno
    rec = by_file.get(fn)
    v = file2v.get(fn, {})
    text = (rec or {}).get("text") or ""
    if rec is None:
        missing.append(fn)
    if len(text) < 2000:
        ocr_needed.append({"idx": idx, "newno": newno, "file": fn})
    ev = "\n".join("【%s】%s" % (k, v.get(k) or "") for k in
                   ["判定", "触发条款", "介质原文", "成形与组织原文", "热处理原文", "Epit与参比", "依据"])
    hdr = ("【新编号】%d\n【原文件】%s\n%s\n【DOI】%s\n【全文字数】%d\n"
           % (newno, fn, ev, (rec or {}).get("doi") or "", len(text)))
    if len(text) > 90000:
        body = text[:60000] + "\n\n……［中段略］……\n\n" + text[-30000:]
    else:
        body = text
    with open(os.path.join(TEXT_DIR, idx + ".txt"), "w", encoding="utf-8") as f:
        f.write(hdr + MARK + body)
    manifest.append({
        "idx": idx, "newno": newno, "file": fn,
        "verdict": v.get("判定") or "", "chars": len(text),
        "text": os.path.join("备份", "entry_20261108", "text", idx + ".txt"),
    })

with open(os.path.join(WORK, "manifest.json"), "w", encoding="utf-8") as f:
    json.dump(manifest, f, ensure_ascii=False, indent=2)
with open(os.path.join(WORK, "ocr_needed.json"), "w", encoding="utf-8") as f:
    json.dump(ocr_needed, f, ensure_ascii=False, indent=2)
with open(os.path.join(WORK, "numbering.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["新编号", "原文件名", "判级", "全文字数", "目标文献编号"])
    for m in manifest:
        w.writerow([m["newno"], m["file"], m["verdict"], m["chars"],
                    "文献\\%03d.pdf" % m["newno"]])

print("missing_text=", len(missing), missing[:5], flush=True)
print("ocr_needed=", len(ocr_needed), [o["file"] for o in ocr_needed], flush=True)
print("verdicts=", dict(Counter(m["verdict"] for m in manifest)), flush=True)
print("newno range=64..%d" % (64 + len(files) - 1), flush=True)
