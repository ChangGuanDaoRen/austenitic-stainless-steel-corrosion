# -*- coding: utf-8 -*-
r"""第 9 批：把 准\ 83 篇【移动】重编号为 文献\64.pdf..146.pdf，并出入库对照报告。

历史文件名沿用（copy_lit_and_report），但按用户 2026-11-08 12:45 指令
（m01578「已经录入的文献就重新编号移动到文献文件夹」）实际执行【移动】：
  dry  只检查不移动
  run  shutil.move：准\ 原件移走、准\ 腾空；目标已存在则保留源并登记冲突，不覆盖、不删除

对照报告按 RULES 四段：背景 → 方法 → 结论 → 建议。
"""
import csv
import hashlib
import json
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "备份", "entry_20261108")
NUMCSV = os.path.join(WORK, "numbering.csv")
FINAL = os.path.join(WORK, "final_rows.json")
LEDGER = os.path.join(ROOT, "准_入0与重编号对照_20261108.csv")
REPORT = os.path.join(ROOT, "准文献入库对照_20261108.md")
SRC_DIR = os.path.join(ROOT, "准")
DST_DIR = os.path.join(ROOT, "文献")


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    with open(NUMCSV, encoding="utf-8-sig") as f:
        numbering = list(csv.DictReader(f))
    with open(FINAL, encoding="utf-8") as f:
        data = {int(p["no"]): p for p in json.load(f)}
    if len(numbering) != 83 or len(data) != 83:
        raise SystemExit("编号台账或 final_rows 不是 83 篇")

    if any(len(p["rows"]) < 1 for p in data.values()):
        raise SystemExit("存在 0 数据行的卡片，先补证再移动重编号")
    for rec in numbering:
        no = int(rec["新编号"])
        rec["目标文献编号"] = "文献\\" + ("%02d.pdf" % no)
    with open(NUMCSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(numbering[0].keys()))
        w.writeheader()
        w.writerows(numbering)

    before = len([x for x in os.listdir(DST_DIR) if x.lower().endswith(".pdf")])
    rows, problems = [], []
    for rec in numbering:
        no = int(rec["新编号"])
        src_name = rec["原文件名"]
        src = os.path.join(SRC_DIR, src_name)
        dst_name = "%02d.pdf" % no          # 64..99 两位；100..146 三位，沿用 01.pdf..63.pdf 的自然编号习惯
        dst = os.path.join(DST_DIR, dst_name)
        p = data[no]
        item = {
            "新编号": no, "原文件名": src_name, "判级": rec["判级"],
            "题录": p["citation"].replace("\n", " / "), "备注": p["notes"].replace("\n", " "),
            "数据行数": len(p["rows"]), "样本名": "；".join(str(r["sample"]) for r in p["rows"]),
            "目标文件": "文献\\" + dst_name, "源字节": 0, "源MD5": "", "动作": "",
        }
        if not os.path.exists(src):
            problems.append("源缺失：" + src_name)
            item["动作"] = "源缺失"
            rows.append(item)
            continue
        item["源字节"] = os.path.getsize(src)
        item["源MD5"] = md5(src)
        if os.path.exists(dst):
            if md5(dst) == item["源MD5"]:
                item["动作"] = "冲突保留源（目标已存在同MD5）"
                problems.append("目标已存在同 MD5，准\\ 源件未移动（需人工处置）：" + src_name)
            else:
                item["动作"] = "冲突，未覆盖"
                problems.append("目标已存在且内容不同：文献\\" + dst_name)
            rows.append(item)
            continue
        with open(src, "rb") as f:
            if f.read(5) != b"%PDF-":
                problems.append("源不是 PDF：" + src_name)
                item["动作"] = "坏件，未移动"
                rows.append(item)
                continue
        if mode == "run":
            src_md5 = item["源MD5"]
            shutil.move(src, dst)
            if md5(dst) != src_md5:
                item["动作"] = "移动后 MD5 不符"
                problems.append("移动后 MD5 不符：" + dst_name)
            elif os.path.exists(src):
                item["动作"] = "移动后源仍在"
                problems.append("移动后源仍在：" + src_name)
            else:
                item["动作"] = "已移动"
        else:
            item["动作"] = "[dry] 待移动"
        rows.append(item)

    with open(LEDGER, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    after = len([x for x in os.listdir(DST_DIR) if x.lower().endswith(".pdf")])
    remain = len([x for x in os.listdir(SRC_DIR) if x.lower().endswith(".pdf")])
    total_rows = sum(x["数据行数"] for x in rows)
    good = sum(1 for x in rows if x["判级"] == "合规")
    base = sum(1 for x in rows if x["判级"] == "合规-基准组")
    done = sum(1 for x in rows if x["动作"] == "已移动")
    conflict = sum(1 for x in rows if x["动作"].startswith("冲突"))

    lines = [
        "# 准\\ 第 9 批文献入库对照（2026-11-08）",
        "",
        "## 一、背景",
        "",
        "用户 2026-11-08 10:09 指令「先把准里的文献全部入0」，10:32 批准执行并允许多开子智能体；",
        "12:45 追加「已经录入的文献就重新编号移动到文献文件夹」——重编号环节由【复制】改为【移动】，`准\\` 腾空。",
        "`准\\` 内 83 篇为 11-08 逐篇精读后判定「合规 / 合规-基准组」的文献；按定稿口径",
        "`口径基准_20261107.md`（N1–N4）逐篇抽取初始信息，录入 `0.xlsx` 并移动重编号入 `文献\\`。",
        "本文件只记编号与文件的对应关系；逐格数据与校验结论见 `0_第9批入库校验_20261108.md`。",
        "",
        "## 二、方法",
        "",
        "- 编号：以 `文献\\01.pdf`–`63.pdf` 为既有馆藏，新编号连续取 **64–146**，共 83 篇；命名沿用自然编号 `64.pdf`…`146.pdf`。",
        "- 抽取：83 个子智能体并行读 `备份\\entry_20261108\\digest\\NNN.txt`（不全再读 `text\\NNN.txt`），按 `备份\\entry_20261108\\RULES.md` 出 23 字段卡片。",
        "- 入库：`0.xlsx` 以 `备份\\0_基准_第9批入库前_20261108.xlsx` 为底追加，前 285 行零差异，样式/合并区/列宽逐格校验。",
        "- 移动：`shutil.move`，`准\\` 原件移走后腾空；目标已存在则保留源并登记冲突，不覆盖、不删除（用户未授权删除）。",
        "- 台账：`准_入0与重编号对照_20261108.csv`（含源 MD5、字节数、题录、备注、数据行数、样本名、动作）。",
        "",
        "## 三、结论",
        "",
        "| 项目 | 结果 |",
        "| --- | --- |",
        "| 准\\ 待入库 | 83 篇 |",
        "| 判级 | 合规 %d 篇 / 合规-基准组 %d 篇 |" % (good, base),
        "| 抽取数据行 | %d 条 |" % total_rows,
        "| 新增文献编号 | 64–146（连续） |",
        "| 文献\\ 篇数 | %d → %d |" % (before, after),
        "| 本次移动 | %d 篇（冲突保留源 %d 篇） |" % (done, conflict),
        "| 准\\ 剩余 PDF | %d 篇（run 模式应为 0） |" % remain,
        "| 异常 | %d 项 |" % len(problems),
        "",
        "### 3.1 逐篇对照",
        "",
        "| 新编号 | 判级 | 条数 | 原文件名 | 样本名 | 目标文件 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for x in rows:
        lines.append("| %s | %s | %s | %s | %s | `%s` |" % (
            x["新编号"], x["判级"], x["数据行数"],
            x["原文件名"].replace("|", "\\|")[:70],
            str(x["样本名"]).replace("|", "\\|")[:90],
            x["目标文件"]))
    lines += [
        "",
        "### 3.2 异常",
        "",
    ]
    if problems:
        lines += ["- " + p for p in problems]
    else:
        lines.append("- 无。")
    lines += [
        "",
        "## 四、建议",
        "",
        "1. 按 `准_入0与重编号对照_20261108.csv` 抽查 64–146 的题录、样本数与 `0.xlsx` 是否一致。",
        "2. 「合规-基准组」%d 篇只取独立未处理熔铸基准组，备注已写「仅取基准组」，后续建模前需确认是否单列。" % base,
        "3. `准\\` 腾空后，本轮 ima 二次核查判定符合的文献从 `待\\` 放入 `准\\`（只放、不编号、不写 `0.xlsx`）。",
        "4. `0.xlsx` 仍只记初始信息；Epit 统一 SCE、元素补 0、Fe 实算等后续处理留在 `1/2.xlsx`。",
        "5. 暂停中的「精简块 28–63 的 35 条长备注 + 同步两份文档」待用户恢复后再执行。",
        "",
        "指纹 1111155371",
    ]
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print("模式=%s / 文献目录 %d → %d 篇 / 准\\ 剩余 %d 篇 / 数据 %d 条 / 移动 %d / 冲突 %d / 异常 %d"
          % (mode, before, after, remain, total_rows, done, conflict, len(problems)))
    if problems:
        for p in problems[:20]:
            print("  -", p)
    print("台账：" + LEDGER)
    print("报告：" + REPORT)
    if mode == "run" and problems:
        sys.exit(1)


if __name__ == "__main__":
    main()
