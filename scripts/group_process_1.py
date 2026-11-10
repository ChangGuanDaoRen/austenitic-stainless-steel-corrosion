# -*- coding: utf-8 -*-
"""按师兄 2026-10-24 口径筛工艺：成形方式 + 热处理制度（「都选热轧+固溶态，或都选冷轧+退火态」）。

解析逻辑沿用 scripts\split_process.py，并扩充新数据的写法：
  成形：新增「熔炼/浇注/铸造」归为铸造类、「热轧/锻」归为锻轧类
  热处理词表：新增 热老化 / 均匀化 / 正火 / 回火 / 固溶退火

数据源 1.xlsx（150 条）。
用法：python scripts/group_process_1.py [dry|write]
  write 在 1.xlsx 新增/替换「工艺分组」sheet
"""
import collections
import re
import sys

import openpyxl

SRC = "1.xlsx"
COL_ID, COL_NAME, COL_PROC, COL_EPIT = 1, 2, 18, 23
COL_CR, COL_MO, COL_NI = 10, 11, 12

COOL = ["CNT淬火", "水淬", "水冷", "空冷", "炉冷", "油冷"]
STEP_WORDS = ["固溶处理", "固溶退火", "固溶", "退火处理", "退火", "均匀化",
              "敏化处理", "敏化", "稳定化处理", "稳定化", "时效处理", "时效",
              "热老化", "正火", "回火", "保温", "淬火"]
STEP_RE = re.compile(
    r"(\d{3,4}(?:[–\-~]\d{3,4})?)\s*℃\s*(" + "|".join(STEP_WORDS) + r")?"
    r"\s*(\d+(?:\.\d+)?(?:[–\-~]\d+(?:\.\d+)?)?)?\s*(h|min|小时|分钟)?"
)


def parse_form(s):
    s = s or ""
    if s.startswith("铸态"):
        return "铸态（造型未注明）"
    for k in ["砂型铸造", "离心铸造", "熔模铸造", "真空精密铸造"]:
        if s.startswith(k):
            return k
    if "热锻" in s and "冷轧" in s:
        return "铸锭→热锻→冷轧"
    if "热轧" in s and "锻" in s:
        return "锻造→热轧"
    if "热轧" in s:
        return "热轧"
    if "冷轧" in s:
        return "冷轧"
    if "锻" in s:
        return "锻造"
    if "离心" in s:
        return "离心铸造"
    if "砂型" in s:
        return "砂型铸造"
    if "熔模" in s:
        return "熔模铸造"
    if "核级铸造" in s or "铸造" in s or "浇注" in s or "熔炼" in s or "铸锭" in s:
        return "铸造（造型未注明）"
    if s.startswith("锻材"):
        return "商用锻材"
    if "未提及" in s or "未注明" in s:
        return "未注明"
    return "未注明（有热处理记录）"


def parse_heat(s):
    s = s or ""
    steps = [{"温度": m.group(1), "处理": m.group(2) or "",
              "时间": (m.group(3) or "") + (m.group(4) or "")}
             for m in STEP_RE.finditer(s)]
    cool = "水淬（CNT蒸馏水）" if "CNT" in s else next((c for c in COOL if c in s), "")
    surface = ""
    if "未钝化" in s:
        surface = "未钝化"
    else:
        segs = [x.strip() for x in re.split(r"[，,+]", s) if ("钝化" in x or "酸洗" in x)]
        surface = "+".join(segs)
    extra = [w for w in ["敏化", "稳定化", "时效", "热老化"] if w in s]

    main_type, main_t, main_d, main_i = "", "", "", None
    for idx, st in enumerate(steps):
        w = st["处理"]
        if w.startswith(("固溶", "均匀化")):
            main_type, main_t, main_d, main_i = ("固溶" if w.startswith("固溶") else w), st["温度"], st["时间"], idx
            break
        if w.startswith("退火"):
            main_type, main_t, main_d, main_i = "退火", st["温度"], st["时间"], idx
            break
    if not main_type:
        for idx, st in enumerate(steps):
            if st["处理"] in ("保温", "淬火", "热老化", "时效", "正火", "回火", "敏化", "稳定化"):
                main_type, main_t, main_d, main_i = st["处理"], st["温度"], st["时间"], idx
                break
    others = [f'{st["处理"] or "热处理"}{st["温度"]}℃{st["时间"]}'
              for i, st in enumerate(steps) if i != main_i]
    if not main_type:
        if "未热处理" in s or "无热处理" in s:
            main_type = "无热处理"
        elif "未固溶处理" in s:
            main_type = "未固溶（铸态直接用）"
        elif "未热老化" in s:
            main_type = "未热老化"
        elif "未提及" in s:
            main_type = "未提及"
        else:
            main_type = "未提及热处理（原文未写）"
    others_str = "+".join(others)
    attach = "+".join([x for x in [others_str, "+".join(w for w in extra if w not in others_str)] if x])
    return {"类型": main_type, "温度": main_t, "时间": main_d,
            "冷却": cool, "附加": attach, "表面": surface}


def route(form, h):
    parts = [form, h["类型"]]
    for k in ("温度", "时间", "冷却", "附加", "表面"):
        v = h[k]
        parts.append((v + "℃") if k == "温度" and v else v)
    return " | ".join(p for p in parts if p)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    wb = openpyxl.load_workbook(SRC)
    ws = wb["Sheet1"]
    recs = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=COL_NAME).value
        if not name:
            continue
        proc = ws.cell(row=r, column=COL_PROC).value
        f, h = parse_form(str(proc)), parse_heat(str(proc))
        recs.append({"id": ws.cell(row=r, column=COL_ID).value, "name": name,
                     "proc": proc, "form": f, "heat": h, "route": route(f, h),
                     "epit": ws.cell(row=r, column=COL_EPIT).value,
                     "cr": ws.cell(row=r, column=COL_CR).value,
                     "ni": ws.cell(row=r, column=COL_NI).value,
                     "mo": ws.cell(row=r, column=COL_MO).value})

    n = len(recs)
    print(f"总条数 {n}\n")
    print("【1】成形方式分布")
    for k, v in collections.Counter(x["form"] for x in recs).most_common():
        papers = sorted({str(x["id"]) for x in recs if x["form"] == k}, key=int)
        print(f"{v:>4}  {k:<18} 文献{','.join(papers)}")

    print("\n【2】热处理类型分布")
    for k, v in collections.Counter(x["heat"]["类型"] for x in recs).most_common():
        print(f"{v:>4}  {k}")

    print("\n【3】成形 + 热处理类型（师兄口径的粒度）")
    for (f, t), v in collections.Counter((x["form"], x["heat"]["类型"]) for x in recs).most_common():
        items = [x for x in recs if (x["form"], x["heat"]["类型"]) == (f, t)]
        papers = sorted({str(x["id"]) for x in items}, key=int)
        tag = ""
        if "热轧" in f and t == "固溶":
            tag = "  ← 师兄举例A：热轧+固溶"
        elif "冷轧" in f and t == "退火":
            tag = "  ← 师兄举例B：冷轧+退火"
        print(f"{v:>4}  {f} + {t:<12} 文献{','.join(papers)}{tag}")

    print("\n【4】完整工艺路线簇（≥2 条）")
    cl = collections.Counter(x["route"] for x in recs)
    for k, v in cl.most_common():
        if v >= 2:
            items = [x for x in recs if x["route"] == k]
            papers = sorted({str(x["id"]) for x in items}, key=int)
            eps = [x["epit"] for x in items if isinstance(x["epit"], (int, float))]
            print(f"{v:>4} 条 | 文献{','.join(papers)} | Epit {min(eps):.3f}~{max(eps):.3f} | {k}")
    print(f"单条路线 {sum(1 for v in cl.values() if v == 1)} 组；≥2 条 {sum(1 for v in cl.values() if v >= 2)} 组")

    # 师兄举例命中
    hot_sol = [x for x in recs if "热轧" in x["form"] and x["heat"]["类型"] == "固溶"]
    cold_ann = [x for x in recs if "冷轧" in x["form"] and x["heat"]["类型"] == "退火"]
    cold_any = [x for x in recs if "冷轧" in x["form"]]
    hot_any = [x for x in recs if "热轧" in x["form"]]
    cast_sol = [x for x in recs if x["form"].startswith(("铸",)) and x["heat"]["类型"] == "固溶"]
    print(f"\n【5】师兄举例命中")
    print(f"  热轧+固溶：{len(hot_sol)} 条")
    print(f"  冷轧+退火：{len(cold_ann)} 条（冷轧共 {len(cold_any)} 条，其中退火 {len(cold_ann)} 条，"
          f"原文未写退火 {sum(1 for x in cold_any if '未提及' in x['heat']['类型'])} 条）")
    print(f"  热轧（不论后续热处理）：{len(hot_any)} 条 -> {[x['name'] for x in hot_any][:8]}")
    print(f"  铸件+固溶（最接近的替代）：{len(cast_sol)} 条，"
          f"文献{','.join(sorted({str(x['id']) for x in cast_sol}, key=int))}")

    if mode != "write":
        print("\n[dry] 未写盘")
        return

    if "工艺分组" in wb.sheetnames:
        del wb["工艺分组"]
    g = wb.create_sheet("工艺分组")
    g.append(["A. 成形方式"])
    g.append(["成形方式", "条数", "文献编号"])
    for k, v in collections.Counter(x["form"] for x in recs).most_common():
        g.append([k, v, ",".join(sorted({str(x["id"]) for x in recs if x["form"] == k}, key=int))])
    g.append([])
    g.append(["B. 成形 + 热处理类型"])
    g.append(["成形方式", "热处理类型", "条数", "文献编号", "是否师兄举例"])
    for (f, t), v in collections.Counter((x["form"], x["heat"]["类型"]) for x in recs).most_common():
        items = [x for x in recs if (x["form"], x["heat"]["类型"]) == (f, t)]
        tag = "热轧+固溶" if ("热轧" in f and t == "固溶") else ("冷轧+退火" if ("冷轧" in f and t == "退火") else "")
        g.append([f, t, v, ",".join(sorted({str(x["id"]) for x in items}, key=int)), tag])
    g.append([])
    g.append(["C. 完整工艺路线簇"])
    g.append(["工艺路线", "条数", "文献编号", "Epit范围(V)"])
    for k, v in cl.most_common():
        items = [x for x in recs if x["route"] == k]
        eps = [x["epit"] for x in items if isinstance(x["epit"], (int, float))]
        g.append([k, v, ",".join(sorted({str(x["id"]) for x in items}, key=int)),
                  f"{min(eps):.3f}~{max(eps):.3f}" if eps else "-"])
    g.append([])
    g.append(["D. 逐条明细"])
    g.append(["文献编号", "样本", "工艺原文", "成形方式", "热处理类型", "温度(℃)", "时间", "冷却", "Epit(V)"])
    for x in recs:
        h = x["heat"]
        g.append([x["id"], x["name"], x["proc"], x["form"], h["类型"],
                  h["温度"], h["时间"], h["冷却"], x["epit"]])
    wb.save(SRC)
    print(f"\n[write] 已写入 1.xlsx 的「工艺分组」sheet；sheets={wb.sheetnames}")


if __name__ == "__main__":
    main()
