# -*- coding: utf-8 -*-
"""
把清洗版「加工工艺」列拆成 成形方式 + 热处理结构化字段，并统计同工艺簇。
用法：
  python split_process.py dry      # 只打印解析结果，不写文件
  python split_process.py write    # 生成 _清洗版_v2.xlsx
"""
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版.xlsx"
DST = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）_清洗版_v2.xlsx"

COL_FORM_IN = 17          # 加工工艺（0 基）
COL_PAPER, COL_SAMPLE = 0, 1
COL_EPIT, COL_TEMP = 23, 21
COL_EL = {"C": 2, "Si": 3, "N": 4, "P": 5, "S": 6, "Cu": 7, "Mn": 8,
          "Cr": 9, "Mo": 10, "Ni": 11, "Ti": 12, "Nb": 13, "Co": 14,
          "Sn": 15, "Fe": 16}

COOL = ["CNT淬火", "水淬", "水冷", "空冷", "炉冷", "油冷"]
STEP_WORDS = ["固溶处理", "固溶", "退火处理", "退火", "敏化处理", "敏化",
              "稳定化处理", "稳定化", "时效处理", "时效", "保温", "淬火"]
STEP_RE = re.compile(
    r"(\d{3,4}(?:[–\-~]\d{3,4})?)\s*℃\s*(" + "|".join(STEP_WORDS) + r")?"
    r"\s*(\d+(?:\.\d+)?(?:[–\-~]\d+(?:\.\d+)?)?)?\s*(h|min|小时|分钟)?"
)
NUM_RE = re.compile(r"[–\-~]")


def parse_form(s):
    """成形方式（只看成形，不看热处理）"""
    if s is None:
        return "未注明"
    if s.startswith("铸态"):
        return "铸态（造型未注明）"
    for k in ["砂型铸造", "离心铸造", "熔模铸造", "真空精密铸造"]:
        if s.startswith(k):
            return k
    if "热锻" in s and "冷轧" in s:
        return "铸锭→热锻→冷轧"
    if s.startswith("锻材"):
        return "商用锻材"
    if "未提及" in s or "未注明" in s:
        return "未注明"
    return "未注明（有热处理记录）"


def parse_heat(s):
    """返回 dict：类型/温度/时间/冷却/附加/表面"""
    s = s or ""
    steps = []
    for m in STEP_RE.finditer(s):
        t, w, d, u = m.group(1), m.group(2), m.group(3), m.group(4)
        steps.append({"温度": t, "处理": w or "", "时间": (d or "") + (u or "")})

    # 冷却：CNT 蒸馏水淬火单独标出
    if "CNT" in s:
        cool = "水淬（CNT蒸馏水）"
    else:
        cool = next((c for c in COOL if c in s), "")

    # 表面处理：保留原文细节（钝化液新旧、时长），因为它们是真实工艺差异
    surface = ""
    if "未钝化" in s:
        surface = "未钝化"
    else:
        segs = [seg.strip() for seg in re.split(r"[，,+]", s)
                if ("钝化" in seg or "酸洗" in seg)]
        surface = "+".join(segs)

    extra = [w for w in ["敏化", "稳定化", "时效"] if w in s]

    main_type, main_t, main_d, main_i = "", "", "", None
    for idx, st in enumerate(steps):
        w = st["处理"]
        if w.startswith("固溶"):
            main_type, main_t, main_d, main_i = "固溶", st["温度"], st["时间"], idx
            break
        if w.startswith("退火"):
            main_type, main_t, main_d, main_i = "退火", st["温度"], st["时间"], idx
            break
    if not main_type:
        for idx, st in enumerate(steps):
            if st["处理"] in ("保温", "淬火"):
                main_type, main_t, main_d, main_i = st["处理"], st["温度"], st["时间"], idx
                break
    # 主工序之外的其余工序（第二段保温等）记入附加
    others = [f'{st["处理"] or "热处理"}{st["温度"]}℃{st["时间"]}'
              for i, st in enumerate(steps) if i != main_i]
    if not main_type:
        if "未热处理" in s or "无热处理" in s:
            main_type = "无热处理"
        elif "未固溶处理" in s:
            main_type = "未固溶（铸态直接用）"
        elif "未提及" in s:
            main_type = "未提及"
        else:
            main_type = "未提及热处理（原文未写）"
    others_str = "+".join(others)
    missing = [w for w in extra if w not in others_str]
    attach = "+".join([x for x in [others_str, "+".join(missing)] if x])
    return {"类型": main_type, "温度": main_t, "时间": main_d,
            "冷却": cool, "附加": attach, "表面": surface}


def build_route(form, heat):
    parts = [form, heat["类型"]]
    if heat["温度"]:
        parts.append(heat["温度"] + "℃")
    if heat["时间"]:
        parts.append(heat["时间"])
    if heat["冷却"]:
        parts.append(heat["冷却"])
    if heat["附加"]:
        parts.append(heat["附加"])
    if heat["表面"]:
        parts.append(heat["表面"])
    return " | ".join(parts)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "dry"
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["01_清洗总表"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[0] is not None]

    recs = []
    for r in rows:
        form = parse_form(r[COL_FORM_IN])
        heat = parse_heat(r[COL_FORM_IN])
        recs.append({"row": r, "成形": form, "热": heat,
                     "路线": build_route(form, heat)})

    if mode == "dry":
        print("总条数:", len(recs))
        print("\n--- 成形方式分布 ---")
        for k, v in Counter(x["成形"] for x in recs).most_common():
            print(f"{v:>4}  {k}")
        print("\n--- 热处理类型分布 ---")
        for k, v in Counter(x["热"]["类型"] for x in recs).most_common():
            print(f"{v:>4}  {k}")
        print("\n--- 去重后的解析明细（原工艺串 → 拆解）---")
        seen = {}
        for x in recs:
            seen.setdefault(str(x["row"][COL_FORM_IN]), []).append(x)
        for k, v in sorted(seen.items(), key=lambda kv: -len(kv[1])):
            h = v[0]["热"]
            print(f"[{len(v):>2}条] {k}")
            print(f"        → 成形={v[0]['成形']} 类型={h['类型']} 温度={h['温度']} "
                  f"时间={h['时间']} 冷却={h['冷却']} 附加={h['附加']} 表面={h['表面']}")
            print(f"        → 路线ID={v[0]['路线']}")
        print("\n--- 路线簇（≥2 条）---")
        cl = Counter(x["路线"] for x in recs)
        for k, v in cl.most_common():
            if v >= 2:
                papers = sorted({str(x["row"][COL_PAPER]) for x in recs if x["路线"] == k},
                                key=lambda s: int(s) if s.isdigit() else 99)
                print(f"{v:>4} 条 | 文献{papers} | {k}")
        print("单条路线数:", sum(1 for v in cl.values() if v == 1),
              "；≥2 条路线数:", sum(1 for v in cl.values() if v >= 2))
        return

    # ---------- write ----------
    wbv = openpyxl.load_workbook(SRC)
    head_font = Font(bold=True, size=10)
    fill = PatternFill("solid", fgColor="DDEBF7")

    def put_sheet(title, header, data, widths, freeze="A2"):
        if title in wbv.sheetnames:
            del wbv[title]
        sh = wbv.create_sheet(title)
        sh.append(header)
        for c in sh[1]:
            c.font = head_font
            c.fill = fill
            c.alignment = Alignment(horizontal="center", vertical="center")
        for d in data:
            sh.append(d)
        for i, w in enumerate(widths, 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        sh.freeze_panes = freeze
        sh.auto_filter.ref = f"A1:{get_column_letter(len(header))}{len(data) + 1}"
        return sh

    # 02_工艺拆解：逐条
    h1 = ["文献编号", "样本（原文命名）", "加工工艺（原文）", "成形方式",
          "热处理类型", "热处理温度(℃)", "热处理时间", "冷却方式",
          "附加处理", "表面处理", "工艺路线ID", "Epit(V)", "腐蚀温度"]
    d1 = []
    for x in recs:
        r, h = x["row"], x["热"]
        d1.append([r[COL_PAPER], r[COL_SAMPLE], r[COL_FORM_IN], x["成形"], h["类型"],
                   h["温度"], h["时间"], h["冷却"], h["附加"], h["表面"],
                   x["路线"], r[COL_EPIT], r[COL_TEMP]])
    put_sheet("02_工艺拆解", h1, d1, [9, 26, 42, 22, 18, 13, 11, 10, 12, 14, 46, 9, 10])

    # 03_同工艺簇：按条数排序
    by_route = defaultdict(list)
    for x in recs:
        by_route[x["路线"]].append(x)

    def el_range(items, el):
        vals = [x["row"][COL_EL[el]] for x in items]
        nums = []
        for v in vals:
            try:
                nums.append(float(v))
            except (TypeError, ValueError):
                pass
        if not nums:
            return ""
        return f"{min(nums):g}~{max(nums):g}"

    def epit_cov(items):
        num = na = 0
        for x in items:
            v = x["row"][COL_EPIT]
            try:
                float(v)
                num += 1
            except (TypeError, ValueError):
                na += 1
        return num, na

    h2 = ["工艺路线ID", "成形方式", "热处理类型", "热处理温度(℃)", "热处理时间",
          "冷却方式", "附加处理", "表面处理", "条数", "文献编号", "文献数", "样本数",
          "Epit数值条数", "Epit缺失/NA", "Cr范围(wt.%)", "Ni范围(wt.%)",
          "Mo范围(wt.%)", "是否匹配师兄举例"]
    d2 = []
    for route, items in sorted(by_route.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        x0 = items[0]
        h = x0["热"]
        papers = sorted({str(x["row"][COL_PAPER]) for x in items},
                        key=lambda s: int(s) if s.isdigit() else 99)
        num, na = epit_cov(items)
        is_cast = x0["成形"].startswith("铸") or "铸造" in x0["成形"]
        if "热轧" in x0["成形"] and h["类型"] == "固溶":
            tag = "是（热轧+固溶）"
        elif "冷轧" in x0["成形"] and h["类型"] == "退火":
            tag = "是（冷轧+退火）"
        elif "冷轧" in x0["成形"]:
            tag = "否（冷轧但原文未写退火）"
        elif is_cast and h["类型"] == "固溶":
            tag = "否（铸件+固溶，非师兄举例）"
        elif is_cast:
            tag = "否（铸件，无固溶）"
        else:
            tag = "否（其他）"
        d2.append([route, x0["成形"], h["类型"], h["温度"], h["时间"], h["冷却"],
                   h["附加"], h["表面"], len(items), "、".join(papers), len(papers),
                   len(items), num, na, el_range(items, "Cr"), el_range(items, "Ni"),
                   el_range(items, "Mo"), tag])
    put_sheet("03_同工艺簇", h2, d2,
              [48, 22, 22, 13, 11, 14, 20, 30, 8, 12, 8, 8, 11, 11, 13, 13, 13, 30])

    # 04_工艺汇总：两级聚合，直接回答"哪种数据多"
    h3 = ["聚合层级", "分组键", "条数", "文献编号", "Epit数值条数", "占比"]
    d3 = []
    for k, v in Counter(x["成形"] for x in recs).most_common():
        items = [x for x in recs if x["成形"] == k]
        papers = sorted({str(x["row"][COL_PAPER]) for x in items},
                        key=lambda s: int(s) if s.isdigit() else 99)
        d3.append(["成形方式", k, v, "、".join(papers), epit_cov(items)[0],
                   f"{v / len(recs):.1%}"])
    for k, v in Counter((x["成形"], x["热"]["类型"]) for x in recs).most_common():
        items = [x for x in recs if (x["成形"], x["热"]["类型"]) == k]
        papers = sorted({str(x["row"][COL_PAPER]) for x in items},
                        key=lambda s: int(s) if s.isdigit() else 99)
        d3.append(["成形+热处理类型", f"{k[0]} + {k[1]}", v, "、".join(papers),
                   epit_cov(items)[0], f"{v / len(recs):.1%}"])
    put_sheet("04_工艺汇总", h3, d3, [18, 40, 8, 16, 11, 9])

    wbv.save(DST)
    print("已写出:", DST)


if __name__ == "__main__":
    main()
