# -*- coding: utf-8 -*-
r"""把第 6 轮检索的 8 篇 / 25 条整理成 0.xlsx 追加用的 final_rows3.json。

源：第6轮检索_可录候选_20261105.csv（已逐条取证）
文献编号 56-63，与 文献\56.pdf–63.pdf 一一对应（MD5 已核）。

相对台账做的取证修正（全部来自本轮重新打开 PDF 逐字核对，不是改作者数据）：
  1) WE 列：台账前 5 篇误写「铂」（那是 CE），按原文改回「试样」等；
  2) 文献60 参比：原文 "silver/saturated silver chloride (Ag/AgCl)"，
     台账写「SCE（待换算）」会把未换算数值冒充 SCE 值，改为 Ag/AgCl(sat.KCl)
     （清洗脚本按 -0.042 V 换到 SCE），Epit 保留原值 0.343 V；
  3) 文献60 工艺：原文 "hot rolled above 1200 C and annealed at 1000 to 1100 C"，
     台账「板材，未注明」改为该工艺链；
  4) 新增 X 列「测试前稳定时间」：台账无此列，从各篇实验段取证补齐；
  5) 文献61 温度：298K(25℃) -> 298 K（保留原文开尔文记法，清洗时换算 25℃）；
  6) CE 列按原文细化为 铂片/铂电极/铂丝 等（清洗后都归「铂」，不影响 1.xlsx）。

用法：D:/Python/python.exe scripts/build_entry3_20261105.py
输出：temp\lit\entry3\final_rows3.json
"""
import csv
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

WS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV = os.path.join(WS, "第6轮检索_可录候选_20261105.csv")
OUT = os.path.join(WS, "temp", "lit", "entry3", "final_rows3.json")

# 0.xlsx 的 B..X 共 23 列（26 列版本，pH 列已在 11-05 删酸时移除）
FIELDS = ["sample", "C", "Si", "N", "P", "S", "Cu", "Mn", "Cr", "Mo", "Ni",
          "Ti", "Nb", "Co", "Sn", "Fe", "process", "WE", "RE", "CE",
          "temp", "Epit", "stable"]
CSVKEY = {"sample": "样本", "process": "加工工艺", "WE": "WE", "RE": "RE",
          "CE": "CE", "temp": "温度", "Epit": "Epit(V,vs SCE)"}

PAPERS = [
    dict(no=56, src="Metals2021_11_1124",
         citation="Enhancement of Corrosion Properties of Fe-18Cr-9Mn-5.5Ni-0.3(C + N) Austenitic Stainless Steels by Carbon Alloying\n"
                  "Metals, 2021, 11(7): 1124\nDOI: 10.3390/metals11071124",
         notes="【提取方式】成分取自 p2 Table 1，四组合仅 C/N 与固溶制度不同；Si/P/S/Cu/Mo/Ti/Nb/Co/Sn 未给 → 未提及，Fe 余量。"
               "Epit 取 p8 正文 8–10 次重复均值 0.194/0.353/0.160/0.222 V vs SCE（表内 AISI 304 对比值 0.282 未录）。"
               "工艺：1 kg 真空/压力铸锭→1250℃×2h 均匀化→>1050℃ 热轧 75%→1100/1150/1230℃ 固溶水淬。"
               "稳定：阴极极化 −1 V/SCE 300 s 去空气氧化膜 + 溶液中静置 500 s。介质 3.5 wt.% NaCl、25±1℃、SCE、Pt 对电极。"),
    dict(no=57, src="Coatings2022_12_1769",
         citation="Corrosion Behaviors of Fe-22Cr-16Mn-0.55N High-Nitrogen Austenitic Stainless Steel in 3.5% NaCl Solution\n"
                  "Coatings, 2022, 12(11): 1769\nDOI: 10.3390/coatings12111769",
         notes="【提取方式】成分取自 p2 成分段（C 0.044/Si 0.24/Mn 16.84/P 0.017/S 0.005/Cr 22.40/N 0.55/Ni 1.50/Fe 余），"
               "Cu/Mo/Ti/Nb/Co/Sn 未给 → 未提及。Epit 取 p5 Table 1：锻态 HNS −0.12±0.02、固溶态 ST −0.03±0.01 V vs SCE，只录均值。"
               "溶液 3.5% NaCl 蒸馏水配制并完全除气；SCE 参比、铂片辅助电极。动电位扫描前 OCP 记录 3600 s 至稳态（→ 60 min）。"
               "温度原文未注明，0.xlsx 记「未提及」，1.xlsx 按默认 25℃。"),
    dict(no=58, src="材料研究学报2017_31_9_665",
         citation="三种常用不锈钢的耐局部腐蚀性能\n材料研究学报. 2017, 31(9): 665-671\nDOI: 10.11901/1005.3093.2016.662",
         notes="【提取方式】成分取自 p2 表1；Epit 取 p2 表2 的 Eb（1Cr18Ni9Ti 0.092 / 304 0.413 / 316L 0.268 V）。"
               "工艺：3 mm 冷轧板→1050℃×20 min 水冷固溶。介质 3.5% NaCl、室温、饱和甘汞电极 SCE、铂片辅助电极，"
               "工作电极为 6 mm×6 mm（0.36 cm²）方形面。原文未给 OCP 稳定时间 → 0.xlsx「未提及」、1.xlsx 默认 30 min。"),
    dict(no=59, src="中国腐蚀与防护学报2021_41_5_653",
         citation="两种不锈钢在模拟重水堆一回路溶液和3.5%NaCl溶液中的点蚀行为\n中国腐蚀与防护学报. 2021, 41(5): 653-657",
         notes="【提取方式】成分取自 p2 表1（304L/316L 引漏管取样，成形工艺原文未注明）；Epit 取 p3 表2 在 3.5% NaCl 列的 Eb10 三次重复均值"
               "（304L 30℃ 256 mV / 60℃ 100 mV；316L 30℃ 332 mV / 60℃ 148 mV），换算为 V。Eb100 另 4 条为备用未录；"
               "模拟重水堆一回路溶液（含 Li⁺/Cl⁻）数据不属 3.5% NaCl，未录。SCE 参比、铂对电极、高纯氮除氧，扫描 20 mV/min 自开路电位起。"
               "【去重】腐蚀与防护 2021, 42(9) 同团队 304L 在 30℃ 的 Eb10 同为 256 mV、成分表一致，判为同一批数据分投两刊，未重复录入。"),
    dict(no=60, src="REM2025_78_2_e240044",
         citation="Corrosive behavior of austenitic stainless steel 321 in environments containing ammonium chloride\n"
                  "REM - International Engineering Journal, 2025, 78(2): e240044\nDOI: 10.1590/0370-44672024780044",
         notes="【提取方式】成分取自 p3/p4 Table 1（ASTM A-240 TP321，厂家 Aperam 提供）；Epit 取 p6 Table 3 NaCl 列 321 的 +343±17 mV，"
               "只录均值 0.343 V。表 3 中 304/316L 数值注明引自 Pereira Nicacio et al. 2024，属二手数据，未录；NH4Cl 介质数据不属 NaCl，未录。"
               "【参比已核定】p4 原文 \"The silver/saturated silver chloride (Ag/AgCl) reference electrode was used\"，"
               "为饱和 KCl 型 Ag/AgCl，故 0.xlsx 记 Ag/AgCl(sat.KCl) 并保留原始 0.343 V，1.xlsx 按 −0.042 V 换算为 SCE（0.301 V）；"
               "台账初稿的「SCE（待换算）」已修正。工艺按 p3 原文 \"hot rolled above 1200 °C and annealed at 1000 to 1100 °C\" 补全。"
               "稳定：开路电位稳定 1 h（→ 60 min），扫描 0.16 mV/s，−0.50~+0.70 V，三次重复。介质 3.5% wt./v NaCl，温度原文未注明。"),
    dict(no=61, src="SciRep2018_8_4830",
         citation="Correlation between evolution of inclusions and pitting corrosion in 304 stainless steel with yttrium addition\n"
                  "Scientific Reports, 2018, 8: 4830\nDOI: 10.1038/s41598-018-23273-x",
         notes="【提取方式】成分取自 p2 Table 1；Epit 取 p5 Table 2 五个 Y 含量（0 / 0.007 / 0.013 / 0.019 / 0.049%）的 "
               "−0.19 / <−0.35 / −0.18 / −0.11 / −0.22 V vs SCE。Y 含量不在 0.xlsx 元素列内，写入样本名与证据说明。"
               "工艺：Si–Mo 电阻炉氩气保护熔炼（铸造），炉冷后未经任何热处理、铸态直接使用——按用户 11-05 口径，"
               "铸造成形合法（只排除增材/焊接/粉末冶金）。介质 3.5% NaCl、298 K、SCE 参比、铂片对电极、工作电极暴露 1 cm²。"
               "【上限值】0.007%Y 的 Epit 原文记 <−0.35（低到扫描下限未测出精确值），0.xlsx 按惯例保留字符串；"
               "1.xlsx 按「上限值取该数值」清洗为 −0.35，用于建模时需注意其右删失性质。"),
    dict(no=62, src="中国腐蚀与防护学报2024_44_2_429",
         citation="形变强化和逆相变细晶强化对Fe-18Cr-8Ni钢耐腐蚀性能的影响\n"
                  "中国腐蚀与防护学报. 2024, 44(2): 429-436\nDOI: 10.11902/1005.4537.2023.093",
         notes="【提取方式】成分取自 p2 实验段（3 mm 商用 Fe-18Cr-8Ni 冷轧板）；Epit 取 p5 表1 三个状态的 Epit/V "
               "（Original 0.41 / CR 0.48 / UFG/FG 0.75）。工艺：原样 / 室温 10% 冷轧 / 室温 60% 冷轧→900℃×3 min 退火（逆相变细晶）；"
               "其中 UFG/FG 为「冷轧+退火」单步热处理，正落师兄 10-29 口径的 B 组。介质 3.5% NaCl、25℃、SCE、Pt 辅助电极，"
               "扫描速率 0.5 mV/s。原文未给 OCP 稳定时间 → 1.xlsx 默认 30 min；650℃×1h 敏化是双环再活化试验，与本 Epit 无关。"),
    dict(no=63, src="腐蚀与防护2010_31_10_756",
         citation="节镍型不锈钢的耐腐蚀性能比较\n腐蚀与防护. 2010, 31(10): 756-759",
         notes="【提取方式】成分取自 p2 表1；Epit 取 p2 表2 / 摘要：201 −32 mV、202 −22 mV、304 312 mV（vs SCE），换算为 V。"
               "按 GB/T 17899-1999 动电位法、3.5% NaCl、50 mV/min。工艺：市售不锈钢冷轧薄板，热处理原文未注明。"
               "【排除】439/443 为铁素体不锈钢（165/227 mV），不属奥氏体，未录。原文未给 OCP 稳定时间 → 1.xlsx 默认 30 min。"),
]

# 逐篇取证修正：key = (文献号, 字段名)，value = 新值；每行通用
FIX = {}
for no, we, ce, stab in [
    (56, "试样(10×15×2 mm³)", "铂", "阴极极化−1 V/SCE 300 s，500 s 稳定"),
    (57, "试样", "铂片", "3600 s OCP"),
    (58, "试样(暴露 0.36 cm²)", "铂片", "未提及"),
    (59, "试样", "铂电极", "未提及"),
    (60, "试样(暴露 1 cm²)", "铂丝", "1 h OCP"),
    (61, "试样1cm2", "铂片", "未提及"),
    (62, "试样", "Pt", "未提及"),
    (63, "试样1cm2", "铂", "未提及"),
]:
    FIX[(no, "WE")] = we
    FIX[(no, "CE")] = ce
    FIX[(no, "stable")] = stab

FIX[(60, "RE")] = "Ag/AgCl(sat.KCl)"
FIX[(60, "process")] = "1200℃以上热轧→1000~1100℃退火"
FIX[(61, "temp")] = "298 K"

NUM_RE = re.compile(r"^\s*-?\d+(?:\.\d+)?\s*$")


def to_cell(v):
    """纯数值转 float（与 0.xlsx 既有单元格一致），其余保持原始字符串。"""
    s = str(v).strip()
    if NUM_RE.match(s):
        return float(s)
    return s


def main():
    raw = list(csv.DictReader(open(CSV, encoding="utf-8-sig")))
    raw = [r for r in raw if r.get("来源", "").strip()]
    print("台账 %d 条 / %d 篇" % (len(raw), len({r["来源"] for r in raw})))

    by_src = {}
    for r in raw:
        by_src.setdefault(r["来源"], []).append(r)

    data = []
    for p in PAPERS:
        src_rows = by_src.get(p["src"])
        assert src_rows, "台账缺来源 " + p["src"]
        rows = []
        for src in src_rows:
            rec = {}
            for k in FIELDS:
                key = CSVKEY.get(k, k)
                v = src.get(key, "")
                rec[k] = to_cell(v) if k != "sample" else str(v).strip()
            # 台账无稳定时间列，靠取证补齐
            rec.setdefault("stable", "未提及")
            for (no, field), val in FIX.items():
                if no == p["no"] and field in rec:
                    rec[field] = val
                elif no == p["no"] and field == "stable":
                    rec["stable"] = val
            rows.append(rec)
        data.append(dict(no=p["no"], citation=p["citation"], notes=p["notes"],
                         rows=rows, source=p["src"], doi=src_rows[0]["DOI/出处"],
                         evidence=[r["原始数值与证据"] for r in src_rows]))

    # ---------- 自检 ----------
    assert len(data) == 8
    assert [p["no"] for p in data] == list(range(56, 64))
    total = sum(len(p["rows"]) for p in data)
    assert total == 25, "条数 %d != 25" % total
    for p in data:
        assert p["citation"].strip() and p["notes"].strip()
        for r in p["rows"]:
            assert len(r) == len(FIELDS), (p["no"], len(r))
            for k in FIELDS:
                assert r[k] is not None and r[k] != "", (p["no"], r["sample"], k)
            assert r["WE"] != "铂", "WE 未修正：文献 %d" % p["no"]
            assert r["RE"] in ("SCE", "Ag/AgCl(sat.KCl)"), (p["no"], r["RE"])
            assert isinstance(r["Epit"], (int, float)) or r["Epit"] == "<-0.35"
    # 文献编号与文件一一对应
    for p in data:
        f = os.path.join(WS, "文献", "%02d.pdf" % p["no"])
        assert os.path.isfile(f), f

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(data, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("输出 %s" % OUT)
    print("\n逐篇明细：")
    for p in data:
        ep = [r["Epit"] for r in p["rows"]]
        print("  %d | %-30s %d 条 | WE=%-18s RE=%-18s 稳定=%s\n      Epit=%s" %
              (p["no"], p["source"][:30], len(p["rows"]), p["rows"][0]["WE"][:18],
               p["rows"][0]["RE"], p["rows"][0]["stable"], ep))
    print("\n取证修正共 %d 处（WE/CE/稳定时间/参比/工艺/温度）" % len(FIX))


if __name__ == "__main__":
    main()
