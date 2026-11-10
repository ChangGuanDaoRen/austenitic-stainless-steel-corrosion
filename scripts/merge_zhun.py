# -*- coding: utf-8 -*-
r"""合并三组核定 JSON + 补录 2 篇，输出统一录入数据 temp\lit\entry\final_rows.json。

补录来源（均为回原文/渲染取证，非推测）：
  - 王聪 2024 309S：p3 表1 成分（图片）、p5 表2 击穿电位（图片），4 条
  - 304L 晶粒尺寸篇：原文只给统计分布，无逐试样 Epit，录 1 条占位并标注
修正：
  - 200 系 Table 1 的 Mo 列文字层只有 4 个数值，后 3 组归属不明 → 改「未提及」
  - Epit 缺值统一写「N/A」（与 0.xlsx 现有 2 条一致）
文献编号按 准\_说明.md 清单顺序 28–42 分配。
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

WS = r"C:\Users\WuGuanDaoRen\Desktop\铸造奥氏体不锈钢腐蚀（3.5wt.%NaCl）"
D = os.path.join(WS, "temp", "lit", "entry")

ORDER = [
    # (编号, 原PDF文件名, 组, 篇内索引/匹配用文件名主干)
    (28, "SD_S2405844024019959_Combined_role_of_molybdenum_and_nitrogen_in_Limiting_corrosi.pdf", "g1"),
    (29, "SD_S2238785426027481_304_Nb_microalloying_hot_rolled_solution_treated.pdf", "g1"),
    (30, "WF_王聪_2024_N、Nb合金化对309S奥氏体不锈钢腐蚀性能的影响.pdf", "g1"),
    (31, "SP_s11665_023_07956_9.pdf", "g1"),
    (32, "SD_S0010938X15000992_Effect_of_grain_size_on_pitting_corrosion_of_304L_austenitic.pdf", "g2"),
    (33, "SD_S025405842300737X_Extraordinary_pitting_resistance_obtained_by_forming_a_parti.pdf", "g2"),
    (34, "OAh_Adesina2022_N08029_固溶时长温度_Epit_3.5NaCl.pdf", "g2"),
    (35, "CR_310S_thermomech_anneal_pitting_2023.pdf", "g2"),
    (36, "SP_s12034_025_03477_4.pdf", "g2"),
    (37, "SP_s11665_020_05289_5.pdf", "g3"),
    (38, "WY_srin_201200304.pdf", "g3"),
    (39, "SP_s12666_015_0587_8.pdf", "g1"),
    (40, "SD_S2238785426018016_Effects_of_Mg_and_S_content_on_the_passive_film_characterist.pdf", "g3"),
    (41, "SD_S2352492822005025_Comparison_of_the_microstructure_corrosion_resistance_and_ha.pdf", "g3"),
    (42, "SD_S1044580322006428_Effect_of_grain_size_on_mechanical_property_and_corrosion_be.pdf", "g3"),
]

# 王聪 2024 309S：p3 表1 成分 / p5 表2 Eb（图片放大逐格读出）
WANGCONG_COMP = [
    ("1#", dict(C=0.055, Si=0.32, N=0.068, Nb="未提及", Mn=1.42, Cr=22.27, Ni=12.1), 0.20),
    ("2#", dict(C=0.058, Si=0.54, N=0.171, Nb=0.25, Mn=1.51, Cr=22.54, Ni=12.83), 0.50),
    ("3#", dict(C=0.004, Si=0.58, N=0.005, Nb=0.25, Mn=1.53, Cr=22.42, Ni=12.36), -0.07),
    ("4#", dict(C=0.061, Si=0.52, N=0.170, Nb=0.45, Mn=1.56, Cr=22.35, Ni=12.21), 0.96),
]
WANGCONG_PROC = "真空感应炉冶炼铸锭→1200℃×3h均热→锻成40mm板坯→1200℃×2h→7道次热轧至3mm(开轧1150℃/终轧800℃)→1150℃×1h固溶退火后立即淬火"

# 304L 晶粒尺寸篇：成分 Table 1 (p2)，Epit 仅统计分布
GRAIN_COMP = dict(C=0.03, Si=0.31, N=0.075, P=0.004, S=0.009,
                  Cu=0.05, Mn=1.26, Cr=17.75, Mo=0.06, Ni=8.33,
                  Ti=0.008, Nb=0.002, Co=0.11, Sn="未提及", Fe="余量")


def blank_row(sample, comp, process, WE, RE, CE, temp, pH, Epit, stable, page, quote):
    r = dict(sample=sample, C="未提及", Si="未提及", N="未提及", P="未提及",
             S="未提及", Cu="未提及", Mn="未提及", Cr="未提及", Mo="未提及",
             Ni="未提及", Ti="未提及", Nb="未提及", Co="未提及", Sn="未提及",
             Fe="余量", process=process, WE=WE, RE=RE, CE=CE, temp=temp,
             pH=pH, Epit=Epit, stable=stable, source_page=page, source_quote=quote)
    r.update(comp)
    return r


def main():
    groups = {}
    for g in ("g1", "g2", "g3"):
        groups[g] = json.load(open(os.path.join(D, "data_%s.json" % g), encoding="utf-8"))

    by_file = {}
    for g, d in groups.items():
        for p in d["papers"]:
            key = p["file"]
            by_file[key] = p

    out = []
    for no, fn, _g in ORDER:
        if fn not in by_file:
            raise SystemExit("找不到核定数据: " + fn)
        p = dict(by_file[fn])
        rows = [dict(r) for r in p["rows"]]

        # --- 补录：王聪 309S ---
        if fn.startswith("WF_王聪"):
            rows = []
            for sample, comp, eb in WANGCONG_COMP:
                rows.append(blank_row(
                    sample=sample, comp=comp, process=WANGCONG_PROC,
                    WE="309S被测试样(10mm×10mm×3mm)", RE="SCE(饱和KCl甘汞)",
                    CE="铂电极", temp="未提及", pH="未提及", Epit=eb, stable="60min",
                    page=5, quote="表2 不同N、Nb含量下309S不锈钢的Ecorr、Jcorr和Eb值"))
            p["notes"] = ("成分取 p3 表1、击穿电位 Eb 取 p5 表2（两表均为图片，"
                          "经 200/400 dpi 渲染放大逐格读出，非文字层）。"
                          "介质 3.5%(0.61 mol/L) NaCl，三电极为 SCE / 铂电极 / 试样，"
                          "OCP 稳定 60 min，扫描 0.33 mV/s。击穿电位定义＝腐蚀电流密度"
                          "随阳极极化电位升高而开始快速增大时的电位（阳极区拐点），"
                          "符合师兄 10-05 口径。原文另有 6%FeCl3 浸泡试验，本表只录 3.5% NaCl 极化数据。")

        # --- 补录：304L 晶粒尺寸篇（仅统计分布，Epit 缺） ---
        if fn.startswith("SD_S0010938X15000992"):
            rows = [blank_row(
                sample="304L(晶粒尺寸 5/11/28 μm 统计组)",
                comp=GRAIN_COMP,
                process="1050℃×1h固溶水淬→冷轧70%→1050℃再结晶退火3/15/30min(晶粒5/11/28μm)",
                WE="304L试样(预钝化850mV SCE/30min in 0.1M Na2SO4)",
                RE="SCE", CE="铂片", temp="室温", pH="未提及", Epit="N/A",
                stable="30min", page=3,
                quote="The median of pitting potentials was between 270 mV and 290 mV (SCE).")]
            p["notes"] = ("原文以统计分布研究点蚀电位：每种晶粒尺寸重复 15 次，"
                          "Fig.7 给累积概率分布，正文只给中位数区间 270–290 mV (SCE)，"
                          "并明确「grain refinement has no effect on the pitting potential」。"
                          "无逐试样 Epit 数值 → 本行 Epit 记 N/A，仅录成分与工艺占位；"
                          "若后续按师兄「读阳极区拐点」口径读 Fig.7 三条曲线中位值，可再补 3 条。")

        # --- 修正：200 系 Mo 列归属不明 ---
        if fn == "WY_srin_201200304.pdf":
            for r in rows:
                if r["sample"] in ("204Cu", "204Cu/Nb", "304"):
                    r["Mo"] = "未提及"
            p["notes"] += (" 【主程复核】p3 Table 1 的 Mo 行文字层只有 4 个数值"
                           "(0.09/0.11/0.21/0.03)，后 3 组(204Cu/204Cu-Nb/304)归属不明"
                           "(疑为合并单元格)，本轮改记「未提及」，不猜。")

        # --- 口径冲突标注（写入备注，交师兄裁决） ---
        extra = []
        if fn.startswith("OAh_Adesina"):
            extra.append("【口径存疑·重要】原文 p3 明确所有固溶试样水淬后还经 900℃×0.5h 时效再空冷，"
                         "属「两步热处理」，与师兄 10-29「不要两步热处理的」直接冲突；"
                         "A29 为 as-received 冷轧态未退火。本轮按用户「15 篇全录」录入，待师兄裁决是否保留。")
        if fn == "SP_s12666_015_0587_8.pdf":
            extra.append("【待裁决】原文主体为 600–1000℃ 时效系列，只取 1250℃×3h 固溶态 1 条，"
                         "是否算「单步」待定；且方法节溶液为 14 g NaCl/1000 mL + pH 3.0，"
                         "与摘要/图注的 3.5% NaCl 表述矛盾，已照原文记 pH 3.0。")
        if fn == "WY_srin_201200304.pdf":
            extra.append("【口径存疑】溶液为 3.56 wt% NaCl（ASTM G61 标准液，≈0.609 mol/L）；"
                         "库内文献27(C002 林晓娉1992) 曾因 3.56% NaCl 被整篇删除，"
                         "两处口径不一致，待师兄裁决。Table 1 无 Si、P 两列，按「未提及」填。")
        if fn == "SD_S2405844024019959" or fn.startswith("SD_S2405844024019959"):
            extra.append("【数据待补】3.5% NaCl/50℃ 的动电位极化曲线仅在 p6 Fig.3(a) 图中，"
                         "正文无数值型点蚀电位，3 条 Epit 暂记 N/A；"
                         "按师兄「读阳极区拐点」口径可读图补录，本轮未做读图取点。")
        if fn.startswith("SD_S2352492822005025"):
            extra.append("【口径存疑】As received 两组为热轧原始态，是否归「冷轧+退火」待定；"
                         "参比仅写 Ag/AgCl，未注明 KCl 浓度。")
        if extra:
            p["notes"] = (p.get("notes", "") + "\n" + "\n".join(extra)).strip()

        # Epit 缺值统一为 N/A
        for r in rows:
            if r.get("Epit") in ("未提及", "", None):
                r["Epit"] = "N/A"

        out.append(dict(no=no, file=fn, title=p.get("title", ""),
                        citation=p.get("citation", ""), notes=p.get("notes", ""),
                        evidence=p.get("evidence", []), rows=rows))

    path = os.path.join(D, "final_rows.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    total = 0
    for p in out:
        n = len(p["rows"])
        total += n
        eps = [r["Epit"] for r in p["rows"]]
        print("%2d | %-62s | %2d条 | Epit=%s" % (p["no"], p["file"][:62], n, eps))
    print("合计 %d 篇 / %d 条 -> %s" % (len(out), total, path))


if __name__ == "__main__":
    main()
