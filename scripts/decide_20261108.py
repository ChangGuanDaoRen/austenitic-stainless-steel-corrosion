# -*- coding: utf-8 -*-
"""2026-11-08 定稿判级：在初筛证据基础上套用人工逐篇裁定，输出最终清单。
只读 待\\*.pdf（不移动），只写 CSV / MD。判定：合规 / 存疑-待补证 / 不符 / 重复-已入库 / 重复-库内重复
用法: D:\\Python\\python.exe -X utf8 scripts\\decide_20261108.py
"""
import csv
import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.join(ROOT, "temp", "rescreen_20261108")
SRC = os.path.join(BASE, "text_all.jsonl")
TRI = os.path.join(BASE, "triage_all.csv")
OUT = os.path.join(ROOT, "待_复筛判定_20261108.csv")

# 人工裁定：唯一子串 -> (判定, 依据)
OV = [
    # ---- 合规 ----
    ("s42243_025_01509_1", "合规", "高氮奥氏体不锈钢 V 微合金化；热轧+1050℃×1h固溶单步；3.5 wt% NaCl；SCE；有Epit。6wt%FeCl3为夹杂物试验非本表介质"),
    ("CN8_10.11902_1005.4537.2013.259", "合规", "304奥氏体；3.5%NaCl pH序列；SCE；Eb。pH为条件变量"),
    ("CN9_jcscp_pid20789", "合规", "304；3.5%NaCl；SCE；扫描速率序列。文本层乱码，入库前需人工读图核对"),
    ("maco.202414280", "合规", "304；3.5 wt% NaCl；Ag/AgCl(需−0.031换算)；温度序列"),
    ("mawe.202100178", "合规", "15Cr-9Mn低镍奥氏体；铸锭→热轧→1150℃×2h固溶；3.5 wt% NaCl（另有0.1M H2SO4组不取）；SCE"),
    ("srin.202400144", "合规", "as-cast 304；1100℃×30min固溶单步；3.5 wt% NaCl；SCE；Epit 0.34–0.40 V"),
    ("s10126-005-5145-7", "合规", "UNS S31603；墨西哥湾天然海水（A2保留须标注）；SCE"),
    ("s11837-021-04569-2", "合规", "高氮奥氏体含少量δ相（A5标注）；1100℃固溶单步；3.5 wt% NaCl；SCE；充氢为试验条件"),
    ("s41024-024-00479-3", "合规", "UNS S31603；原位天然海水（A2标注）；SCE"),
    ("s1018-3639_18_30603-2", "合规", "铸块→1050℃锻/热轧；3.5%NaCl；SCE；有Epit"),
    ("a_1020338103964", "合规", "AISI 304超细晶；3.5%NaCl 303K；SCE；Ep-晶粒尺寸"),
    ("JP01_jcorr1974_34_271", "合规", "3.5%NaCl 换热条件；SUS316/317奥氏体（329J1双相组不入库）；SCE"),
    ("JP8_jcorr1974_26_539", "合规", "方法标准文但含SUS304成分表与孔食电位；3.5%NaCl 30℃；V vs SCE"),
    ("JP9_008_isijinternational1989", "合规", "0.6 kmol·m⁻³ NaCl（≡0.6 mol/L等价）；奥氏体；1323K×30min固溶；Ag/AgCl"),
    ("JP9_034_jcorr1991_47_9", "合规", "304不同N；3.5%NaCl；Ag/AgCl(3.3 kmol/m³ KCl)；流速序列"),
    ("JP9_040_jinstmet1952", "合规", "304；3.5 mass% NaCl；Ag/AgCl；紫外光为试验条件"),
    ("JP9_041_matertrans", "合规", "3.5 mass% NaCl；SCE；超声/表面状态条件"),
    ("JP9_061_jcorr_59_12", "合规", "SUS304/430；3.5 mass% NaCl；Ag/AgCl；硝酸钝化前处理（标注）"),
    ("JP9_jcorr1974_26_1", "合规", "304/316；3.5%NaCl 35℃；SCE；孔食电位统计分布"),
    ("JP9_sfj1950_39_12", "合规", "304；3.5 mass%（7.0 mass%组不取）；刻蚀/钝化前处理（标注）"),
    ("a1021625117639", "合规", "高氮低镍奥氏体；1100℃退火；3.5%NaCl 303K；SCE"),
    ("e3sconf_201912104002", "合规", "含氮奥氏体；3.5%NaCl(GOST 9.912)；SSCE需换算；退火温度序列"),
    ("s1516-14392002000100013", "合规", "AISI 304L；3.5%NaCl；SCE；Epit–Ecorr关系"),
    ("1001-1560.2022.0128", "合规", "8367/904L/316L奥氏体；3.5%NaCl；SCE；CPT+Eb"),
    ("2022.04.31", "合规", "304；3.5 wt% NaCl（浓度序列取3.5%组）；SCE"),
    ("iceesd-18.2018.157", "合规", "304/316L海洋大气暴露后3.5%NaCl；SCE"),
    ("intechopen.102381", "合规", "Hercules低镍奥氏体 vs 304；3.5 wt% NaCl；SCE；Tafel仅用于icorr"),
    ("10.5781_jwj", "合规", "超级奥氏体S32654；3.5%NaCl；Ag/AgCl；预应变序列"),
    ("metmateng1402097b", "合规", "316L退火序列单步；3.5%NaCl；SCE；含δ铁素体（标注）"),
    ("s12666_012_0237_3", "合规", "304L；0.5–3.5 wt% NaCl（取3.5%组）；1050℃退火；SCE"),
    ("s42243_024_01365_5", "合规", "316L夹杂物/组织；3.5 wt% NaCl；SCE；含铸态+热轧+IHT单步"),
    ("met10010013", "合规", "Fe-Mn-Al-C奥氏体(SME)；3.5%NaCl；SCE；铸态/热处理/热轧状态序列"),
    ("00035591211265550", "合规", "304/316奥氏体；天然海水+3.5 wt% NaCl；SCE"),
    ("s41598-020-69177-7", "合规", "Fe-Mn-Al-C奥氏体；0.6 M(3.5 wt%) NaCl；SCE；1050℃固溶单步"),
    ("1478422x.2020.1812818", "合规", "316L；3.5%NaCl；1050℃固溶单步；SCC研究但含Epit"),
    ("1743278215y.0000000046", "合规", "SUS304固溶板；3.5%NaCl；SCE；超声为试验条件"),
    ("j.corsci.2016.06.011", "合规", "316L；0.6 M NaCl（另有醋酸钠组不取）；SCE"),
    ("s42243-023-01101-5", "合规", "304含Al2O3夹杂；3.5 wt% NaCl；Epit 0.131/0.304/0.338 V"),
    ("SD8_130504700", "合规", "304形变马氏体（标注）；3.5%NaCl 90℃；Ag/AgCl；击穿电位"),
    ("SD8_22105", "合规", "18Mn-18Cr奥氏体；3.5 wt% NaCl；SCE；温度序列"),
    ("SD8_4236_msa.2017.812065", "合规", "304等；0.6 M NaX（含NaCl组）；SCE"),
    ("SD8_90200778", "合规", "316L；3.5%NaCl；Ag/AgCl；静水压序列"),
    ("SD_S0010938X13001534", "合规", "AISI 316；3.5%NaCl；SCE；扫描速率依赖"),
    ("SD_S0010938X21001554", "合规", "316L固溶态；3.5 wt% NaCl；SCE"),
    ("SD_S0167577X06012717", "合规", "Nb微合金化奥氏体；铸锭→热变形→1050℃退火；3.5%NaCl；SCE"),
    # ---- 不符（人工确认） ----
    ("CH_10_1149_2_0601509jes", "不符", "研究对象为焊接粗晶热影响区(CGHAZ)→N1焊接"),
    ("s11665-012-0467-x", "不符", "介质为16%H2SO4+CuSO4(DL-EPR)，无3.5%NaCl→N3"),
    ("s11837-015-1292-0", "不符", "介质为酸化NaCl(2 mL/L HCl)→N3酸性"),
    ("s12540-020-00789-4", "不符", "316L焊接接头→N1焊接"),
    ("s12666-025-03709-5", "不符", "高氮钢焊接接头(WM/HAZ)→N1焊接"),
    ("maco.202112517", "不符", "固溶1050℃后640℃×100/700h时效→N4两步"),
    ("srin.202300860", "不符", "固溶1050℃后再时效(Effect of Aging)→N4两步"),
    ("JP9_jinstmet1952_57_3", "不符", "18%Cr铸造不锈钢为马氏体+奥氏体+铁素体多相→N2"),
    ("OA9_038_1516-1439", "不符", "12Cr-11Ni-1.7Ti马氏体沉淀硬化钢→N2"),
    ("PW9_iop_10.1149", "不符", "介质为含硼酸模拟浓缩地下水→N3"),
    ("1478422x.2022.2105682", "不符", "搅拌摩擦焊(FSW)接头→N1焊接"),
    ("UP9_018_2446719", "不符", "搅拌摩擦焊(FSW)接头→N1焊接"),
    ("WY_maco_201709753", "不符", "对象为铁素体UNS S43000→N2"),
    # ---- 存疑（人工确认） ----
    ("10.5539_mas.v9n11p119", "存疑-待补证", "3.5与5 wt%并列，试样含900℃时效组，需回原文确认是否两步"),
    ("s12666-016-0851-6", "存疑-待补证", "成形方式与试样状态未在摘要明确，焊接仅见讨论"),
    ("b_jach.0000005579", "存疑-待补证", "人工海水→A2待补证"),
    ("08927019109378161", "存疑-待补证", "天然海水/人工海水/NaCl三介质混合，生物膜条件"),
    ("corrrev.2000.18.6.429", "存疑-待补证", "综述性长文，原始试验数据需核"),
    ("JP8_jcorr1991_42_93", "存疑-待补证", "随机脉冲法观测亚稳孔蚀，Epit数值需核"),
    ("JP9_039_tetsutohagane", "存疑-待补证", "人工海水与天然海水对比，需确认3.5%组"),
    ("31026_j.eng.2018.01.04", "存疑-待补证", "人工海水浓度+溴离子混合卤化物"),
    ("SD8_81212476", "存疑-待补证", "含Alloy 59镍基与1.4462双相，需拆分奥氏体组"),
    ("SD_S0010938X19305487", "存疑-待补证", "介质含硫醇类缓蚀剂→介质被改性"),
    ("SP_bf01161451", "存疑-待补证", "含固溶+时效组与Fe-Mn-Al-C合金，需拆单步组"),
    ("UP9R_052", "存疑-待补证", "人工海水+线材/绳索试样→A2待补证"),
    ("PB8_10_1007_s42243_023_00967_9", "存疑-待补证", "304与超纯铁素体Cr26Mo1并列，含摩擦腐蚀条件"),
]

# 内部重复（保留前者，后者标重复）
DUP_INNER = [
    ("R10_springer_10.1007_s42243-025-01509-1.pdf", "重复-库内重复", "与 CH_10_1007_s42243_025_01509_1.pdf 同一 DOI 10.1007/s42243-025-01509-1"),
    ("R10_hindawi_10.1155_2020_6697227.pdf", "重复-库内重复", "与 H9_009.pdf 同一 DOI 10.1155/2020/6697227"),
    ("SP_s11665_020_05098_w.pdf", "重复-库内重复", "与 R10_springer_10.1007_s11665-020-05098-w.pdf 同一 DOI 10.1007/s11665-020-05098-w"),
    ("SD_S1452398123074965", "重复-库内重复", "与 SD8_130707095 同一 DOI 10.20964/2018.07.74"),
    ("SD_S1452398123077568", "重复-库内重复", "与 SD8_90200778 同一 DOI（静水压 316L）"),
]

ADD = re.compile(r"additive manufactur|selective laser melt|\bSLM\b|\bLPBF\b|laser powder bed|\bDED\b|\bWAAM\b|wire arc|3D print|electron beam melt|增材|激光选区|电弧增材", re.I)
POW = re.compile(r"powder metallurg|\bSPS\b|spark plasma sinter|hot isostatic|粉末冶金", re.I)
COAT = re.compile(r"plasma nitrid|nitriding|carburiz|渗氮|渗碳|镀层|涂层|coating", re.I)
DUP = re.compile(r"duplex stainless|lean duplex|双相不锈钢|duplex\b", re.I)
WELD = re.compile(r"weld\b|welding|weldment|堆焊|焊接", re.I)


def refine(rec, v, t):
    """对自动判「不符」的排除项做 head 命中校验，降低参考文献误报"""
    head = rec["head"][:1500]
    text = rec["text"]
    if t.startswith("N1-增材"):
        if ADD.search(head):
            return "不符", "N1-增材(标题/摘要命中)"
        if len(ADD.findall(text)) >= 3:
            return "存疑-待补证", "N1-增材(正文多处提及，需核是否为研究对象)"
        return "存疑-待补证", "增材仅在参考文献/讨论中提及"
    if t.startswith("N1-表面改性"):
        if COAT.search(head):
            return "不符", "N1-表面改性/涂层(标题/摘要命中)"
        if len(COAT.findall(text)) >= 5:
            return "存疑-待补证", "N1-涂层字样(正文多处，需核)"
        return "存疑-待补证", "涂层/渗氮仅在参考文献中提及"
    if t.startswith("N2-双相"):
        if DUP.search(head):
            return "不符", "N2-双相(标题/摘要命中)"
        if len(DUP.findall(text)) >= 3:
            return "存疑-待补证", "N2-双相字样(正文多处，需核研究对象)"
        return "存疑-待补证", "双相仅在参考文献中提及"
    if t.startswith("N1-粉末"):
        if POW.search(head):
            return "不符", "N1-粉末冶金(标题/摘要命中)"
        return "存疑-待补证", "粉末冶金仅在参考文献中提及"
    return v, t


def main():
    recs = {r["file"]: r for r in (json.loads(l) for l in open(SRC, encoding="utf-8")) if r["tag"] == "DAI"}
    rows = list(csv.DictReader(open(TRI, encoding="utf-8-sig")))
    out = []
    used = set()
    for r in rows:
        f = r["文件"]
        v, t = r["预判定"], r["触发条款"]
        note = ""
        # 1) 显式裁定
        hit = None
        for key, vv, reason in OV:
            if key in f:
                hit = (key, vv, reason)
                break
        if hit:
            used.add(hit[0])
            v, note = hit[1], hit[2]
        elif v == "不符":
            v, t = refine(recs[f], v, t)
            note = t
        # 2) 内部重复
        for key, vv, reason in DUP_INNER:
            if f.startswith(key) or key in f:
                v, note = vv, reason
                used.add(key)
                break
        out.append({"文件名": f, "判定": v, "依据": note or t,
                    "页数": r["页数"], "字数": r["字数"], "DOI": r["DOI"],
                    "介质证据": r["介质证据"][:220], "成形证据": r["成形证据"][:180],
                    "组织证据": r["组织证据"][:150], "热处理证据": r["热处理证据"][:180],
                    "电化学证据": r["电化学证据"][:180]})
    cols = list(out[0].keys())
    with open(OUT, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(out)
    from collections import Counter
    c = Counter(x["判定"] for x in out)
    print("wrote", OUT, len(out))
    print("dist", dict(c))
    miss = [k for k, _, _ in OV if k not in used]
    if miss:
        print("!! 未匹配到文件的裁定键:", miss)


if __name__ == "__main__":
    main()
