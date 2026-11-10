#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把第一批待读图的图件渲染成放大 PNG（v2）。

定位策略：
1. 逐行扫文字层，把 `Fig.3` / `Figure 2 (a)` / `図1` / `图 3`（含全角句点、全角括号）都归一后匹配；
2. 候选行按「上方是否有图像块 → 图像块面积 → 行文本长度」排序，取最像图注的一行；
3. 裁剪图注上方同栏区域；每张图另外再出一张整页 PNG 作为兜底，索引里标明用的是哪种。

输出：备份\00_录入_20261109\读图素材\第1批\
"""
import os
import re

import fitz

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIT = os.path.join(ROOT, "文献")
OUT = os.path.join(ROOT, "备份", "00_录入_20261109", "读图素材", "第1批")

BATCH = {
    1: [("图3a", [("3", "QCNT-treated LC-ASS")])],
    3: [("图2", [("14", "C10")])],
    42: [("图6a", [("204", "CG")])],
    47: [("图12b", [("249", "As-received-950")])],
    49: [("图4a", [("250", "Super 304H SS"), ("251", "Super 304H SS（预浸泡1h）")])],
    80: [("图7b", [("381", "304L"), ("382", "316L"), ("383", "904L")])],
    84: [("图5", [("391", "HNSS（充氢3h）"), ("392", "HNSS（充氢5h）")])],
    85: [("图3", [("394", "UNS S31603（浸泡2天）"), ("396", "UNS S31603（浸泡9天）")])],
    88: [("图1", [("407", "alloy 926")])],
    109: [("图2(b)", [("487", "AISI 316"), ("488", "AISI 316")])],
    137: [("图7", [("597", "304 SS")])],
    144: [("图6", [("600", "Base Material (BM)")])],
    150: [("图1", [("619", "SUS304標準試料（板材）")]), ("图2", [("620", "SUS304線材（多点式試样）")])],
    166: [("图3", [("691", "SUS304")])],
    167: [("图6", [("692", "C"), ("693", "D")])],
    170: [("图8", [("705", "Blank matrix"), ("706", "Matrix containing single MnS"),
                    ("707", "Matrix containing single oxy-sulfide")])],
    173: [("图2a", [("718", "316L 180 MPa"), ("719", "316L 200 MPa"), ("717", "316L unstressed")])],
}


def norm_text(s):
    s = s.replace("．", ".").replace("（", "(").replace("）", ")")
    s = s.replace("\u3000", " ").replace("－", "-")
    return re.sub(r"[ \t]+", " ", s)


def parse_fig(fig):
    m = re.search(r"(\d+)\s*[\(（]?\s*([a-zA-Z])?", fig)
    return int(m.group(1)), (m.group(2) or "").lower()


def line_iter(page):
    out = []
    for b in page.get_text("dict").get("blocks", []):
        for ln in b.get("lines", []):
            txt = "".join(s.get("text", "") for s in ln.get("spans", []))
            if txt.strip():
                out.append((fitz.Rect(ln["bbox"]), norm_text(txt).strip()))
    return out


def candidates(page, num):
    rx = re.compile(rf"(?i)(?:^|[\s\(])(?:fig\.?|figure|图|図)\s*{num}(?![0-9])")
    out = []
    for rect, txt in line_iter(page):
        if rx.search(txt):
            imgs = [i for i in page.get_image_info()
                    if i["bbox"][3] <= rect.y0 + 14 and i["bbox"][1] > 30
                    and not (i["bbox"][2] < rect.x0 - page.rect.x1 * 0.35 or i["bbox"][0] > rect.x1 + page.rect.x1 * 0.35)]
            area = sum((i["bbox"][2] - i["bbox"][0]) * (i["bbox"][3] - i["bbox"][1]) for i in imgs)
            out.append((bool(imgs), area, len(txt), rect, txt))
    out.sort(key=lambda x: (-int(x[0]), -x[1], -x[2]))
    return out


def crop_rect(page, cap):
    W, H = page.rect.x1, page.rect.y1
    allimg = page.get_image_info()
    page_area = W * H
    if allimg and max((i["bbox"][2] - i["bbox"][0]) * (i["bbox"][3] - i["bbox"][1]) for i in allimg) > page_area * 0.6:
        # 扫描页：整页就是一张图，按图注位置取上半页或下半页
        if cap.y0 > H * 0.35:
            return fitz.Rect(0, 0, W, min(H, cap.y1 + 4))
        return fitz.Rect(0, max(0, cap.y0 - 6), W, H)
    imgs = [i for i in page.get_image_info()
            if i["bbox"][3] <= cap.y0 + 14 and i["bbox"][1] > 30
            and not (i["bbox"][2] < cap.x0 - W * 0.35 or i["bbox"][0] > cap.x1 + W * 0.35)]
    if imgs:
        imgs.sort(key=lambda i: i["bbox"][1])
        r = fitz.Rect(imgs[-1]["bbox"])
        for i in reversed(imgs[:-1]):
            b = i["bbox"]
            if r.y0 - b[3] < H * 0.12:
                r = fitz.Rect(min(r.x0, b[0]), min(r.y0, b[1]), max(r.x1, b[2]), max(r.y1, b[3]))
            else:
                break
    else:
        r = fitz.Rect(0, max(0, cap.y0 - H * 0.55), W, cap.y0 - 2)
    r = fitz.Rect(max(0, r.x0 - 10), max(0, r.y0 - 10), min(W, r.x1 + 10), min(H, cap.y1 + 4))
    if r.height < 80:
        r = fitz.Rect(0, max(0, cap.y0 - H * 0.55), W, cap.y1)
    return r


def main():
    os.makedirs(OUT, exist_ok=True)
    idx = ["# 第 1 批读图素材索引（2026-11-09 20:25）\n",
           "每个待填行配两张 PNG：`NN_图X_p页码.png` 是按图注裁出的图区（4 倍放大），`NN_图X_p页码_整页.png` 是整页（3 倍，兜底用）。",
           "读数后请填 `读图待办_20261109.csv` 的第 8–10 列，行号见下表。\n",
           "| 裁剪图 | 整页图 | 文献 | 图号 | 覆盖待填行 | 定位依据 |", "|---|---|---|---|---|---|"]
    ok, fallback, bad = 0, [], []
    for no, figs in sorted(BATCH.items()):
        pdf = os.path.join(LIT, f"{no:02d}.pdf" if no < 100 else f"{no}.pdf")
        if not os.path.exists(pdf):
            bad.append((no, "PDF 缺失"))
            continue
        doc = fitz.open(pdf)
        for fig, items in figs:
            print(f"[{no}] {fig} 开始", flush=True)
            num, sub = parse_fig(fig)
            best = None
            for pno in range(len(doc)):
                cs = candidates(doc[pno], num)
                if not cs:
                    continue
                c = cs[0]
                score = (int(c[0]), c[1], c[2])
                if best is None or score > best[0]:
                    best = (score, pno, c[3], c[4], c[0])
            if best is None:
                # 无图注文字层（扫描页，图注在图里）：输出正文引用该图的页面，整页给出
                rx = re.compile(rf"(?i)(?:fig|图|図)[\.．]?\s*{num}(?![0-9])")
                pages_hit = [p for p in range(len(doc)) if rx.search(doc[p].get_text().replace("．", "."))]
                if not pages_hit:
                    area, pbest = -1, 0
                    for pno in range(len(doc)):
                        a = sum(i["bbox"][2] - i["bbox"][0] for i in doc[pno].get_image_info())
                        if a > area:
                            area, pbest = a, pno
                    pages_hit = [pbest]
                pages_hit = pages_hit[:2]
                names = []
                for pbest in pages_hit:
                    name = f"{no:02d}_{fig.replace('图','fig').replace('(','').replace(')','')}_p{pbest + 1}_整页.png"
                    doc[pbest].get_pixmap(matrix=fitz.Matrix(2.5, 2.5)).save(os.path.join(OUT, name))
                    names.append((pbest + 1, name))
                idx.append("| — | " + "、".join(n for _, n in names) + f" | {no} | {fig} | "
                           + "、".join(f"行{r}（{s}）" for r, s in items)
                           + " | 扫描页无图注文字层，按正文引用页整页给出（p" + "、p".join(str(p) for p, _ in names) + "） |")
                bad.append((no, fig))
                continue
            _, pno, cap, cap_text, has_img = best
            page = doc[pno]
            top_caption = cap.y0 < page.rect.y1 * 0.15
            rect = crop_rect(page, cap)
            stem = f"{no:02d}_{fig.replace('图','fig').replace('(','').replace(')','')}_p{pno + 1}"
            if top_caption or rect.height < page.rect.y1 * 0.18:
                # 图注在页顶或裁出区域过扁：说明图在别处，直接用整页
                page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5)).save(os.path.join(OUT, stem + ".png"))
                how = "图注在页顶，整页兜底"
            else:
                page.get_pixmap(clip=rect, matrix=fitz.Matrix(4, 4)).save(os.path.join(OUT, stem + ".png"))
                how = "图注定位"
            page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5)).save(os.path.join(OUT, stem + "_整页.png"))
            ok += 1
            if not has_img:
                fallback.append((no, fig))
            print(f"[{no}] {fig} → p{pno+1} {how}", flush=True)
            idx.append(f"| {stem}.png | {stem}_整页.png | {no} | {fig} | "
                       + "、".join(f"行{r}（{s}）" for r, s in items) + f" | {how}：{cap_text[:36]} |")
        doc.close()
    idx.append(f"\n裁剪图 {ok} 张，整页图 {ok + len(bad)} 张；未定位到图注 {len(bad)} 处"
               + (("：" + "、".join(f"{n}/{f}" for n, f in bad)) if bad else "。"))
    if fallback:
        idx.append("其中无嵌入图块、靠图注位置裁剪的有 " + str(len(fallback)) + " 处：" + "、".join(f"{n}/{f}" for n, f in fallback) + "。")
    open(os.path.join(OUT, "索引.md"), "w", encoding="utf-8").write("\n".join(idx))
    print(f"裁剪图 {ok} 张；整页兜底 {len(bad)} 处 → {OUT}")
    for n, f in bad:
        print("  兜底", n, f)


if __name__ == "__main__":
    main()
