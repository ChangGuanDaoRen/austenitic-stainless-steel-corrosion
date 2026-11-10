#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""在用户已登录的调试 Edge 里，用 Kimi 网页版读一张极化曲线图，结果落盘供比对。

只做一件事：上传一张 PNG → 发一段固定提示词 → 等回答结束 → 抓回答文本 + 截图。
"""
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, "备份", "00_录入_20261109", "读图素材", "第1批", "03_fig2_p3.png")
OUTDIR = os.path.join(ROOT, "备份", "00_录入_20261109", "kimi_试点")
os.makedirs(OUTDIR, exist_ok=True)

PROMPT = (
    "这是一篇论文的 Fig.2，动电位阳极极化曲线，介质 3.5% NaCl，参比 SCE。"
    "只做一件事：读出图例中 S0、S1、S5、S12、S13、C0、C9、C11、C12 这九条曲线各自的点蚀电位"
    "（电流密度突然增大、曲线近乎垂直上升处的电位），单位 V vs SCE，保留三位小数。"
    "读不出的写「无法判读」。请用两列表格输出（试样 / 点蚀电位），"
    "最后一行写你采用的判据（例如电流密度 100 μA/cm² 或曲线拐点）。不要给其他分析。"
)


def main():
    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp("http://127.0.0.1:9223")
        ctx = b.contexts[0]
        pages = [x for x in ctx.pages if "kimi.com" in x.url]
        if not pages:
            print("没有 kimi 页面"); sys.exit(1)
        pg = pages[-1]
        pg.bring_to_front()

        # 1 上传
        fi = pg.query_selector("input[type=file]")
        fi.set_input_files(IMG)
        print("已投递文件，等待上传…", flush=True)
        for _ in range(30):
            pg.wait_for_timeout(1000)
            body = pg.inner_text("body")
            if "03_fig2_p3.png" in body or "fig2" in body.lower():
                break
        pg.screenshot(path=os.path.join(OUTDIR, "01_上传后.png"))

        # 2 输入
        ed = pg.query_selector("div[contenteditable='true']")
        ed.click()
        pg.wait_for_timeout(300)
        pg.keyboard.type(PROMPT, delay=8)
        pg.wait_for_timeout(500)
        pg.screenshot(path=os.path.join(OUTDIR, "02_已输入.png"))

        # 3 发送
        pg.keyboard.press("Enter")
        print("已发送，等待回答…", flush=True)

        last_len, stable, text = 0, 0, ""
        for i in range(60):
            pg.wait_for_timeout(3000)
            text = pg.inner_text("body")
            if len(text) == last_len:
                stable += 1
            else:
                stable = 0
                last_len = len(text)
            if stable >= 3 and i > 2:
                break
        pg.wait_for_timeout(1500)
        pg.screenshot(path=os.path.join(OUTDIR, "03_回答.png"), full_page=False)

        # 4 抓最后一条回答
        ans = pg.evaluate(
            """() => {
              const nodes = document.querySelectorAll('.markdown, [class*="markdown"]');
              if (nodes.length) return nodes[nodes.length-1].innerText;
              return document.body.innerText.slice(-3000);
            }"""
        )
        open(os.path.join(OUTDIR, "回答文本.txt"), "w", encoding="utf-8").write(ans)
        print("=== 回答 ===\n" + ans[:2500])


if __name__ == "__main__":
    main()
