#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""对 9 个教学互动 HTML 做真实浏览器交互冒烟测试（Playwright + 本机 Edge）。"""
import pathlib
import sys
from playwright.sync_api import sync_playwright

SRC = pathlib.Path(r"D:/BaiduSyncdisk/工作台_Codex/knowledge/lesson/电气控制技术实训/2026-2027-1/项目一/教学互动")
OUT = pathlib.Path(r"D:/BaiduSyncdisk/工作台_Codex/work/p1/_html_shot/verify")
OUT.mkdir(parents=True, exist_ok=True)

fails = []


def check(cond, msg):
    print(("  OK  " if cond else "  XX  ") + msg)
    if not cond:
        fails.append(msg)


def run(page, idx, path):
    errs = []
    page.on("pageerror", lambda e: errs.append(str(e)))
    page.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
    page.goto(path.as_uri())
    page.wait_for_timeout(300)
    name = path.stem
    print("[%d] %s" % (idx, name))

    # 1) 主按钮：显示全部答案 -> .answer 可见
    primary = page.locator("#btnAll")
    if not primary.count():
        primary = page.locator("button.solid").first
    if primary.count():
        primary.click()
        page.wait_for_timeout(250)
        ans = page.locator(".answer").first
        if ans.count():
            check(ans.is_visible(), "主按钮揭示了 .answer")
        primary.click()  # 收起
        page.wait_for_timeout(200)
        if ans.count():
            check(not ans.is_visible(), "再点一次收起了 .answer")

    # 2) 逐类交互元素点击，确认不报错
    for sel in [".card", ".opt", ".step", ".hot", ".it", ".pick button", ".opts button",
                ".write .line input"]:
        loc = page.locator(sel)
        n = loc.count()
        for i in range(min(n, 8)):
            try:
                loc.nth(i).click(timeout=1500)
            except Exception:
                pass
    page.wait_for_timeout(300)
    check(not errs, "无 JS 报错%s" % ("" if not errs else " -> " + " | ".join(errs[:3])))

    page.screenshot(path=str(OUT / ("v%02d_after.png" % idx)), full_page=True)
    return errs


with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=True)
    page = browser.new_page(viewport={"width": 1340, "height": 1000})
    files = sorted(SRC.glob("*.html"))
    for i, f in enumerate(files, 1):
        run(page, i, f)
    browser.close()

print("\n==== 结果：%d 个断言失败 ====" % len(fails))
for m in fails:
    print(" -", m)
sys.exit(1 if fails else 0)
