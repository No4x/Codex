#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 7 个交互动画 + 9 个教学互动，共 16 个单文件 HTML，
以「嵌入对象」方式挂到 PPT 对应页的页脚中段（导出后处理工序，可重复执行）。

为什么要单独一道工序：
  ppt-master 导出契约禁止 file:// 本地链接，所以链接不能在 SVG 源里写；
  嵌入对象只能在 PPTX 层完成，故本脚本在 svg_to_pptx 导出之后单独跑。

用法：
  python embed_html_objects.py                 # 真跑，产出最终 pptx
  python embed_html_objects.py --src <pptx>    # 指定源 pptx
  python embed_html_objects.py --out <pptx>    # 指定输出 pptx
"""
from __future__ import annotations

import argparse
import pathlib
import shutil
import sys

import pythoncom
import win32com.client as win32

ROOT = pathlib.Path(r"D:/BaiduSyncdisk/工作台_Codex")
PROJ = ROOT / "work/p1/项目一_常用低压电器_ppt169_20260930"
LESSON = ROOT / "knowledge/lesson/电气控制技术实训/2026-2027-1/项目一"
ANIM = LESSON / "交互动画"
INTER = LESSON / "教学互动"

DEFAULT_SRC = PROJ / "exports/项目一_常用低压电器_20260930_142440.pptx"
DEFAULT_OUT = PROJ / "exports/项目一_常用低压电器_20260930_embedded.pptx"

# 幻灯片号（1 基） -> (子目录, HTML 文件名, 标签首行)
PLAN = {
    # —— 交互动画（7）——
    13: (ANIM, "动画1_低压断路器三种脱扣动作.html", "动画1 三种脱扣"),
    26: (ANIM, "动画2_交流接触器工作原理.html", "动画2 接触器动作"),
    36: (ANIM, "动画3_热继电器双金属片动作.html", "动画3 双金属片动作"),
    41: (ANIM, "动画4_通电延时与断电延时对比.html", "动画4 通断电延时"),
    46: (ANIM, "动画5_按钮按下与松开触点动作.html", "动画5 按钮触点动作"),
    51: (ANIM, "动画6_接近开关三线检测.html", "动画6 接近开关检测"),
    54: (ANIM, "动画7_万用表测触点练习.html", "动画7 万用表测触点"),
    # —— 教学互动（9）——
    8:  (INTER, "互动1_看图找元件.html", "互动1 看图找元件"),
    15: (INTER, "互动2_判断断路器好坏.html", "互动2 判断断路器"),
    22: (INTER, "互动3_算一算选熔断器.html", "互动3 算一算选熔断器"),
    30: (INTER, "互动4_哪些触点会变.html", "互动4 哪些触点变"),
    32: (INTER, "互动5_找接线图错误.html", "互动5 找接线错误"),
    38: (INTER, "互动6_谁管过载谁管短路.html", "互动6 过载还是短路"),
    43: (INTER, "互动7_看时序图判类型.html", "互动7 时序图判类型"),
    47: (INTER, "互动8_按钮符号与颜色配对.html", "互动8 符号配色配对"),
    57: (INTER, "互动9_不吸合先查哪里.html", "互动9 排故先查哪"),
}

# 位置：SVG 像素 → 磅（1 px = 0.75 pt，幻灯片 960×540 pt = 1280×720 px）
# 落点在页脚中段空白窗：左避「素材来源」注文（最长止于 x=500px），右避页码（起于 x=1158px）
PX = 0.75
ICON = (700 * PX, 666 * PX, 56 * PX, 44 * PX)   # 图标：525, 499.5, 42, 33
LABEL = (516 * PX, 666 * PX, 176 * PX, 44 * PX)  # 标签：387, 499.5, 132, 33（右对齐）
LABEL_RGB = 100 + 116 * 256 + 139 * 65536        # #64748B
LABEL_TEXT = "%s\n双击图标，浏览器打开"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=str(DEFAULT_SRC))
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    args = ap.parse_args()

    src, out = pathlib.Path(args.src), pathlib.Path(args.out)
    if not src.exists():
        print("源 PPTX 不存在：%s" % src)
        return 1
    missing = [(n, d / f) for n, (d, f, _) in PLAN.items() if not (d / f).exists()]
    if missing:
        for n, p in missing:
            print("缺 HTML：slide %d -> %s" % (n, p))
        return 1

    tmp = out.with_name("_embed_work.pptx")
    shutil.copyfile(src, tmp)
    print("源：%s" % src.name)
    print("工作副本：%s" % tmp)
    print("图标落点（磅）：%s" % (tuple(round(v, 1) for v in ICON),))
    print("标签落点（磅）：%s" % (tuple(round(v, 1) for v in LABEL),))

    pythoncom.CoInitialize()
    app = win32.Dispatch("PowerPoint.Application")
    pres = None
    ok, bad = [], []
    try:
        pres = app.Presentations.Open(str(tmp), WithWindow=False)
        print("页数：%d ｜ 画布：%.0f × %.0f pt"
              % (pres.Slides.Count, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight))
        for idx in sorted(PLAN):
            folder, fname, label = PLAN[idx]
            path = folder / fname
            if idx > pres.Slides.Count:
                bad.append((idx, "超出页数"))
                continue
            sl = pres.Slides(idx)
            try:
                sh = sl.Shapes.AddOLEObject(FileName=str(path), DisplayAsIcon=True)
                sh.Left, sh.Top, sh.Width, sh.Height = ICON
                # PowerPoint 的 OLEFormat 没有 IconLabel，图标只会显示「HTML Document」，
                # 所以在图标左侧补一条中文小标签，讲清「这是哪一个」。
                tb = sl.Shapes.AddTextbox(1, LABEL[0], LABEL[1], LABEL[2], LABEL[3])
                tb.Name = "html-entry-label"
                tb.TextFrame.MarginLeft = tb.TextFrame.MarginRight = 0
                tb.TextFrame.MarginTop = tb.TextFrame.MarginBottom = 0
                tb.TextFrame.WordWrap = True
                tb.TextFrame.VerticalAnchor = 3          # msoAnchorMiddle
                tr = tb.TextFrame.TextRange
                tr.Text = LABEL_TEXT % label
                tr.Font.Size = 9
                tr.Font.Name = "Microsoft YaHei"
                tr.Font.NameFarEast = "Microsoft YaHei"
                tr.Font.Color.RGB = LABEL_RGB
                tr.ParagraphFormat.Alignment = 3         # msoAlignRight
                ok.append((idx, fname, sh.Width, sh.Height))
                print("  ✔ slide %-3d %-8s %s" % (idx, label, fname))
            except Exception as e:
                bad.append((idx, "%s -> %s" % (fname, e)))
                print("  ✘ slide %-3d %s -> %s" % (idx, fname, e))

        if out.exists():
            out.unlink()
        pres.SaveAs(str(out))
        print("已保存：%s（%.2f MB）" % (out.name, out.stat().st_size / 1024 / 1024))
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    print("\n嵌入成功 %d 个，失败 %d 个" % (len(ok), len(bad)))
    for idx, msg in bad:
        print("  - slide %s: %s" % (idx, msg))
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
