#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把一批单文件 HTML（交互动画 / 教学互动）以「嵌入对象」方式挂到 PPTX 对应页的页脚中段。

为什么必须是「导出后工序」：
  ppt-master 的 SVG→PPTX 导出契约禁止 file:// 本地链接（只允许 https:// 与 #slide-N），
  所以「让 PPT 打开本地 HTML」写不进 SVG 源，只能在 PPTX 层完成。
  ⇒ 每次重新导出 PPT 之后，都必须重跑本脚本，否则嵌入对象全部丢失。

用法
----
  1) 先改下面 CONFIG 段（或全部用命令行参数覆盖）
  2) python embed_html_objects.py --dry-run          # 只检查 HTML 是否齐、算落点
  3) python embed_html_objects.py                    # 真跑

  python embed_html_objects.py \
      --src exports/xxx_151612.pptx \
      --out exports/xxx_embedded.pptx \
      --map map.json \
      --icon 700,666,56,44 --label 516,666,176,44

map.json 格式（键 = 幻灯片 1 基页号）：
  { "13": ["动画/动画1_xxx.html", "动画1 三种脱扣"], "8": ["互动/互动1_xxx.html", "互动1 看图找元件"] }
  路径相对于 --html-root（默认 = 当前工作目录）。

依赖：pywin32（pythoncom / win32com.client）+ 本机安装的 PowerPoint。
"""
from __future__ import annotations

import argparse
import json
import pathlib
import shutil
import sys

# ══════════════════════════════════════════════════════════════════════
# CONFIG —— 改这三处即可复用
# ══════════════════════════════════════════════════════════════════════
DEFAULT_SRC = ""          # 导出中间件 PPTX（每次导出后要换成最新时间戳那份！）
DEFAULT_OUT = ""          # 输出 PPTX
DEFAULT_HTML_ROOT = "."   # HTML 根目录（map 里的相对路径基于它解析）
DEFAULT_MAP = ""          # PLAN 的 JSON 文件路径；留空则用下面的内联 PLAN

# 内联挂接表：幻灯片 1 基页号 -> (HTML 相对路径, 图标左侧的中文短标题)
PLAN: dict[int, tuple[str, str]] = {
    # 13: ("交互动画/动画1_低压断路器三种脱扣动作.html", "动画1 三种脱扣"),
}

# 落点（SVG px，1 px = 0.75 pt；画布 1280×720 px = 960×540 pt）
#   默认值 = 页脚中段空白窗：左避「素材来源」注文（最长止于 x≈500），右避页码（起于 x≈1158）
PX = 0.75
ICON_PX = (700, 666, 56, 44)      # 图标：x y w h
LABEL_PX = (516, 666, 176, 44)    # 中文标签：x y w h（右对齐）

LABEL_RGB = 100 + 116 * 256 + 139 * 65536     # #64748B
LABEL_TEXT = "%s\n双击图标，浏览器打开"
FONT_NAME = "Microsoft YaHei"
# ══════════════════════════════════════════════════════════════════════


def load_plan(map_path: str, html_root: pathlib.Path) -> dict[int, tuple[pathlib.Path, str]]:
    if map_path:
        raw = json.loads(pathlib.Path(map_path).read_text(encoding="utf-8"))
    else:
        raw = {str(k): [v[0], v[1]] for k, v in PLAN.items()}
    if not raw:
        raise SystemExit("挂接表为空：请填写 CONFIG.PLAN，或用 --map 指定 JSON。")
    return {int(k): (html_root / v[0], v[1]) for k, v in raw.items()}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--map", default=DEFAULT_MAP)
    ap.add_argument("--html-root", default=DEFAULT_HTML_ROOT)
    ap.add_argument("--icon", default="", help="SVG px，格式 x,y,w,h")
    ap.add_argument("--label", default="", help="SVG px，格式 x,y,w,h")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not args.src or not args.out:
        raise SystemExit("必须指定 --src 与 --out（或在 CONFIG 里填 DEFAULT_SRC / DEFAULT_OUT）。")

    html_root = pathlib.Path(args.html_root)
    plan = load_plan(args.map, html_root)

    icon_px = tuple(float(v) for v in args.icon.split(",")) if args.icon else ICON_PX
    label_px = tuple(float(v) for v in args.label.split(",")) if args.label else LABEL_PX
    icon = tuple(v * PX for v in icon_px)
    label = tuple(v * PX for v in label_px)

    src, out = pathlib.Path(args.src), pathlib.Path(args.out)
    if not src.exists():
        print("源 PPTX 不存在：%s" % src)
        return 1
    missing = [(n, p) for n, (p, _) in plan.items() if not p.exists()]
    if missing:
        for n, p in missing:
            print("缺 HTML：slide %d -> %s" % (n, p))
        return 1

    print("挂接 %d 个 HTML" % len(plan))
    print("图标落点（磅）：%s" % (tuple(round(v, 1) for v in icon),))
    print("标签落点（磅）：%s" % (tuple(round(v, 1) for v in label),))
    for n in sorted(plan):
        print("  slide %-3d %-22s %s" % (n, plan[n][1], plan[n][0].name))
    if args.dry_run:
        print("\n[dry-run] 检查通过，未写入。")
        return 0

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        print("缺 pywin32：请用 ppt-master 的 venv 运行，或 pip install pywin32")
        return 1

    tmp = out.with_name("_embed_work.pptx")
    shutil.copyfile(src, tmp)
    print("\n源：%s\n工作副本：%s" % (src.name, tmp))

    pythoncom.CoInitialize()
    app = win32.Dispatch("PowerPoint.Application")
    pres = None
    ok, bad = [], []
    try:
        pres = app.Presentations.Open(str(tmp), WithWindow=False)
        print("页数：%d ｜ 画布：%.0f × %.0f pt"
              % (pres.Slides.Count, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight))
        for idx in sorted(plan):
            path, label_txt = plan[idx]
            if idx > pres.Slides.Count:
                bad.append((idx, "超出页数"))
                continue
            sl = pres.Slides(idx)
            try:
                # ⚠️ AddOLEObject 只吃关键字参数；位置参数报 could not convert string to float
                sh = sl.Shapes.AddOLEObject(FileName=str(path), DisplayAsIcon=True)
                # ⚠️ 命名参数里的 Left/Top 会被忽略 → 必须创建后再赋值
                sh.Left, sh.Top, sh.Width, sh.Height = icon
                # ⚠️ PowerPoint 的 OLEFormat 没有 IconLabel → 图标只显示「HTML Document」
                #    所以在图标左侧补一条中文小标签，讲清「这是哪一个」
                tb = sl.Shapes.AddTextbox(1, label[0], label[1], label[2], label[3])
                tb.Name = "html-entry-label"
                tf = tb.TextFrame
                tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
                tf.WordWrap = True
                tf.VerticalAnchor = 3            # msoAnchorMiddle
                tr = tf.TextRange
                tr.Text = LABEL_TEXT % label_txt
                tr.Font.Size = 9
                tr.Font.Name = FONT_NAME
                tr.Font.NameFarEast = FONT_NAME
                tr.Font.Color.RGB = LABEL_RGB
                tr.ParagraphFormat.Alignment = 3  # msoAlignRight
                ok.append(idx)
                print("  OK  slide %-3d %s" % (idx, label_txt))
            except Exception as e:                # noqa: BLE001
                bad.append((idx, "%s -> %s" % (path.name, e)))
                print("  XX  slide %-3d %s -> %s" % (idx, path.name, e))

        if out.exists():
            out.unlink()
        pres.SaveAs(str(out))
        print("已保存：%s（%.2f MB）" % (out.name, out.stat().st_size / 1024 / 1024))
    finally:
        for fn in (lambda: pres and pres.Close(), lambda: app.Quit()):
            try:
                fn()
            except Exception:                     # noqa: BLE001
                pass
        pythoncom.CoUninitialize()

    print("\n嵌入成功 %d 个，失败 %d 个" % (len(ok), len(bad)))
    for idx, msg in bad:
        print("  - slide %s: %s" % (idx, msg))
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
