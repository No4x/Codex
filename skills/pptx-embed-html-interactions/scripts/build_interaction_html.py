#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 work/p1/html_src/*.tpl.html 里的 %%IMG:文件名:宽度%% 占位符
替换成内嵌 base64 data URI（等比缩放到指定宽度、白底、JPEG q82），
输出单文件零依赖 HTML 到 knowledge/.../项目一/教学互动/。

用法：
  python build_interaction_html.py            # 构建
  python build_interaction_html.py --check    # 只校验有没有残留占位符
"""
from __future__ import annotations
import base64
import io
import re
import sys
from pathlib import Path

ROOT = Path(r"D:/BaiduSyncdisk/工作台_Codex")
IMAGES = ROOT / "work/p1/项目一_常用低压电器_ppt169_20260930/images"
SRC = ROOT / "work/p1/html_src"
OUT = ROOT / "knowledge/lesson/电气控制技术实训/2026-2027-1/项目一/教学互动"

PAT = re.compile(r"%%IMG:([^:%]+):(\d+)%%")


def data_uri(name: str, width: int) -> str:
    from PIL import Image
    p = IMAGES / name
    if not p.exists():
        raise FileNotFoundError(p)
    im = Image.open(p)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    w, h = im.size
    if w > width:
        im = im.resize((width, max(1, round(h * width / w))), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82, optimize=True, progressive=True)
    b = buf.getvalue()
    sys.stderr.write("  %-16s -> %4dx%-4d %6.1f KB\n" % (name, im.size[0], im.size[1], len(b) / 1024))
    return "data:image/jpeg;base64," + base64.b64encode(b).decode("ascii")


def main() -> int:
    check = "--check" in sys.argv
    tpls = sorted(SRC.glob("*.tpl.html"))
    if not tpls:
        print("没有找到模板：%s/*.tpl.html" % SRC)
        return 1
    OUT.mkdir(parents=True, exist_ok=True)
    bad = 0
    for tpl in tpls:
        text = tpl.read_text(encoding="utf-8")
        names = PAT.findall(text)
        if check:
            if names:
                print("[残留占位符] %s -> %s" % (tpl.name, names))
                bad += 1
            continue
        print("[构建] %s" % tpl.name)
        for fname, width in names:
            uri = data_uri(fname, int(width))
            # 注意：不要用 "%"-格式化构造替换串，%% 会被当成转义而少一个百分号
            text = text.replace("%%IMG:" + fname + ":" + width + "%%", uri)
        left = PAT.findall(text)
        if left:
            print("  !! 仍有未替换：%s" % left)
            bad += 1
        if "%%" in text:
            print("  !! 仍残留 %% 标记（替换串被转义吃掉）")
            bad += 1
        if text.count("data:image/jpeg;base64") != len(names):
            print("  !! 内嵌图片数量不符：期望 %d，实际 %d"
                  % (len(names), text.count("data:image/jpeg;base64")))
            bad += 1
        dst = OUT / tpl.name.replace(".tpl.html", ".html")
        dst.write_text(text, encoding="utf-8")
        print("  -> %s  (%.1f KB)" % (dst.name, dst.stat().st_size / 1024))
    if check:
        print("校验完成，问题 %d 个" % bad)
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
