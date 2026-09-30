#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
嵌入对象保真校验：证明 ppt/embeddings/oleObjectN.bin 里的 HTML 与源文件**逐字节一致**。

原理
----
PowerPoint 存 OLE 对象时，把源文件**原样**封进一个 OLE 复合文档（.bin）。
因此可以：
  1) 用 olefile 打开 .bin，取出内嵌流；
  2) 直接把「源 HTML 字节」拿去 .bin 里做子串查找 —— 命中即证明原样保存。

⚠️ 关键坑：原始文件名**不在 OLE 目录项里**（目录项只有 `\x01Ole10Native` 之类），
   而是存在 `\x01Ole10Native` 流内部，且是 **ANSI = mbcs / GBK 编码**，不是 UTF-8。
   用 UTF-8 正则去匹配 `.html` 会得 0 命中（假失败）。
   必须 `bytes.decode("mbcs")` 之后再 `endswith(".html")`。

`\x01Ole10Native` 流布局（小端）：
   DWORD 总长度 ｜ WORD 标志(0x0002) ｜ ANSI\\0 标签(=文件名)
   ｜ ANSI\\0 原始完整路径 ｜ DWORD ｜ DWORD ｜ ANSI\\0 临时路径 ｜ 原始文件字节…

用法
----
  python verify_embeddings.py --pptx <文件.pptx> --html-dir <HTML 根目录> [--expect 16]

  退出码 0 = 全部通过；1 = 有失败项。
"""
from __future__ import annotations

import argparse
import hashlib
import pathlib
import struct
import sys
import zipfile

try:
    import olefile
except ImportError:
    olefile = None


NATIVE_STREAM = "\x01Ole10Native"


def iter_html(html_dir: pathlib.Path):
    yield from sorted(html_dir.rglob("*.html"))


def _read_cstr(data: bytes, pos: int) -> tuple[str, int]:
    """读一个以 \\0 结尾的 ANSI(mbcs) 字符串，返回 (文本, 新位置)。"""
    end = data.index(b"\x00", pos)
    return data[pos:end].decode("mbcs", errors="replace"), end + 1


def parse_ole10native(payload: bytes) -> dict[str, str]:
    """从 \\x01Ole10Native 流里取出 label / 原始路径（ANSI = mbcs 编码）。"""
    out: dict[str, str] = {}
    pos = 4                                    # 跳过 DWORD 总长度
    if payload[pos:pos + 2] == b"\x02\x00":    # WORD 标志
        pos += 2
    out["label"], pos = _read_cstr(payload, pos)
    out["original_path"], pos = _read_cstr(payload, pos)
    pos += 8                                   # 两个 DWORD
    out["temp_path"], pos = _read_cstr(payload, pos)
    return out


def ole_probe(data: bytes) -> dict[str, str]:
    """把 .bin 字节写临时文件 → olefile 读 → 解析 Ole10Native。"""
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".bin", delete=False) as f:
        f.write(data)
        tmp = pathlib.Path(f.name)
    try:
        ole = olefile.OleFileIO(str(tmp))
        try:
            names = ["/".join(e) for e in ole.listdir(streams=True, storages=True)]
            info: dict[str, str] = {"streams": ";".join(names)}
            for entry in ole.listdir(streams=True):
                if entry and entry[-1] == NATIVE_STREAM:
                    raw = ole.openstream(entry).read()
                    info.update(parse_ole10native(raw))
                    break
            return info
        finally:
            ole.close()
    finally:
        tmp.unlink(missing_ok=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--html-dir", required=True, help="源 HTML 根目录（递归）")
    ap.add_argument("--expect", type=int, default=0, help="期望的嵌入对象数量；0 = 不校验数量")
    args = ap.parse_args()

    pptx = pathlib.Path(args.pptx)
    html_dir = pathlib.Path(args.html_dir)
    if not pptx.exists():
        print("PPTX 不存在：%s" % pptx)
        return 1

    sources = {p.name: p for p in iter_html(html_dir)}
    src_bytes = {n: p.read_bytes() for n, p in sources.items()}
    print("源 HTML %d 个：%s" % (len(sources), "、".join(sorted(sources))))

    with zipfile.ZipFile(pptx) as z:
        bins = sorted(n for n in z.namelist()
                      if n.startswith("ppt/embeddings/") and n.lower().endswith(".bin"))
        print("嵌入对象 %d 个" % len(bins))
        if args.expect and len(bins) != args.expect:
            print("!! 数量不符：期望 %d，实际 %d" % (args.expect, len(bins)))
            return 1

        hit, miss, warn = [], [], []
        for name in bins:
            data = z.read(name)
            matched = [n for n, b in src_bytes.items() if b in data]
            if not matched:
                miss.append((name, "无源 HTML 字节命中"))
                print("  XX %-42s 未命中" % name)
                continue
            # 校验 OLE 包内的原始文件名（ANSI = mbcs 编码）
            ext_ok, fname = False, ""
            try:
                info = ole_probe(data)
                fname = info.get("label") or info.get("original_path") or ""
                ext_ok = pathlib.Path(fname).name.lower().endswith(".html")
            except Exception as e:                 # noqa: BLE001
                fname = "(解析失败: %s)" % e
            tag = "OK" if ext_ok else "WARN"
            hit.append(name)
            print("  %-28s %-5s 包内文件名=%s" % (name, tag, fname))
            print("      sha256=%s" % hashlib.sha256(data).hexdigest()[:12])
            for n in matched:
                print("      <- %s（%d 字节，逐字节命中）" % (n, len(src_bytes[n])))
            if not ext_ok:
                warn.append(name)

    print("\n命中 %d / %d ｜ 扩展名保留 %d / %d" % (len(hit), len(bins), len(bins) - len(warn), len(bins)))
    for name, why in miss:
        print("  - %s: %s" % (name, why))
    return 1 if miss else 0


if __name__ == "__main__":
    raise SystemExit(main())
