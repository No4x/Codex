#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 olefile 解析 pptx 里嵌入对象的 OLE 复合文档，读出 OLE10Native 头里的原始文件名与大小。"""
import pathlib
import struct
import zipfile

import olefile

PPTX = pathlib.Path(
    r"D:/BaiduSyncdisk/工作台_Codex/knowledge/lesson/电气控制技术实训/2026-2027-1/项目一/ppt/"
    r"项目一_常用低压电器_v2.pptx"
)


def parse_native(blob: bytes):
    """解析 \x01Ole10Native 流：DWORD 总长 | WORD 标志 | [DWORD 版本] | 标签 | 原始路径 | ... | 数据"""
    off = 0
    total = struct.unpack_from("<I", blob, off)[0]
    off += 4
    off += 2                      # 0x0002
    # 跳过可选 OLE 版本头（有的写 0x0003 + 4 字节）
    ver = struct.unpack_from("<I", blob, off)[0]
    if ver == 3:
        off += 4

    def read_cstr():
        nonlocal off
        end = blob.index(b"\x00", off)
        s = blob[off:end]
        off = end + 1
        return s

    label = read_cstr()
    path = read_cstr()
    # 头部剩余字段各版本略有差异，直接以「数据段起点」定位，比按偏移硬算可靠
    start = blob.index(b"<!DOCTYPE html>")
    size = struct.unpack_from("<I", blob, start - 4)[0]
    data = blob[start:start + size]
    return total, label, path, size, data


def main() -> None:
    z = zipfile.ZipFile(PPTX)
    bins = sorted(
        (n for n in z.namelist() if n.startswith("ppt/embeddings/")),
        key=lambda x: int("".join(c for c in x if c.isdigit())),
    )
    print("嵌入对象 OLE10Native 头（决定双击后释放成什么文件名）：\n")
    for n in bins:
        ole = olefile.OleFileIO(z.read(n))
        streams = ["/".join(s) for s in ole.listdir()]
        native = None
        for s in streams:
            if "Ole10Native" in s:
                native = ole.openstream(s).read()
        if native is None:
            print("  %-26s 流=%s（无 Ole10Native）" % (n, streams))
            ole.close()
            continue
        total, label, path, size, data = parse_native(native)
        print("  %-26s 释放名=%s" % (n, path.decode("mbcs", "ignore")))
        print("     标签=%s 大小=%d B（=源 %d B: %s）"
              % (label.decode("mbcs", "ignore"), size, size, size == len(data)))
        ole.close()


if __name__ == "__main__":
    main()
