from __future__ import annotations

import argparse
import re
from pathlib import Path

try:
    import cv2
except ImportError as exc:
    raise SystemExit("缺少依赖 opencv-python；不要在未获许可时自动安装。") from exc


def parse_time(value: str) -> float:
    value = value.strip()
    if re.fullmatch(r"\d+(?:\.\d+)?", value):
        return float(value)
    parts = value.split(":")
    if len(parts) not in (2, 3):
        raise argparse.ArgumentTypeError(f"无法解析时间：{value}")
    try:
        numbers = [float(part) for part in parts]
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"无法解析时间：{value}") from exc
    if len(numbers) == 2:
        minutes, seconds = numbers
        return minutes * 60 + seconds
    hours, minutes, seconds = numbers
    return hours * 3600 + minutes * 60 + seconds


def file_stamp(seconds: float) -> str:
    total_ms = int(round(seconds * 1000))
    hours, rem = divmod(total_ms, 3_600_000)
    minutes, rem = divmod(rem, 60_000)
    secs, millis = divmod(rem, 1000)
    return f"{hours:02d}-{minutes:02d}-{secs:02d}-{millis:03d}"


def main() -> None:
    parser = argparse.ArgumentParser(description="从视频指定时间点导出关键帧。")
    parser.add_argument("video")
    parser.add_argument("output_dir")
    parser.add_argument("--at", nargs="+", required=True, type=parse_time, metavar="TIME")
    args = parser.parse_args()

    video = Path(args.video).resolve()
    if not video.is_file():
        raise SystemExit(f"找不到视频：{video}")
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    cap = cv2.VideoCapture(str(video))
    if not cap.isOpened():
        raise SystemExit(f"无法打开视频：{video}")
    fps = float(cap.get(cv2.CAP_PROP_FPS))
    frame_count = float(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    duration = frame_count / fps if fps > 0 else 0.0

    written = 0
    for seconds in args.at:
        if seconds < 0 or (duration and seconds > duration + 0.5):
            print(f"skip out-of-range timestamp: {seconds:.3f}s", flush=True)
            continue
        cap.set(cv2.CAP_PROP_POS_MSEC, seconds * 1000)
        ok, frame = cap.read()
        if not ok:
            print(f"failed to read frame at {seconds:.3f}s", flush=True)
            continue
        output = output_dir / f"frame_{file_stamp(seconds)}.png"
        if not cv2.imwrite(str(output), frame):
            print(f"failed to write: {output}", flush=True)
            continue
        written += 1
        print(output, flush=True)
    cap.release()
    if not written:
        raise SystemExit("没有成功导出关键帧")


if __name__ == "__main__":
    main()
