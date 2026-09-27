from __future__ import annotations

import argparse
import json
import re
from difflib import SequenceMatcher
from pathlib import Path

try:
    import cv2
    from rapidocr_onnxruntime import RapidOCR
except ImportError as exc:
    raise SystemExit(
        "缺少依赖。请先安装 opencv-python 和 rapidocr_onnxruntime；不要在未获许可时自动安装。"
    ) from exc


def norm(text: str) -> str:
    return re.sub(r"[\s，。！？、,.!?~～…：:；;‘’“”\"'（）()【】\[\]]+", "", text).lower()


def stamp(seconds: float) -> str:
    hours = int(seconds // 3600)
    minutes = int(seconds % 3600 // 60)
    secs = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"


def similar(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return 1.0 if a == b else SequenceMatcher(None, a, b).ratio()


def main() -> None:
    parser = argparse.ArgumentParser(description="从聊天录屏抽帧 OCR，生成待人工复核的时间线。")
    parser.add_argument("video")
    parser.add_argument("output_md")
    parser.add_argument("output_json")
    parser.add_argument("--step", type=float, default=1.25, help="抽帧间隔（秒）")
    parser.add_argument("--scale", type=float, default=1.0, help="OCR 前缩放比例")
    parser.add_argument("--min-confidence", type=float, default=0.72)
    parser.add_argument("--left-max-ratio", type=float, default=0.43)
    parser.add_argument("--right-min-ratio", type=float, default=0.57)
    parser.add_argument("--dedupe-window", type=float, default=35.0)
    parser.add_argument("--left-label", default="左侧（未确认）")
    parser.add_argument("--right-label", default="右侧（未确认）")
    parser.add_argument("--title", default="聊天记录")
    args = parser.parse_args()

    if args.step <= 0 or args.scale <= 0:
        parser.error("--step 和 --scale 必须大于 0")
    if not 0 < args.left_max_ratio < args.right_min_ratio < 1:
        parser.error("说话人横向阈值必须满足 0 < left < right < 1")

    video = Path(args.video).resolve()
    if not video.is_file():
        raise SystemExit(f"找不到视频：{video}")
    output_md = Path(args.output_md).resolve()
    output_json = Path(args.output_json).resolve()
    output_md.parent.mkdir(parents=True, exist_ok=True)
    output_json.parent.mkdir(parents=True, exist_ok=True)

    cap = cv2.VideoCapture(str(video))
    if not cap.isOpened():
        raise SystemExit(f"无法打开视频：{video}")
    fps = float(cap.get(cv2.CAP_PROP_FPS))
    frame_count = float(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if fps <= 0 or frame_count <= 0:
        cap.release()
        raise SystemExit("无法读取有效的帧率或帧数")
    duration = frame_count / fps
    source_width = float(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    source_height = float(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    engine = RapidOCR()

    entries: list[dict] = []
    raw_frames: list[dict] = []
    t = 0.0
    sample_index = 0
    while t <= duration + 0.001:
        cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
        ok, frame = cap.read()
        if not ok:
            t += args.step
            continue
        if args.scale != 1.0:
            interpolation = cv2.INTER_CUBIC if args.scale > 1 else cv2.INTER_AREA
            frame = cv2.resize(frame, None, fx=args.scale, fy=args.scale, interpolation=interpolation)
        width = float(frame.shape[1])
        result, _ = engine(frame)
        lines: list[dict] = []
        for item in result or []:
            box, text, score = item
            text = str(text).strip()
            score = float(score)
            if not text or score < args.min_confidence:
                continue
            x1 = min(float(point[0]) for point in box)
            x2 = max(float(point[0]) for point in box)
            y1 = min(float(point[1]) for point in box)
            y2 = max(float(point[1]) for point in box)
            center_x = (x1 + x2) / 2
            if center_x < width * args.left_max_ratio:
                speaker = args.left_label
            elif center_x > width * args.right_min_ratio:
                speaker = args.right_label
            else:
                speaker = "系统/时间/居中内容"
            lines.append(
                {
                    "speaker": speaker,
                    "text": text,
                    "score": round(score, 3),
                    "box": [round(x1, 1), round(y1, 1), round(x2, 1), round(y2, 1)],
                    "y": y1,
                }
            )
        lines.sort(key=lambda item: (item["y"], item["box"][0]))
        raw_frames.append({"video_seconds": round(t, 2), "video_time": stamp(t), "lines": lines})

        for line in lines:
            key = norm(line["text"])
            if len(key) < 2:
                continue
            duplicate = False
            for prior in reversed(entries[-80:]):
                if t - prior["video_seconds"] > args.dedupe_window:
                    break
                if prior["speaker"] == line["speaker"] and similar(key, prior["norm"]) >= 0.94:
                    duplicate = True
                    break
            if not duplicate:
                entries.append(
                    {
                        "video_seconds": round(t, 2),
                        "video_time": stamp(t),
                        "speaker": line["speaker"],
                        "text": line["text"],
                        "confidence": line["score"],
                        "norm": key,
                    }
                )

        sample_index += 1
        if sample_index % 20 == 0:
            print(f"processed {stamp(t)} / {stamp(duration)}; entries={len(entries)}", flush=True)
        t += args.step
    cap.release()

    mapping = f"左侧＝{args.left_label}；右侧＝{args.right_label}；居中＝系统/时间/引用等待复核内容"
    md = [
        f"# {args.title}（视频 OCR 草稿）",
        "",
        f"- 来源视频：`{video.name}`",
        f"- 视频时长：{stamp(duration)}",
        f"- 源分辨率：{int(source_width)}×{int(source_height)}；帧率：{fps:.2f}",
        f"- 抽帧间隔：{args.step:.2f} 秒；OCR 缩放：{args.scale:.2f}",
        f"- 说话人映射：{mapping}",
        "- 说明：未读取音轨；已自动去除大量重复画面。OCR 可能有误字，关键原话必须对照原帧复核。",
        "",
        "## OCR 时间线",
        "",
    ]
    for entry in entries:
        md.append(f"- `[录屏 {entry['video_time']}]` **{entry['speaker']}**：{entry['text']}")
    output_md.write_text("\n".join(md) + "\n", encoding="utf-8")

    serializable_entries = []
    for entry in entries:
        clean = dict(entry)
        clean.pop("norm", None)
        serializable_entries.append(clean)
    output_json.write_text(
        json.dumps(
            {
                "video": str(video),
                "duration_seconds": round(duration, 2),
                "fps": round(fps, 3),
                "resolution": [int(source_width), int(source_height)],
                "step_seconds": args.step,
                "scale": args.scale,
                "mapping": {"left": args.left_label, "right": args.right_label},
                "entries": serializable_entries,
                "frames": raw_frames,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"done; entries={len(entries)}; md={output_md}; json={output_json}", flush=True)


if __name__ == "__main__":
    main()
