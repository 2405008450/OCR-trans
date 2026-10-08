"""本地 PaddleOCR：延迟加载模型、串行推理、分块渲染和逐页缓存。"""
from __future__ import annotations

import hashlib
import importlib.metadata
import json
import threading
import time
import subprocess
import tempfile
import sys
import os
from pathlib import Path
from typing import Any

from app.core.config import settings

MODEL_ID = "local/paddleocr"
_lock = threading.Lock()
_engine = None
_engine_key = None
_worker_lock = threading.Lock()


def _configuration():
    return {
        "device": settings.PADDLEOCR_DEVICE,
        "text_detection_model_name": settings.PADDLEOCR_DETECTION_MODEL,
        "text_recognition_model_name": settings.PADDLEOCR_RECOGNITION_MODEL,
        "cpu_threads": max(1, settings.PADDLEOCR_CPU_THREADS),
        "enable_mkldnn": False,
        "use_doc_orientation_classify": False,
        "use_doc_unwarping": False,
        "use_textline_orientation": True,
        "text_rec_score_thresh": 0.0,
        "text_det_limit_side_len": 1600,
        "text_det_limit_type": "max",
    }


def _get_engine(config):
    global _engine, _engine_key
    key = json.dumps(config, sort_keys=True)
    if _engine is None or _engine_key != key:
        try:
            os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "bos")
            os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
            from paddleocr import PaddleOCR
            _engine = PaddleOCR(**config)
        except Exception as exc:
            raise RuntimeError("PaddleOCR 初始化失败，请检查独立运行环境及模型配置：" + str(exc)) from exc
        _engine_key = key
    return _engine


def _same_region(a, b):
    """按位置消除分块重叠，保留不同位置出现的相同文字。"""
    x = max(0, min(a[2], b[2]) - max(a[0], b[0]))
    y = max(0, min(a[3], b[3]) - max(a[1], b[1]))
    area = min(max(0, a[2]-a[0]) * max(0, a[3]-a[1]),
               max(0, b[2]-b[0]) * max(0, b[3]-b[1]))
    return area > 0 and x * y / area >= 0.6


def _recognize_page(page, engine, dpi):
    import fitz
    import numpy as np
    scale = dpi / 72
    # 每块最大1600像素，避免A0图纸整页高分辨率渲染占满内存。
    size, step = 1600 / scale, 1400 / scale
    lines = []
    rect = page.rect
    # 普通图片整张识别，保留完整长句；超大图及PDF继续分块控制内存。
    if not page.parent.is_pdf and rect.width * rect.height * scale * scale <= 12_000_000:
        size = step = max(rect.width, rect.height) + 1
    y = rect.y0
    while y < rect.y1:
        x = rect.x0
        while x < rect.x1:
            clip = fitz.Rect(x, y, min(x+size, rect.x1), min(y+size, rect.y1))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False, colorspace=fitz.csRGB)
            # PaddleOCR接收BGR图像。每次仅保留当前块。
            array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)[:, :, ::-1].copy()
            for result in engine.predict(array):
                data = result.json
                if isinstance(data, str):
                    data = json.loads(data)
                data = data.get("res", data)
                for text, score, polygon in zip(data.get("rec_texts", []), data.get("rec_scores", []), data.get("rec_polys", [])):
                    if not text.strip():
                        continue
                    xs = [float(p[0]) / scale + pix.x / scale for p in polygon]
                    ys = [float(p[1]) / scale + pix.y / scale for p in polygon]
                    line = {"text": text, "score": float(score), "bbox": [min(xs), min(ys), max(xs), max(ys)]}
                    _merge_line(lines, line)
            x += step
        y += step
    return sorted(lines, key=lambda item: (round(item["bbox"][1] / 5), item["bbox"][0]))


def _merge_line(lines, line):
    """分块重叠优先保留完整文字框，避免高置信度截断片段覆盖长句。"""
    overlaps = [old for old in lines if _same_region(old["bbox"], line["bbox"])]
    if not overlaps:
        lines.append(line)
        return
    candidates = overlaps + [line]
    accepted = [item for item in candidates if item["score"] >= settings.PADDLEOCR_MIN_SCORE]
    pool = accepted or candidates
    def area(item):
        x0, y0, x1, y1 = item["bbox"]
        return max(0, x1-x0) * max(0, y1-y0)
    largest = max(area(item) for item in pool)
    # 大小接近的完整框仍按置信度选择，明显截断的框不参与覆盖。
    best = max((item for item in pool if area(item) >= largest * 0.85), key=lambda item:item["score"])
    for old in overlaps:
        lines.remove(old)
    lines.append(best)


def extract_paddle_plain_text(*, file_path, page_numbers=None, page_progress_callback=None, status_callback=None, **_):
    if not settings.PADDLEOCR_ENABLED:
        raise RuntimeError("PaddleOCR 尚未启用，请配置 PADDLEOCR_ENABLED=true 并安装识别依赖")
    if settings.PADDLEOCR_PYTHON:
        return _extract_in_worker(file_path, page_numbers, status_callback, page_progress_callback)
    import fitz
    config = _configuration()
    dpi = min(400, max(150, settings.PADDLEOCR_DPI))
    versions = {name: importlib.metadata.version(name) for name in ("paddleocr", "paddlex", "PyMuPDF")}
    min_score = min(1.0, max(0.0, settings.PADDLEOCR_MIN_SCORE))
    signature = json.dumps({"schema": 4, "config": config, "dpi": dpi, "min_score": min_score, "versions": versions}, sort_keys=True)
    digest = hashlib.sha256()
    with open(file_path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    cache = Path(settings.PADDLEOCR_CACHE_DIR) / digest.hexdigest() / hashlib.sha256(signature.encode()).hexdigest()[:20]
    cache.mkdir(parents=True, exist_ok=True)
    results = []
    with _lock, fitz.open(file_path) as document:
        numbers = list(page_numbers) if page_numbers is not None else list(range(1, len(document)+1))
        for index, number in enumerate(numbers, 1):
            if not 1 <= number <= len(document):
                raise ValueError(f"页码超出范围：{number}")
            if status_callback:
                status_callback(f"PaddleOCR 第 {number} 页（{index}/{len(numbers)}）")
            started = time.monotonic()
            cache_file = cache / f"{number}.json"
            try:
                saved = None
                if cache_file.exists():
                    try:
                        saved = json.loads(cache_file.read_text(encoding="utf-8"))
                    except (ValueError, OSError):
                        pass
                if saved is None:
                    lines = _recognize_page(document[number-1], _get_engine(config), dpi)
                    accepted = [line for line in lines if line["score"] >= min_score]
                    saved = {"page_number": number, "lines": lines, "text": "\n".join(line["text"] for line in accepted),
                             "raw_text": "\n".join(line["text"] for line in lines),
                             "excluded_low_score_lines": [line for line in lines if line["score"] < min_score],
                             "error": "", "blank": False, "elapsed_seconds": round(time.monotonic()-started, 3)}
                    # 没有识别结果不证明空白，必须保留复核提示。
                    saved["review_required"] = not accepted or any(line["score"] < 0.85 for line in lines)
                    temp = cache_file.with_suffix(".tmp")
                    temp.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")
                    temp.replace(cache_file)
                results.append(saved)
            except Exception as exc:
                results.append({"page_number": number, "text": "", "error": str(exc), "blank": False})
            if page_progress_callback:
                page_progress_callback(index, len(numbers))
        total = len(document)
    return {"text": "\n\n".join(row["text"] for row in results if not row.get("error")),
            "page_results": results, "total_pages": total,
            "processed_pages": [row["page_number"] for row in results],
            "failed_pages": [row["page_number"] for row in results if row.get("error")],
            "review_pages": [row["page_number"] for row in results if row.get("review_required")],
            "cache_directory": str(cache)}


def _extract_in_worker(file_path, page_numbers, status_callback, progress_callback):
    """独立Python环境运行OCR，避免模型依赖影响Web服务。"""
    root = Path(__file__).resolve().parents[2]
    config = {key: getattr(settings, key) for key in (
        "PADDLEOCR_DEVICE", "PADDLEOCR_DETECTION_MODEL", "PADDLEOCR_RECOGNITION_MODEL",
        "PADDLEOCR_DPI", "PADDLEOCR_CPU_THREADS", "PADDLEOCR_CACHE_DIR", "PADDLEOCR_MIN_SCORE")}
    if status_callback:
        status_callback("正在使用独立 PaddleOCR 环境识别；首次运行需要加载模型")
    with _worker_lock, tempfile.TemporaryDirectory(prefix="paddleocr-") as folder:
        request = Path(folder) / "request.json"
        response = Path(folder) / "response.json"
        request.write_text(json.dumps({"file_path": str(Path(file_path).resolve()),
                                      "page_numbers": list(page_numbers) if page_numbers is not None else None,
                                      "config": config}, ensure_ascii=False), encoding="utf-8")
        try:
            process = subprocess.run([settings.PADDLEOCR_PYTHON, "-m", "app.service.paddle_ocr_service",
                                      str(request), str(response)], cwd=root, capture_output=True,
                                     timeout=max(30, settings.PADDLEOCR_TIMEOUT_SECONDS))
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("PaddleOCR超时；已保存的逐页缓存可在重试时复用") from exc
        if process.returncode or not response.exists():
            error = (process.stderr or process.stdout).decode("utf-8", errors="replace")[-2000:]
            raise RuntimeError("PaddleOCR独立环境执行失败：" + error)
        result = json.loads(response.read_text(encoding="utf-8"))
        if progress_callback:
            total = len(result.get("processed_pages") or [])
            progress_callback(total, total)
        return result


if __name__ == "__main__":
    request = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    settings.PADDLEOCR_ENABLED = True
    settings.PADDLEOCR_PYTHON = ""
    for key, value in request["config"].items():
        setattr(settings, key, value)
    result = extract_paddle_plain_text(file_path=request["file_path"], page_numbers=request["page_numbers"])
    Path(sys.argv[2]).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
