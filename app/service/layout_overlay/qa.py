"""有界质检：检查、渲染比对、白名单修补，禁止模型执行任意代码。"""
from __future__ import annotations

import io
import math
import re
import subprocess
import tempfile
import time
from pathlib import Path

import fitz
from PIL import Image

from app.core.config import settings
from app.service.libreoffice_service import resolve_libreoffice_path
from .imaging import fit_segment
from .models import overlaps, validate_bbox
from .ocr_providers import OcrRouter
from .translation import call_vision, translate_segments
from .budget import bounded_call

ALLOWED_ACTIONS = {"retranslate", "set_font_scale", "nudge_bbox", "re_erase", "mark_protected", "re_ocr"}


def _recognize(image_bytes, primary, fallbacks, timeout, allow_empty=False):
    return OcrRouter(primary, fallbacks).recognize(image_bytes, timeout=timeout, allow_empty=allow_empty)


def _shorter_translation(segment, options, timeout):
    translate_segments([segment], options["source_lang"], options["target_lang"], options["translation_engine"],
                       options["gemini_route"], shorter=True, timeout=timeout)
    return segment.translation


def deterministic_issues(pages, source_lang, target_lang):
    issues = []
    for page in pages:
        for segment in page.segments:
            if segment.protected:
                continue
            if segment.overflow:
                issues.append({"id": segment.segment_id, "kind": "overflow", "action": "retranslate", "shorter": True})
            text = segment.translation
            residue = False
            if source_lang.startswith("zh") and not target_lang.startswith("zh") and target_lang != "ja":
                residue = bool(re.search(r"[\u3400-\u9fff]", text))
            elif source_lang == "ja" and target_lang != "ja":
                residue = bool(re.search(r"[\u3040-\u30ff]", text))
            elif source_lang == "ko" and target_lang != "ko":
                residue = bool(re.search(r"[\uac00-\ud7af]", text))
            elif source_lang in {"ru", "uk", "bg"} and target_lang not in {"ru", "uk", "bg", "sr", "kk", "mn"}:
                residue = bool(re.search(r"[\u0400-\u04ff]", text))
            if residue:
                issues.append({"id": segment.segment_id, "kind": "source_language", "action": "retranslate", "shorter": True})
    return issues


def residual_issues(page, router, timeout=None):
    duration = timeout or settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS
    lines = bounded_call(_recognize, page.clean_path.read_bytes(), router.order[0], router.order[1:], duration, True,
                         budget_seconds=duration)
    issues = []
    for segment in page.segments:
        if segment.protected or segment.translation.strip() == segment.text.strip():
            continue
        for line in lines:
            if (len(line.text.strip()) >= 2 and line.text.strip() in segment.text and
                    any(overlaps(line.bbox, source.bbox) for source in segment.lines)):
                issues.append({"id": segment.segment_id, "kind": "residual", "action": "re_erase"})
                break
    return issues


def render_docx(docx_path, output_dir, timeout):
    """复用可执行文件发现；单独设置超时，避免质检挂住整个任务队列。"""
    executable = resolve_libreoffice_path()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    expected = output_dir / (Path(docx_path).stem + ".pdf")
    expected.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(prefix="overlay-lo-", ignore_cleanup_errors=True) as profile:
        result = subprocess.run([executable, f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
                                 "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir",
                                 str(output_dir), str(Path(docx_path).resolve())],
                                capture_output=True, timeout=max(1, min(45, timeout)),
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if result.returncode or not expected.is_file():
        raise RuntimeError("LibreOffice 未生成质检 PDF")
    return expected


def visual_issues(pages, pdf_path, model, route, timeout):
    issues = []
    deadline = time.monotonic() + timeout
    with fitz.open(pdf_path) as document:
        if len(document) != len(pages):
            raise ValueError(f"渲染页数不一致：预期 {len(pages)}，实际 {len(document)}")
        for index, page in enumerate(pages):
            remaining = deadline - time.monotonic()
            if remaining < 1:
                raise TimeoutError("视觉质检预算耗尽")
            with Image.open(page.image_path) as image:
                original = image.convert("RGB")
            pixmap = document[index].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            rendered = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
            original.thumbnail((1300, 1800))
            rendered = rendered.resize(original.size)
            comparison = Image.new("RGB", (original.width * 2, original.height), "white")
            comparison.paste(original, (0, 0))
            comparison.paste(rendered, (original.width, 0))
            data = io.BytesIO()
            comparison.save(data, "PNG")
            comparison.save(Path(pdf_path).parent / f"comparison-{index + 1}.png")
            mapping = [{"id": s.segment_id, "source": s.text, "translation": s.translation,
                        "bbox": s.bbox, "protected": s.protected} for s in page.segments]
            import json
            prompt = (
                "左边原件，右边 Word 实际渲染。检查文字溢出、漏译、原文残影、照片印章或签名变化。"
                "只针对给定 id 输出；无问题返回空数组。操作白名单："
                "retranslate(id,shorter=true)、set_font_scale(id,s=0.8..1.2)、"
                "nudge_bbox(id,dx,dy,dw)、re_erase(id)、mark_protected(id)、re_ocr(id,provider)。"
                '返回 {"issues":[{"id":"p1s1","kind":"overflow","action":"retranslate","shorter":true}]}。'
                "位移为原图像素，仅建议小幅调整。分段：" + json.dumps(mapping, ensure_ascii=False))
            duration = min(remaining, settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS)
            result = bounded_call(call_vision, data.getvalue(), prompt, model, route, duration, budget_seconds=duration)
            if not isinstance(result, dict) or not isinstance(result.get("issues"), list):
                raise ValueError("视觉质检响应格式错误")
            page_ids = {s.segment_id for s in page.segments}
            if any(not isinstance(issue, dict) or issue.get("id") not in page_ids for issue in result["issues"]):
                raise ValueError("视觉质检包含未知分段")
            issues.extend(result["issues"][:50])
    return issues


def apply_patch_action(issue, page, *, retranslate, reocr):
    """校验后执行单次修补；无效提议不改变当前状态。"""
    action = issue.get("action")
    segment = next((s for s in page.segments if s.segment_id == issue.get("id")), None)
    if action not in ALLOWED_ACTIONS or segment is None:
        raise ValueError("修补操作或分段 id 不在白名单")
    if segment.protected and action != "mark_protected":
        raise ValueError("禁止修补保护分段")
    if action == "retranslate":
        if issue.get("shorter", True) is not True:
            raise ValueError("retranslate 只允许请求简洁等义译文")
        retranslate(segment)
        fit_segment(segment, page)
    elif action == "set_font_scale":
        scale = float(issue.get("s"))
        if not math.isfinite(scale) or not 0.8 <= scale <= 1.2:
            raise ValueError("字号缩放超出 0.8..1.2")
        segment.font_scale = scale
        fit_segment(segment, page)
    elif action == "nudge_bbox":
        dx, dy, dw = (float(issue.get(key, 0)) for key in ("dx", "dy", "dw"))
        limit = min(30, (segment.bbox[3] - segment.bbox[1]) * 0.3)
        if not all(math.isfinite(v) and abs(v) <= limit for v in (dx, dy, dw)):
            raise ValueError("文本框移动超过允许范围")
        x1, y1, x2, y2 = segment.bbox
        candidate = (x1 + dx, y1 + dy, x2 + dx + dw, y2 + dy)
        if not (0 <= candidate[0] < candidate[2] <= page.width and 0 <= candidate[1] < candidate[3] <= page.height):
            raise ValueError("修补文本框越界")
        validate_bbox(candidate, page.width, page.height)
        if any(overlaps(candidate, r["bbox"]) for r in page.protected_regions) or any(
                other is not segment and overlaps(candidate, other.bbox) for other in page.segments):
            raise ValueError("修补文本框与其他区域重叠")
        segment.bbox = candidate
        fit_segment(segment, page)
    elif action == "re_erase":
        segment.erase_padding = min(5, segment.erase_padding + 1)
    elif action == "mark_protected":
        segment.kind = "no_translate"
    elif action == "re_ocr":
        provider = issue.get("provider")
        if provider not in {"google", "qwen", "glm"}:
            raise ValueError("交叉识别引擎无效")
        reocr(segment, page, provider)
        retranslate(segment)
        fit_segment(segment, page)


def run_qa(pages, docx_path, debug_dir, options, rebuild, notify=None):
    enabled = options["enable_qa"]
    report = {"status": "disabled" if not enabled else "partial", "rounds": [], "warnings": [], "cross_ocr": [],
              "issues": deterministic_issues(pages, options["source_lang"], options["target_lang"]),
              "visual_checked": False}
    if not enabled:
        return report
    deadline = time.monotonic() + max(1, settings.LAYOUT_OVERLAY_QA_TIMEOUT_SECONDS)
    max_rounds = max(0, min(2, settings.LAYOUT_OVERLAY_MAX_QA_ROUNDS))

    def remaining():
        duration = deadline - time.monotonic()
        if duration < 1:
            raise TimeoutError("质检总预算耗尽")
        return min(duration, settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS)

    def shorter(segment):
        duration = remaining()
        segment.translation = bounded_call(_shorter_translation, segment, options, duration, budget_seconds=duration)

    def reocr(segment, page, provider):
        # 仅识别源文字的原始区域，避免修补后的文本框坐标改变识别目标。
        from .models import union_bbox
        box = union_bbox([line.bbox for line in segment.lines])
        with Image.open(page.image_path) as image:
            crop = image.crop(tuple(map(int, box)))
            data = io.BytesIO()
            crop.save(data, "PNG")
        duration = remaining()
        lines = bounded_call(_recognize, data.getvalue(), provider, [], duration, budget_seconds=duration)
        segment.text = "\n".join(line.text for line in lines)

    def inspect(round_number):
        issues = deterministic_issues(pages, options["source_lang"], options["target_lang"])
        complete = True
        for page in pages:
            if notify:
                notify(78 + round_number * 4, f"质检第 {round_number + 1} 次检查：底图残影识别")
            try:
                issues.extend(residual_issues(page, OcrRouter(options["ocr_provider"]), timeout=remaining()))
            except Exception as exc:
                complete = False
                report["warnings"].append(f"底图残影识别未完成：{type(exc).__name__}")
        if notify:
            notify(79 + round_number * 4, f"质检第 {round_number + 1} 次检查：Word 渲染与视觉比对")
        try:
            pdf = render_docx(docx_path, Path(debug_dir) / f"qa-{round_number}", remaining())
            issues.extend(visual_issues(pages, pdf, options["vision_model"], options["gemini_route"], remaining()))
            report["visual_checked"] = True
        except Exception as exc:
            complete = False
            report["warnings"].append(f"Word 视觉质检未完成：{type(exc).__name__}")
        return issues, complete

    # 低置信度分段裁剪后换引擎复核，仅记录差异，不自动改写证件事实。
    cross_count = 0
    cross_incomplete = False
    for page in pages:
        for segment in page.segments:
            if segment.protected or not any(l.confidence is None or l.confidence < settings.LAYOUT_OVERLAY_MIN_OCR_CONFIDENCE for l in segment.lines):
                continue
            if cross_count >= 10 or time.monotonic() >= deadline:
                cross_incomplete = True
                continue
            cross_count += 1
            provider = next((p for p in OcrRouter(options["ocr_provider"]).order if p != page.ocr_provider), None)
            if not provider:
                cross_incomplete = True
                continue
            if notify:
                notify(76, f"低置信度文字交叉复核：{segment.segment_id}")
            try:
                import copy
                candidate = copy.deepcopy(segment)
                reocr(candidate, page, provider)
                same = re.sub(r"\s+", "", candidate.text) == re.sub(r"\s+", "", segment.text)
                report["cross_ocr"].append({"id": segment.segment_id, "provider": provider,
                                            "text": candidate.text, "matches": same})
                if not same:
                    report["warnings"].append(f"{segment.segment_id} 两个 OCR 引擎识别不一致，请核对原件")
            except Exception as exc:
                cross_incomplete = True
                report["warnings"].append(f"{segment.segment_id} 交叉识别未完成：{type(exc).__name__}")

    import copy
    best = None
    best_score = float("inf")

    for number in range(max_rounds + 1):
        issues, complete = inspect(number)
        score = len(issues) + (0 if complete else 1000)
        if score < best_score:
            best_score = score
            best = (copy.deepcopy([p.segments for p in pages]), Path(docx_path).read_bytes(),
                    [p.clean_path.read_bytes() for p in pages], copy.deepcopy(issues), complete, number)
        report["issues"] = issues
        report["rounds"].append({"round": number, "issues": issues, "patches": []})
        if not issues or number == max_rounds or time.monotonic() >= deadline:
            report["status"] = "passed" if complete and not issues else "needs_review" if issues else "partial"
            break
        # 每轮最多 20 个动作，同一分段的同类动作最多一次；每步回滚失败状态。
        seen = set()
        for issue in issues[:20]:
            key = (issue.get("id"), issue.get("action"))
            if key in seen:
                continue
            seen.add(key)
            page = next((p for p in pages if any(s.segment_id == issue.get("id") for s in p.segments)), None)
            if not page:
                continue
            original = copy.deepcopy(page.segments)
            try:
                remaining()
                apply_patch_action(issue, page, retranslate=shorter, reocr=reocr)
                report["rounds"][-1]["patches"].append({"id": key[0], "action": key[1], "applied": True})
            except Exception as exc:
                page.segments = original
                report["rounds"][-1]["patches"].append({"id": key[0], "action": key[1], "applied": False,
                                                      "error": type(exc).__name__})
        try:
            rebuild()
        except Exception as exc:
            report["warnings"].append(f"修补后构建失败：{type(exc).__name__}")
            break
    if best:
        states, docx, backgrounds, issues, complete, best_round = best
        for page, segments, background in zip(pages, states, backgrounds):
            page.segments = segments
            page.clean_path.write_bytes(background)
        Path(docx_path).write_bytes(docx)
        report["issues"] = issues
        report["selected_round"] = best_round
        uncertain = any(not item["matches"] for item in report["cross_ocr"])
        report["status"] = "needs_review" if issues or uncertain else "passed" if complete and not cross_incomplete else "partial"
    return report
