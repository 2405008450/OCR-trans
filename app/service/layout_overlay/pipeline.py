"""PDF/图片 → 原坐标分段 → 覆盖译文 DOCX 的独立任务入口。"""
from __future__ import annotations

import asyncio
import functools
import io
import json
import zipfile
from dataclasses import asdict
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageOps

from app.core.config import settings
from app.service.doc_translate_service import (
    DOC_TRANSLATE_DEFAULT_MODEL, DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE,
    DOC_TRANSLATE_TRANSLATION_ENGINES, SUPPORTED_LANGUAGES,
    get_doc_translate_models, get_doc_translate_translation_engines,
)
from app.service.gemini_service import get_gemini_routes
from .classification import classify_page
from .docx_builder import build_docx
from .imaging import erase_page, estimate_style, fit_segment
from .models import OverlayPage
from .ocr_providers import OcrRouter, PROVIDERS, extract_pdf_text_lines
from .qa import run_qa
from .translation import translate_segments
from .editable import extract_document, translate_documents, build_editable_docx, review_editable

ALLOWED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}


def get_layout_overlay_config():
    routes = {name: dict(info) for name, info in get_gemini_routes().items()}
    for name, label in {"google": "Google Vertex（需服务账号）", "aistudio": "Google AI Studio（API 密钥）",
                        "openrouter": "OpenRouter（API 密钥）"}.items():
        routes[name]["label"] = label
    models = {name: dict(info, label=f"{name.split('/')[-1]}（{info['label']}）")
              for name, info in get_doc_translate_models().items()}
    engines = {name: dict(info) for name, info in get_doc_translate_translation_engines().items()}
    for name, info in engines.items():
        if name.startswith("google/"):
            info["label"] = f"{name.split('/')[-1]}（{info['label']}）"
    return {"ocr_providers": {"google": "Google Vision", "qwen": "Qwen-VL-OCR", "glm": "GLM-OCR"},
            "default_ocr_provider": settings.LAYOUT_OVERLAY_OCR_PROVIDER,
            "output_modes": {"editable": "清晰可编辑排版（推荐）", "overlay": "原坐标覆盖（保留底纹）"},
            "default_output_mode": "editable",
            "models": models, "default_model": DOC_TRANSLATE_DEFAULT_MODEL,
            "translation_engines": engines,
            "default_translation_engine": DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE,
            "routes": routes, "default_route": settings.LAYOUT_OVERLAY_GEMINI_ROUTE,
            "languages": {code: info["name"] for code, info in SUPPORTED_LANGUAGES.items()},
            "allowed_extensions": sorted(ALLOWED_EXTENSIONS),
            "upload_max_mb": settings.LAYOUT_OVERLAY_UPLOAD_MAX_MB,
            "max_pages": settings.LAYOUT_OVERLAY_MAX_PAGES,
            "max_qa_rounds": min(2, max(0, settings.LAYOUT_OVERLAY_MAX_QA_ROUNDS))}


def normalize_options(*, source_lang="zh", target_lang="en", ocr_provider=None, vision_model=DOC_TRANSLATE_DEFAULT_MODEL,
                      gemini_route=None, translation_engine=DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE, enable_qa=True,
                      output_mode="editable"):
    ocr_provider = ocr_provider or settings.LAYOUT_OVERLAY_OCR_PROVIDER
    gemini_route = gemini_route or settings.LAYOUT_OVERLAY_GEMINI_ROUTE
    if source_lang not in SUPPORTED_LANGUAGES or target_lang not in SUPPORTED_LANGUAGES:
        raise ValueError("不支持的源语言或目标语言")
    if source_lang == target_lang:
        raise ValueError("源语言和目标语言不能相同")
    if ocr_provider not in PROVIDERS:
        raise ValueError("不支持的 OCR 引擎")
    if translation_engine not in DOC_TRANSLATE_TRANSLATION_ENGINES:
        raise ValueError("不支持的翻译模型")
    if vision_model not in get_doc_translate_models():
        raise ValueError("不支持的视觉模型")
    if gemini_route not in get_gemini_routes():
        raise ValueError("不支持的模型路由")
    if type(enable_qa) is not bool:
        raise ValueError("enable_qa 必须为布尔值")
    if output_mode not in {"editable", "overlay"}:
        raise ValueError("不支持的导出排版方式")
    return dict(source_lang=source_lang, target_lang=target_lang, ocr_provider=ocr_provider,
                vision_model=vision_model, gemini_route=gemini_route,
                translation_engine=translation_engine, enable_qa=enable_qa, output_mode=output_mode)


def render_input(input_path, debug_dir, router, notify):
    source = Path(input_path)
    pages = []
    dpi = max(72, min(400, settings.LAYOUT_OVERLAY_RENDER_DPI))
    maximum = max(1, settings.LAYOUT_OVERLAY_MAX_PAGES)
    pixel_limit = max(1, settings.LAYOUT_OVERLAY_MAX_PAGE_PIXELS)

    def add_page(image, width_pt, height_pt, text_page=None):
        if image.width * image.height > pixel_limit:
            raise ValueError("页面像素过多，请缩小图片或降低渲染 DPI")
        index = len(pages)
        image_path = Path(debug_dir) / f"original-{index + 1}.png"
        image.save(image_path, "PNG")
        lines = extract_pdf_text_lines(text_page, image.width, image.height) if text_page is not None else []
        provider = "pdf_text"
        if not lines:
            notify(10 + int(index / maximum * 15), f"第 {index + 1} 页：云端 OCR 识别")
            lines = router.recognize(image_path.read_bytes())
            provider = router.last_provider
        pages.append(OverlayPage(image_path, image.width, image.height, width_pt, height_pt, lines, ocr_provider=provider))

    if source.suffix.lower() == ".pdf":
        with fitz.open(source) as document:
            if document.needs_pass:
                raise ValueError("PDF 已加密，请先解除密码")
            if not 0 < len(document) <= maximum:
                raise ValueError(f"证件页数必须为 1..{maximum}")
            for page in document:
                if page.rect.width * page.rect.height * (dpi / 72) ** 2 > pixel_limit:
                    raise ValueError("PDF 页面尺寸超过像素上限")
                pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
                add_page(Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB"), page.rect.width, page.rect.height, page)
    else:
        with Image.open(source) as image:
            frames = getattr(image, "n_frames", 1)
            if frames > maximum:
                raise ValueError(f"图片页数不能超过 {maximum}")
            for index in range(frames):
                image.seek(index)
                oriented = ImageOps.exif_transpose(image).copy()
                # 透明图片以白底合成，避免透明区域变黑后参与文字擦除。
                rgba = oriented.convert("RGBA")
                background = Image.new("RGBA", rgba.size, "white")
                background.alpha_composite(rgba)
                rgb = background.convert("RGB")
                add_page(rgb, rgb.width / dpi * 72, rgb.height / dpi * 72)
    return pages


def _run_pipeline(*, task_id, display_no, input_path, original_filename, options, notify):
    output_dir = Path(settings.OUTPUT_DIR) / "layout_overlay" / (display_no or task_id)
    output_dir.mkdir(parents=True, exist_ok=True)
    debug_dir = output_dir / "debug"
    debug_dir.mkdir(exist_ok=True)
    output_path = output_dir / "layout_overlay.docx"
    router = OcrRouter(options["ocr_provider"])
    notify(5, "正在读取原件与提取文字坐标")
    pages = render_input(input_path, debug_dir, router, notify)
    if options.get("output_mode", "editable") == "editable":
        return _run_editable(pages, output_dir, debug_dir, options, router, original_filename, notify)
    warnings = []
    for index, page in enumerate(pages):
        notify(25 + int(index / len(pages) * 15), f"第 {index + 1} 页：分析语义块及照片、印章保护区")
        classify_page(page, index, options["vision_model"], options["gemini_route"])
        original = np.array(Image.open(page.image_path).convert("RGB"))
        for segment in page.segments:
            estimate_style(original, segment, page)
            uncertain = [line for line in segment.lines if line.confidence is None or line.confidence < settings.LAYOUT_OVERLAY_MIN_OCR_CONFIDENCE]
            if uncertain and not segment.protected:
                warnings.append(f"{segment.segment_id} OCR 置信度较低或未提供置信度，需重点复核")
    all_segments = [s for p in pages for s in p.segments]
    notify(45, "按分段 id 批量翻译，并校验译文对应关系")
    translate_segments(all_segments, options["source_lang"], options["target_lang"], options["translation_engine"], options["gemini_route"])

    def rebuild():
        for page in pages:
            for segment in page.segments:
                if not segment.protected:
                    fit_segment(segment, page)
            original = np.array(Image.open(page.image_path).convert("RGB"))
            clean = erase_page(original, page)
            page.clean_path = debug_dir / (page.image_path.stem.replace("original", "clean") + ".png")
            Image.fromarray(clean).save(page.clean_path)
        # 先保存临时文件再替换，渲染/保存失败仍保留上一份有效 DOCX。
        candidate = output_path.with_name("layout_overlay_candidate.docx")
        build_docx(pages, candidate, options["target_lang"])
        candidate.replace(output_path)

    notify(65, "擦除原文字、拟合字号并生成可编辑 Word")
    rebuild()
    notify(75, "执行有界质检与修补，最多两轮")
    qa_options = dict(options, vision_model=settings.LAYOUT_OVERLAY_QA_MODEL or options["vision_model"])
    report = run_qa(pages, output_path, debug_dir, qa_options, rebuild, notify)
    report["warnings"].extend(warnings + router.warnings)
    report_path = output_dir / "qa_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    layout_path = debug_dir / "layout.json"
    debug = [{"page": i + 1, "width": page.width, "height": page.height,
              "width_pt": page.width_pt, "height_pt": page.height_pt, "ocr_provider": page.ocr_provider,
              "protected_regions": page.protected_regions, "segments": [asdict(s) for s in page.segments]}
             for i, page in enumerate(pages)]
    layout_path.write_text(json.dumps(debug, ensure_ascii=False, indent=2), encoding="utf-8")
    archive = output_dir / "debug_artifacts.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        bundle.write(report_path, "qa_report.json")
        for file in debug_dir.rglob("*"):
            if file.is_file():
                bundle.write(file, file.relative_to(debug_dir))
    notify(95, "原版式 Word 与质检报告已生成")
    return {"output_docx": str(output_path.resolve()), "qa_report_path": str(report_path.resolve()),
            "debug_archive": str(archive.resolve()), "qa_report": report, "page_count": len(pages),
            "segment_count": len(all_segments), "target_lang": options["target_lang"],
            "original_filename": original_filename, "warnings": report["warnings"]}


def _run_editable(pages, output_dir, debug_dir, options, router, original_filename, notify):
    documents = []
    for index, page in enumerate(pages):
        notify(25 + int(index / len(pages) * 15), f"第 {index + 1} 页：完整读取正文、字段与照片")
        documents.append(extract_document(page, index, options))
    notify(45, "翻译完整语义段，保留姓名、日期、编号与机构信息")
    translate_documents(documents, options)
    notify(65, "生成清晰可编辑 Word，正文自动换行")
    output_path = output_dir / "editable_translation.docx"
    build_editable_docx(pages, documents, output_path, options["target_lang"])
    notify(75, "核对完整性、事实与 Word 实际渲染")
    report = review_editable(pages, documents, output_path, debug_dir, options)
    report["warnings"].extend(router.warnings)
    report_path = output_dir / "qa_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (debug_dir / "layout.json").write_text(json.dumps(documents, ensure_ascii=False, indent=2), encoding="utf-8")
    archive = output_dir / "debug_artifacts.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        bundle.write(report_path, "qa_report.json")
        for file in debug_dir.rglob("*"):
            if file.is_file():
                bundle.write(file, file.relative_to(debug_dir))
    notify(95, "可编辑译文 Word 与质检报告已生成")
    return {"output_docx": str(output_path.resolve()), "qa_report_path": str(report_path.resolve()),
            "debug_archive": str(archive.resolve()), "qa_report": report, "page_count": len(pages),
            "segment_count": sum(len(d["blocks"]) for d in documents), "target_lang": options["target_lang"],
            "output_mode": "editable", "original_filename": original_filename, "warnings": report["warnings"]}


async def execute_layout_overlay_task(*, task_id, display_no, input_path, original_filename="input.pdf",
                                      progress_callback=None, executor=None, **options):
    normalized = normalize_options(**options)
    if Path(input_path).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError("仅支持 PDF 和常见图片文件")
    loop = asyncio.get_running_loop()

    def notify(progress, message):
        if progress_callback:
            asyncio.run_coroutine_threadsafe(progress_callback(progress, message), loop).result()

    work = functools.partial(_run_pipeline, task_id=task_id, display_no=display_no, input_path=input_path,
                             original_filename=original_filename, options=normalized, notify=notify)
    return await loop.run_in_executor(executor, work)
