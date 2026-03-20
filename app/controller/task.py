import os
import traceback
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from app.service.doc_translate_service import get_doc_translate_models, get_supported_languages
from app.service import zhongfanyi_service as zf_service
from app.service.number_check_service import _get_task_progress as get_number_check_progress
from app.service.pdf2docx_service import get_pdf2docx_models
from app.service.task_queue_service import task_queue_service

router = APIRouter(prefix="/task", tags=["Task"])


@router.post("/run")
async def run_task(
    file: UploadFile = File(..., description="image file"),
    from_lang: str = Query("zh"),
    to_lang: str = Query("en"),
    enable_correction: bool = Query(False),
    enable_visualization: bool = Query(True),
    card_side: str = Query("front"),
    doc_type: str = Query("id_card"),
    marriage_page_template: str = Query("page2"),
    registrar_signature_text: Optional[str] = Query(None),
    registered_by_text: Optional[str] = Query(None),
    registered_by_offset_x: int = Query(20),
    registered_by_offset_y: int = Query(-80),
    registrar_signature_offset_x: int = Query(48),
    registrar_signature_offset_y: int = Query(-12),
    enable_merge: bool = Query(True),
    enable_overlap_fix: bool = Query(True),
    enable_colon_fix: bool = Query(True),
    font_size: Optional[int] = Query(None),
):
    try:
        task_id = await task_queue_service.submit_ocr_task(
            file=file,
            from_lang=from_lang,
            to_lang=to_lang,
            enable_correction=enable_correction,
            enable_visualization=enable_visualization,
            card_side=card_side,
            doc_type=doc_type,
            marriage_page_template=marriage_page_template,
            registrar_signature_text=registrar_signature_text,
            registered_by_text=registered_by_text,
            registered_by_offset_x=registered_by_offset_x,
            registered_by_offset_y=registered_by_offset_y,
            registrar_signature_offset_x=registrar_signature_offset_x,
            registrar_signature_offset_y=registrar_signature_offset_y,
            enable_merge=enable_merge,
            enable_overlap_fix=enable_overlap_fix,
            enable_colon_fix=enable_colon_fix,
            font_size=font_size,
        )
        return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}
    except Exception as exc:
        tb = traceback.format_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(exc),
                "type": type(exc).__name__,
                "traceback": tb.split("\n")[-10:] if tb else [],
            },
        )


@router.get("/run/status/{task_id}")
async def get_run_task_status(task_id: str):
    task = task_queue_service.get_task_status(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return task


@router.post("/number-check")
async def run_number_check(
    original_file: UploadFile = File(..., description="original docx"),
    translated_file: UploadFile = File(..., description="translated docx"),
):
    task_id = await task_queue_service.submit_number_check_task(
        original_file=original_file,
        translated_file=translated_file,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}


@router.get("/number-check/status/{task_id}")
async def get_number_check_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="任务不存在")
    progress = get_number_check_progress(task_id)
    if progress and queue_task.get("status") != "queued":
        return progress
    return queue_task


@router.post("/zhongfanyi")
async def run_zhongfanyi(
    original_file: UploadFile = File(..., description="original document"),
    translated_file: UploadFile = File(..., description="translated document"),
    use_ai_rule: bool = Query(False),
    rule_file: Optional[UploadFile] = File(None),
    session_rule_content: Optional[str] = Form(None),
):
    allowed = {".docx", ".doc", ".pdf"}
    ext_orig = os.path.splitext(original_file.filename or "")[1].lower()
    ext_tran = os.path.splitext(translated_file.filename or "")[1].lower()
    if ext_orig not in allowed:
        raise HTTPException(status_code=400, detail=f"不支持的原文格式: {ext_orig}")
    if ext_tran not in allowed:
        raise HTTPException(status_code=400, detail=f"不支持的译文格式: {ext_tran}")

    task_id = await task_queue_service.submit_zhongfanyi_task(
        original_file=original_file,
        translated_file=translated_file,
        use_ai_rule=use_ai_rule,
        rule_file=rule_file,
        session_rule_content=session_rule_content,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}


@router.get("/zhongfanyi/status/{task_id}")
async def get_zhongfanyi_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="任务不存在")
    progress = zf_service.get_task_progress(task_id)
    if progress and queue_task.get("status") != "queued":
        return progress
    return queue_task


class RuleUpdateBody(BaseModel):
    rule_type: str
    content: str


def get_rule_file_path(rule_type: str) -> Path:
    base_path = Path(__file__).resolve().parents[2] / "专检" / "zhongfanyi" / "llm" / "llm_project" / "rule"
    if rule_type == "custom":
        return base_path / "自定义规则.txt"
    if rule_type == "default":
        return base_path / "默认规则.txt"
    raise HTTPException(status_code=400, detail="未知的规则类型")


@router.get("/zhongfanyi/rule")
async def get_zhongfanyi_rule(rule_type: str = Query(..., description="custom or default")):
    file_path = get_rule_file_path(rule_type)
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="规则文件不存在")
    try:
        return {"status": "ok", "content": file_path.read_text(encoding="utf-8")}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"读取规则文件失败: {exc}")


@router.post("/zhongfanyi/rule")
async def update_zhongfanyi_rule(body: RuleUpdateBody):
    raise HTTPException(status_code=400, detail="规则编辑仅在当前会话内生效")


@router.get("/alignment/config")
async def get_alignment_config():
    from app.service.alignment_service import (
        AVAILABLE_MODELS as ALIGNMENT_MODELS,
        BUFFER_CHARS,
        SUPPORTED_LANGUAGES,
        THRESHOLD_MAP,
    )

    return {
        "models": {
            name: {
                "description": info["description"],
                "id": info["id"],
                "max_output": info["max_output"],
            }
            for name, info in ALIGNMENT_MODELS.items()
        },
        "languages": {k: v["description"] for k, v in SUPPORTED_LANGUAGES.items()},
        "thresholds": THRESHOLD_MAP,
        "buffer_chars": BUFFER_CHARS,
    }


@router.post("/alignment")
async def run_alignment(
    original_file: UploadFile = File(..., description="original file"),
    translated_file: UploadFile = File(..., description="translated file"),
    source_lang: str = Query("中文"),
    target_lang: str = Query("英语"),
    model_name: str = Query("Google gemini-3-flash-preview"),
    enable_post_split: bool = Query(True),
    threshold_2: int = Query(25000),
    threshold_3: int = Query(50000),
    threshold_4: int = Query(75000),
    threshold_5: int = Query(100000),
    threshold_6: int = Query(125000),
    threshold_7: int = Query(150000),
    threshold_8: int = Query(175000),
    buffer_chars: int = Query(2000),
):
    allowed_ext = {".docx", ".doc", ".pptx", ".xlsx", ".xls"}
    orig_ext = os.path.splitext(original_file.filename or "")[1].lower()
    trans_ext = os.path.splitext(translated_file.filename or "")[1].lower()
    if orig_ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的原文文件格式: {orig_ext}")
    if trans_ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的译文文件格式: {trans_ext}")

    task_id = await task_queue_service.submit_alignment_task(
        original_file=original_file,
        translated_file=translated_file,
        source_lang=source_lang,
        target_lang=target_lang,
        model_name=model_name,
        enable_post_split=enable_post_split,
        threshold_2=threshold_2,
        threshold_3=threshold_3,
        threshold_4=threshold_4,
        threshold_5=threshold_5,
        threshold_6=threshold_6,
        threshold_7=threshold_7,
        threshold_8=threshold_8,
        buffer_chars=buffer_chars,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}


@router.get("/alignment/status/{task_id}")
async def get_alignment_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="任务不存在")
    from app.service.alignment_service import get_alignment_progress

    progress = get_alignment_progress(task_id)
    if queue_task.get("status") in ("done", "failed"):
        if progress and progress.get("stream_log"):
            queue_task["stream_log"] = progress.get("stream_log")
            if queue_task.get("result"):
                queue_task["result"]["stream_log"] = progress.get("stream_log")
        return queue_task
    if progress and queue_task.get("status") != "queued":
        return progress
    return queue_task


# ── 文档翻译（原营业执照板块，已重构） ──────────────────────────


@router.get("/doc-translate/config")
async def get_doc_translate_config():
    return {
        "models": get_doc_translate_models(),
        "default_model": "google/gemini-3-flash-preview",
        "languages": get_supported_languages(),
    }


@router.post("/doc-translate")
async def submit_doc_translate(
    file: UploadFile = File(..., description="pdf or image file"),
    source_lang: str = Query("zh"),
    target_langs: str = Query("en", description="逗号分隔的目标语言代码，如 en,es,ja"),
    ocr_model: str = Query("google/gemini-3-flash-preview"),
):
    allowed_ext = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的文件格式: {ext}")

    task_id = await task_queue_service.submit_doc_translate_task(
        file=file,
        source_lang=source_lang,
        target_langs=target_langs,
        ocr_model=ocr_model,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}


@router.get("/doc-translate/status/{task_id}")
async def get_doc_translate_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="任务不存在")
    return queue_task


@router.get("/pdf2docx/config")
async def get_pdf2docx_config():
    return {
        "models": get_pdf2docx_models(),
        "default_model": "google/gemini-3-flash-preview",
    }


@router.post("/pdf2docx")
async def run_pdf2docx(
    file: UploadFile = File(..., description="pdf or image file"),
    model: str = Query("google/gemini-3-flash-preview"),
):
    allowed_ext = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的文件格式: {ext}")

    task_id = await task_queue_service.submit_pdf2docx_task(file=file, model=model)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "任务已提交"}


@router.get("/pdf2docx/status/{task_id}")
async def get_pdf2docx_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="任务不存在")
    return queue_task
