import json
import os
import traceback
from pathlib import Path
from typing import List, Optional

from fastapi import APIRouter, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel

from app.db.session import SessionLocal
from app.repository import task_repo
from app.service import zhongfanyi_service as zf_service
from app.service.doc_translate_service import get_doc_translate_models, get_supported_languages
from app.service.drivers_license_service import get_drivers_license_config
from app.service.gemini_service import get_gemini_routes
from app.service.number_check_service import _get_task_progress as get_number_check_progress, get_number_check_models
from app.service.pdf2docx_service import get_pdf2docx_models
from app.service.task_queue_service import task_queue_service

router = APIRouter(prefix="/task", tags=["Task"])
BASE_DIR = Path(__file__).resolve().parents[2]
ZHONGFANYI_RULE_DIR = BASE_DIR / "\u4e13\u68c0" / "zhongfanyi" / "llm" / "llm_project" / "rule"


class RuleUpdateBody(BaseModel):
    rule_type: str
    content: str


def _safe_resolve(file_path: str) -> Path:
    path = Path(file_path)
    if not path.is_absolute():
        path = BASE_DIR / path
    resolved = path.resolve()
    if not str(resolved).startswith(str(BASE_DIR.resolve())):
        raise HTTPException(status_code=403, detail="Path is not allowed")
    return resolved


def _task_to_dict(task) -> dict:
    return {
        "task_id": task.task_id,
        "display_no": task.display_no,
        "task_type": task.task_type,
        "task_label": task.task_label or task.task_type,
        "filename": task.filename,
        "status": task.status,
        "progress": task.progress,
        "message": task.message or "",
        "error": task.error_message,
        "cancel_requested": bool(task.cancel_requested),
        "created_at": task.created_at.isoformat() if task.created_at else None,
        "started_at": task.started_at.isoformat() if task.started_at else None,
        "finished_at": task.finished_at.isoformat() if task.finished_at else None,
    }



def _merge_queue_timestamps(payload: Optional[dict], queue_task: Optional[dict]) -> Optional[dict]:
    if not payload:
        return payload
    merged = dict(payload)
    if queue_task:
        for key in ("task_id", "display_no", "created_at", "started_at", "finished_at"):
            if key in queue_task and merged.get(key) is None:
                merged[key] = queue_task.get(key)
    return merged
@router.get("/list")
async def list_tasks(status: Optional[str] = Query(None), task_type: Optional[str] = Query(None), keyword: Optional[str] = Query(None), page: int = Query(1, ge=1), page_size: int = Query(20, ge=1, le=100)):
    with SessionLocal() as db:
        tasks, total = task_repo.list_tasks(db, status=status, task_type=task_type, keyword=keyword, page=page, page_size=page_size)
        return {"items": [_task_to_dict(task) for task in tasks], "total": total, "page": page, "page_size": page_size}


@router.get("/dashboard/stats")
async def dashboard_stats():
    with SessionLocal() as db:
        return task_repo.count_by_status(db)


@router.get("/{task_id}/detail")
async def task_detail(task_id: str):
    with SessionLocal() as db:
        task = task_repo.get_task_by_task_id(db, task_id)
        if not task:
            raise HTTPException(status_code=404, detail="Task not found")
        payload = _task_to_dict(task)
        payload["params"] = json.loads(task.params_json or "{}")
        payload["input_files"] = json.loads(task.input_files_json or "{}")
        payload["output_files"] = json.loads(task.output_files_json or "[]")
        payload["result"] = json.loads(task.result_json or "null")
        payload["stream_log"] = task_queue_service._task_logs.get(task_id, "")
        return payload


@router.post("/{task_id}/cancel")
async def cancel_task(task_id: str):
    with SessionLocal() as db:
        task = task_repo.get_task_by_task_id(db, task_id)
        if not task:
            raise HTTPException(status_code=404, detail="Task not found")
        if task.status in {"done", "failed", "cancelled"}:
            return {"status": task.status, "message": f"Task already in terminal state: {task.status}"}
        task_repo.cancel_task(db, task_id)
    return {"status": "ok", "message": "Cancel request submitted"}


@router.get("/{task_id}/download")
async def task_download(task_id: str, file_path: str = Query(...), download_name: Optional[str] = Query(None)):
    with SessionLocal() as db:
        task = task_repo.get_task_by_task_id(db, task_id)
        if not task:
            raise HTTPException(status_code=404, detail="Task not found")
    resolved = _safe_resolve(file_path)
    if not resolved.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(str(resolved), filename=download_name or resolved.name, media_type="application/octet-stream")


@router.post("/run")
async def run_task(file: UploadFile = File(...), from_lang: str = Query("zh"), to_lang: str = Query("en"), enable_correction: bool = Query(False), enable_visualization: bool = Query(True), card_side: str = Query("front"), doc_type: str = Query("id_card"), marriage_page_template: str = Query("page2"), registrar_signature_text: Optional[str] = Query(None), registered_by_text: Optional[str] = Query(None), registered_by_offset_x: int = Query(20), registered_by_offset_y: int = Query(-80), registrar_signature_offset_x: int = Query(48), registrar_signature_offset_y: int = Query(-12), enable_merge: bool = Query(True), enable_overlap_fix: bool = Query(True), enable_colon_fix: bool = Query(True), font_size: Optional[int] = Query(None)):
    try:
        task_id = await task_queue_service.submit_ocr_task(file=file, from_lang=from_lang, to_lang=to_lang, enable_correction=enable_correction, enable_visualization=enable_visualization, card_side=card_side, doc_type=doc_type, marriage_page_template=marriage_page_template, registrar_signature_text=registrar_signature_text, registered_by_text=registered_by_text, registered_by_offset_x=registered_by_offset_x, registered_by_offset_y=registered_by_offset_y, registrar_signature_offset_x=registrar_signature_offset_x, registrar_signature_offset_y=registrar_signature_offset_y, enable_merge=enable_merge, enable_overlap_fix=enable_overlap_fix, enable_colon_fix=enable_colon_fix, font_size=font_size)
        return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}
    except Exception as exc:
        tb = traceback.format_exc()
        raise HTTPException(status_code=500, detail={"error": str(exc), "type": type(exc).__name__, "traceback": tb.split("\n")[-10:] if tb else []})


@router.post("/run/batch")
async def run_task_batch(files: List[UploadFile] = File(...), from_lang: str = Query("zh"), to_lang: str = Query("en"), enable_correction: bool = Query(False), enable_visualization: bool = Query(True), card_side: str = Query("front"), doc_type: str = Query("id_card"), marriage_page_template: str = Query("page2"), registrar_signature_text: Optional[str] = Query(None), registered_by_text: Optional[str] = Query(None), registered_by_offset_x: int = Query(20), registered_by_offset_y: int = Query(-80), registrar_signature_offset_x: int = Query(48), registrar_signature_offset_y: int = Query(-12), enable_merge: bool = Query(True), enable_overlap_fix: bool = Query(True), enable_colon_fix: bool = Query(True), font_size: Optional[int] = Query(None)):
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    results = []
    for file in files:
        try:
            task_id = await task_queue_service.submit_ocr_task(file=file, from_lang=from_lang, to_lang=to_lang, enable_correction=enable_correction, enable_visualization=enable_visualization, card_side=card_side, doc_type=doc_type, marriage_page_template=marriage_page_template, registrar_signature_text=registrar_signature_text, registered_by_text=registered_by_text, registered_by_offset_x=registered_by_offset_x, registered_by_offset_y=registered_by_offset_y, registrar_signature_offset_x=registrar_signature_offset_x, registrar_signature_offset_y=registrar_signature_offset_y, enable_merge=enable_merge, enable_overlap_fix=enable_overlap_fix, enable_colon_fix=enable_colon_fix, font_size=font_size)
            results.append({"filename": file.filename, "task_id": task_id, "status": "ACCEPTED"})
        except Exception as exc:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": str(exc)})
    return {"status": "ACCEPTED", "tasks": results, "total": len(results)}


@router.get("/run/status/{task_id}")
async def get_run_task_status(task_id: str):
    task = task_queue_service.get_task_status(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail="Task not found")
    return task


@router.post("/number-check")
async def run_number_check(original_file: UploadFile = File(...), translated_file: UploadFile = File(...), gemini_route: str = Query("openrouter"), model_name: str = Query("gemini-3.1-pro-preview")):
    task_id = await task_queue_service.submit_number_check_task(original_file=original_file, translated_file=translated_file, gemini_route=gemini_route, model_name=model_name)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/number-check/config")
async def get_number_check_config():
    return {"models": get_number_check_models(), "default_model": "gemini-3.1-pro-preview", "routes": get_gemini_routes(), "default_route": "openrouter"}


@router.get("/number-check/status/{task_id}")
async def get_number_check_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    progress = get_number_check_progress(task_id)
    if queue_task.get("status") in {"done", "failed"}:
        if progress and progress.get("stream_log"):
            queue_task["stream_log"] = progress.get("stream_log")
        return queue_task
    if progress and queue_task.get("status") != "queued":
        return _merge_queue_timestamps(progress, queue_task)
    return queue_task


@router.post("/zhongfanyi")
async def run_zhongfanyi(original_file: UploadFile = File(...), translated_file: UploadFile = File(...), use_ai_rule: bool = Query(False), gemini_route: str = Query("openrouter"), rule_file: Optional[UploadFile] = File(None), session_rule_content: Optional[str] = Form(None)):
    allowed = {".docx", ".doc", ".pdf"}
    if os.path.splitext(original_file.filename or "")[1].lower() not in allowed:
        raise HTTPException(status_code=400, detail="Unsupported original file format")
    if os.path.splitext(translated_file.filename or "")[1].lower() not in allowed:
        raise HTTPException(status_code=400, detail="Unsupported translated file format")
    task_id = await task_queue_service.submit_zhongfanyi_task(original_file=original_file, translated_file=translated_file, use_ai_rule=use_ai_rule, gemini_route=gemini_route, rule_file=rule_file, session_rule_content=session_rule_content)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/zhongfanyi/status/{task_id}")
async def get_zhongfanyi_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    progress = zf_service.get_task_progress(task_id)
    if progress and queue_task.get("status") != "queued":
        return _merge_queue_timestamps(progress, queue_task)
    return queue_task


def get_rule_file_path(rule_type: str) -> Path:
    if rule_type == "custom":
        return ZHONGFANYI_RULE_DIR / "\u81ea\u5b9a\u4e49\u89c4\u5219.txt"
    if rule_type == "default":
        return ZHONGFANYI_RULE_DIR / "\u9ed8\u8ba4\u89c4\u5219.txt"
    raise HTTPException(status_code=400, detail="Unknown rule type")


@router.get("/zhongfanyi/rule")
async def get_zhongfanyi_rule(rule_type: str = Query(...)):
    file_path = get_rule_file_path(rule_type)
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Rule file not found")
    return {"status": "ok", "content": file_path.read_text(encoding="utf-8")}


@router.post("/zhongfanyi/rule")
async def update_zhongfanyi_rule(body: RuleUpdateBody):
    raise HTTPException(status_code=400, detail="Rule editing is session-only")


@router.get("/alignment/config")
async def get_alignment_config():
    from app.service.alignment_service import AVAILABLE_MODELS as ALIGNMENT_MODELS, BUFFER_CHARS, SUPPORTED_LANGUAGES, THRESHOLD_MAP
    return {"models": {name: {"description": info["description"], "id": info["id"], "max_output": info["max_output"]} for name, info in ALIGNMENT_MODELS.items()}, "routes": get_gemini_routes(), "default_route": "openrouter", "languages": {k: v["description"] for k, v in SUPPORTED_LANGUAGES.items()}, "thresholds": THRESHOLD_MAP, "buffer_chars": BUFFER_CHARS}


@router.post("/alignment")
async def run_alignment(original_file: UploadFile = File(...), translated_file: UploadFile = File(...), source_lang: str = Query("zh"), target_lang: str = Query("en"), model_name: str = Query("Google gemini-3-flash-preview"), gemini_route: str = Query("openrouter"), enable_post_split: bool = Query(True), threshold_2: int = Query(25000), threshold_3: int = Query(50000), threshold_4: int = Query(75000), threshold_5: int = Query(100000), threshold_6: int = Query(125000), threshold_7: int = Query(150000), threshold_8: int = Query(175000), buffer_chars: int = Query(2000)):
    allowed_ext = {".docx", ".doc", ".pptx", ".xlsx", ".xls"}
    if os.path.splitext(original_file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported original file format")
    if os.path.splitext(translated_file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported translated file format")
    task_id = await task_queue_service.submit_alignment_task(original_file=original_file, translated_file=translated_file, source_lang=source_lang, target_lang=target_lang, model_name=model_name, gemini_route=gemini_route, enable_post_split=enable_post_split, threshold_2=threshold_2, threshold_3=threshold_3, threshold_4=threshold_4, threshold_5=threshold_5, threshold_6=threshold_6, threshold_7=threshold_7, threshold_8=threshold_8, buffer_chars=buffer_chars)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/alignment/status/{task_id}")
async def get_alignment_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    from app.service.alignment_service import get_alignment_progress
    progress = get_alignment_progress(task_id)
    if queue_task.get("status") in {"done", "failed"}:
        if progress and progress.get("stream_log"):
            queue_task["stream_log"] = progress.get("stream_log")
            if queue_task.get("result"):
                queue_task["result"]["stream_log"] = progress.get("stream_log")
        return queue_task
    if progress and queue_task.get("status") != "queued":
        return _merge_queue_timestamps(progress, queue_task)
    return queue_task


@router.get("/drivers-license/config")
async def get_drivers_license_page_config():
    return get_drivers_license_config()


@router.post("/drivers-license")
async def submit_drivers_license(files: List[UploadFile] = File(...), processing_mode: str = Query("merge")):
    allowed_ext = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
    if not files:
        raise HTTPException(status_code=400, detail="At least one image is required")
    if len(files) > 20:
        raise HTTPException(status_code=400, detail="Too many images in one request")
    for file in files:
        if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
            raise HTTPException(status_code=400, detail="Unsupported image format")
    task_id = await task_queue_service.submit_drivers_license_task(files=files, processing_mode=processing_mode)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/drivers-license/status/{task_id}")
async def get_drivers_license_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/doc-translate/config")
async def get_doc_translate_config():
    return {"models": get_doc_translate_models(), "default_model": "google/gemini-3-flash-preview", "routes": get_gemini_routes(), "default_route": "openrouter", "languages": get_supported_languages()}


@router.post("/doc-translate")
async def submit_doc_translate(file: UploadFile = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), ocr_model: str = Query("google/gemini-3-flash-preview"), gemini_route: str = Query("openrouter")):
    allowed_ext = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
    if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported file format")
    task_id = await task_queue_service.submit_doc_translate_task(file=file, source_lang=source_lang, target_langs=target_langs, ocr_model=ocr_model, gemini_route=gemini_route)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.post("/doc-translate/batch")
async def submit_doc_translate_batch(files: List[UploadFile] = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), ocr_model: str = Query("google/gemini-3-flash-preview"), gemini_route: str = Query("openrouter")):
    allowed_ext = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    results = []
    for file in files:
        if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": "Unsupported file format"})
            continue
        try:
            task_id = await task_queue_service.submit_doc_translate_task(file=file, source_lang=source_lang, target_langs=target_langs, ocr_model=ocr_model, gemini_route=gemini_route)
            results.append({"filename": file.filename, "task_id": task_id, "status": "ACCEPTED"})
        except Exception as exc:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": str(exc)})
    return {"status": "ACCEPTED", "tasks": results, "total": len(results)}


@router.get("/doc-translate/status/{task_id}")
async def get_doc_translate_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/pdf2docx/config")
async def get_pdf2docx_config():
    return {"models": get_pdf2docx_models(), "default_model": "google/gemini-3-flash-preview", "routes": get_gemini_routes(), "default_route": "google"}


@router.post("/pdf2docx")
async def run_pdf2docx(file: UploadFile = File(...), model: str = Query("google/gemini-3-flash-preview"), gemini_route: str = Query("google")):
    allowed_ext = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported file format")
    if ext == ".pdf":
        try:
            import fitz
            content = await file.read()
            pdf_doc = fitz.open(stream=content, filetype="pdf")
            if len(pdf_doc) > 100:
                raise HTTPException(status_code=400, detail="PDF has too many pages")
            pdf_doc.close()
            await file.seek(0)
        except HTTPException:
            raise
        except Exception:
            await file.seek(0)
    task_id = await task_queue_service.submit_pdf2docx_task(file=file, model=model, gemini_route=gemini_route)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.post("/pdf2docx/batch")
async def run_pdf2docx_batch(files: List[UploadFile] = File(...), model: str = Query("google/gemini-3-flash-preview"), gemini_route: str = Query("google")):
    allowed_ext = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    results = []
    for file in files:
        ext = os.path.splitext(file.filename or "")[1].lower()
        if ext not in allowed_ext:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": "Unsupported file format"})
            continue
        if ext == ".pdf":
            try:
                import fitz
                content = await file.read()
                pdf_doc = fitz.open(stream=content, filetype="pdf")
                page_count = len(pdf_doc)
                pdf_doc.close()
                if page_count > 100:
                    results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": f"PDF has too many pages ({page_count})"})
                    continue
                await file.seek(0)
            except Exception:
                await file.seek(0)
        try:
            task_id = await task_queue_service.submit_pdf2docx_task(file=file, model=model, gemini_route=gemini_route)
            results.append({"filename": file.filename, "task_id": task_id, "status": "ACCEPTED"})
        except Exception as exc:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": str(exc)})
    return {"status": "ACCEPTED", "tasks": results, "total": len(results)}


@router.get("/pdf2docx/status/{task_id}")
async def get_pdf2docx_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task

