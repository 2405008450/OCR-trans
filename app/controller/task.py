import asyncio
import json
import os
import tempfile
import traceback
import zipfile
from pathlib import Path
from typing import List, Optional

from fastapi import APIRouter, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.background import BackgroundTask

from app.db.session import SessionLocal
from app.core.task_model_display import build_task_model_info
from app.repository import task_repo
from app.service import zhongfanyi_service as zf_service
from app.service.business_licence_service import (
    BUSINESS_LICENCE_DEFAULT_MODEL,
    BUSINESS_LICENCE_DEFAULT_ROUTE,
    extract_business_licence_data,
    get_business_licence_company_name,
    get_business_licence_models,
)
from app.service.doc_translate_service import (
    DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE,
    DOC_TRANSLATE_DEFAULT_MODEL,
    get_doc_translate_allowed_extensions,
    get_doc_translate_models,
    get_supported_languages,
)
from app.service.drivers_license_service import get_drivers_license_config
from app.service.gemini_service import get_gemini_routes
from app.service.number_check_service import (
    _get_task_progress as get_number_check_progress,
    get_number_check_default_mode,
    get_number_check_models,
)
from app.service.pdf2docx_service import (
    PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    PDF2DOCX_DEFAULT_MODEL,
    get_pdf2docx_models,
)
from app.service.task_queue_service import task_queue_service

router = APIRouter(prefix="/task", tags=["Task"])
BASE_DIR = Path(__file__).resolve().parents[2]
ZHONGFANYI_RULE_DIR = BASE_DIR / "\u4e13\u68c0" / "\u4e2d\u7ffb\u8bd1" / "rule"


class RuleUpdateBody(BaseModel):
    rule_type: str
    content: str


class BatchDownloadBody(BaseModel):
    task_ids: List[str]
    extensions: Optional[List[str]] = None
    archive_name: Optional[str] = None


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


def _normalize_extensions(extensions: Optional[List[str]]) -> Optional[set[str]]:
    if not extensions:
        return None
    normalized = set()
    for item in extensions:
        if not item:
            continue
        ext = str(item).strip().lower()
        if not ext:
            continue
        if not ext.startswith("."):
            ext = f".{ext}"
        normalized.add(ext)
    return normalized or None


def _iter_task_output_files(task, allowed_extensions: Optional[set[str]] = None) -> list[dict]:
    files = json.loads(task.output_files_json or "[]")
    if not files and task.output_path:
        files = [{"name": Path(task.output_path).name, "path": task.output_path, "type": "output"}]

    matched: list[dict] = []
    for item in files:
        if not isinstance(item, dict):
            continue
        file_path = item.get("path")
        if not file_path:
            continue
        resolved = _safe_resolve(file_path)
        if not resolved.exists() or not resolved.is_file():
            continue
        if allowed_extensions and resolved.suffix.lower() not in allowed_extensions:
            continue
        matched.append(
            {
                "resolved": resolved,
                "display_name": Path(item.get("name") or resolved.name).name,
            }
        )
    return matched


def _build_archive_name(task, display_name: str, used_names: set[str]) -> str:
    stem = Path(display_name).stem or "output"
    suffix = Path(display_name).suffix
    prefix = task.display_no or task.task_id or "task"
    candidate = f"{prefix}_{stem}{suffix}"
    index = 2
    while candidate in used_names:
        candidate = f"{prefix}_{stem}_{index}{suffix}"
        index += 1
    used_names.add(candidate)
    return candidate



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
        payload["model_info"] = build_task_model_info(task.task_type, payload["params"], payload["result"])
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


@router.post("/batch-download")
async def batch_download(body: BatchDownloadBody):
    if not body.task_ids:
        raise HTTPException(status_code=400, detail="At least one task_id is required")

    allowed_extensions = _normalize_extensions(body.extensions)
    tasks = []
    with SessionLocal() as db:
        for task_id in body.task_ids:
            task = task_repo.get_task_by_task_id(db, task_id)
            if task:
                tasks.append(task)

    archive_sources: list[tuple[Path, str]] = []
    used_names: set[str] = set()
    for task in tasks:
        if task.status != "done":
            continue
        for item in _iter_task_output_files(task, allowed_extensions):
            archive_name = _build_archive_name(task, item["display_name"], used_names)
            archive_sources.append((item["resolved"], archive_name))

    if not archive_sources:
        raise HTTPException(status_code=400, detail="No matching output files found")

    temp_dir = BASE_DIR / "outputs" / "_tmp_batch_downloads"
    temp_dir.mkdir(parents=True, exist_ok=True)
    download_name = (body.archive_name or "batch_outputs.zip").strip() or "batch_outputs.zip"
    if not download_name.lower().endswith(".zip"):
        download_name = f"{download_name}.zip"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip", dir=temp_dir) as temp_file:
        archive_path = Path(temp_file.name)

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for source_path, archive_name in archive_sources:
            zip_file.write(source_path, arcname=archive_name)

    return FileResponse(
        str(archive_path),
        filename=download_name,
        media_type="application/zip",
        background=BackgroundTask(lambda path=archive_path: path.unlink(missing_ok=True)),
    )


@router.post("/run")
async def run_task(file: UploadFile = File(...), from_lang: str = Query("zh"), to_lang: str = Query("en"), enable_correction: bool = Query(False), enable_visualization: bool = Query(True), card_side: str = Query("front")):
    try:
        task_id = await task_queue_service.submit_ocr_task(file=file, from_lang=from_lang, to_lang=to_lang, enable_correction=enable_correction, enable_visualization=enable_visualization, card_side=card_side)
        return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}
    except Exception as exc:
        tb = traceback.format_exc()
        raise HTTPException(status_code=500, detail={"error": str(exc), "type": type(exc).__name__, "traceback": tb.split("\n")[-10:] if tb else []})


@router.post("/run/batch")
async def run_task_batch(files: List[UploadFile] = File(...), from_lang: str = Query("zh"), to_lang: str = Query("en"), enable_correction: bool = Query(False), enable_visualization: bool = Query(True), card_side: str = Query("front")):
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    results = []
    for file in files:
        try:
            task_id = await task_queue_service.submit_ocr_task(file=file, from_lang=from_lang, to_lang=to_lang, enable_correction=enable_correction, enable_visualization=enable_visualization, card_side=card_side)
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
async def run_number_check(
    original_file: Optional[UploadFile] = File(None),
    translated_file: Optional[UploadFile] = File(None),
    single_file: Optional[UploadFile] = File(None),
    mode: Optional[str] = Query(None),
    gemini_route: str = Query("openrouter"),
    model_name: str = Query("gemini-3.1-pro-preview"),
):
    resolved_mode = (mode or get_number_check_default_mode()).strip().lower()
    if resolved_mode == "single":
        if single_file is None:
            raise HTTPException(status_code=400, detail="单文件模式需要上传 single_file")
    elif resolved_mode == "double":
        if original_file is None or translated_file is None:
            raise HTTPException(status_code=400, detail="双文件模式需要同时上传 original_file 和 translated_file")
    else:
        raise HTTPException(status_code=400, detail=f"不支持的数字专检模式: {mode}")

    task_id = await task_queue_service.submit_number_check_task(
        mode=resolved_mode,
        original_file=original_file,
        translated_file=translated_file,
        single_file=single_file,
        gemini_route=gemini_route,
        model_name=model_name,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/number-check/config")
async def get_number_check_config():
    double_file_extensions = [".docx", ".doc"]
    single_file_extensions = [".docx", ".doc", ".pdf", ".xlsx", ".pptx"]
    return {
        "models": get_number_check_models(),
        "default_model": "gemini-3.1-pro-preview",
        "routes": get_gemini_routes(),
        "default_route": "openrouter",
        "default_mode": get_number_check_default_mode(),
        "modes": {
            "double": {
                "label": "双文件模式",
                "description": f"上传原文和译文两个 {' / '.join(double_file_extensions).upper().replace('.', '')} 文件，输出修订版译文。",
            },
            "single": {
                "label": "单文件模式",
                "description": "上传一个中英对照文件；DOC / DOCX 可生成修订版，其它格式仅输出报告。",
            },
        },
        "single_file_extensions": single_file_extensions,
        "double_file_extensions": double_file_extensions,
    }


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
async def run_zhongfanyi(
    original_file: Optional[UploadFile] = File(None),
    translated_file: Optional[UploadFile] = File(None),
    single_file: Optional[UploadFile] = File(None),
    mode: Optional[str] = Query(None),
    use_ai_rule: bool = Query(False),
    gemini_route: str = Query("openrouter"),
    model_name: str = Query(zf_service.ZHONGFANYI_DEFAULT_MODEL),
    rule_file: Optional[UploadFile] = File(None),
    session_rule_content: Optional[str] = Form(None),
):
    resolved_mode = (mode or zf_service.get_zhongfanyi_default_mode()).strip().lower()
    allowed_double = set(zf_service.get_zhongfanyi_double_file_extensions())
    allowed_single = set(zf_service.get_zhongfanyi_single_file_extensions())

    if resolved_mode == zf_service.ZHONGFANYI_MODE_SINGLE:
        if single_file is None:
            raise HTTPException(status_code=400, detail="单文件模式需要上传 single_file")
        if os.path.splitext(single_file.filename or "")[1].lower() not in allowed_single:
            raise HTTPException(status_code=400, detail="Unsupported single file format")
    elif resolved_mode == zf_service.ZHONGFANYI_MODE_DOUBLE:
        if original_file is None or translated_file is None:
            raise HTTPException(status_code=400, detail="双文件模式需要同时上传 original_file 和 translated_file")
        if os.path.splitext(original_file.filename or "")[1].lower() not in allowed_double:
            raise HTTPException(status_code=400, detail="Unsupported original file format")
        if os.path.splitext(translated_file.filename or "")[1].lower() not in allowed_double:
            raise HTTPException(status_code=400, detail="Unsupported translated file format")
    else:
        raise HTTPException(status_code=400, detail=f"不支持的中翻译模式: {mode}")

    task_id = await task_queue_service.submit_zhongfanyi_task(
        mode=resolved_mode,
        original_file=original_file,
        translated_file=translated_file,
        single_file=single_file,
        use_ai_rule=use_ai_rule,
        gemini_route=gemini_route,
        model_name=model_name,
        rule_file=rule_file,
        session_rule_content=session_rule_content,
    )
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.get("/zhongfanyi/config")
async def get_zhongfanyi_config():
    return {
        "models": zf_service.get_zhongfanyi_models(),
        "default_model": zf_service.ZHONGFANYI_DEFAULT_MODEL,
        "routes": get_gemini_routes(),
        "default_route": "openrouter",
        "default_mode": zf_service.get_zhongfanyi_default_mode(),
        "modes": {
            zf_service.ZHONGFANYI_MODE_DOUBLE: {
                "label": "双文件模式",
                "description": "上传原文和译文两个文件，支持 Word / PDF / Excel / PPTX。",
            },
            zf_service.ZHONGFANYI_MODE_SINGLE: {
                "label": "单文件模式",
                "description": "上传一个双语对照文件，自动检查并导出 JSON / Excel 报告。",
            },
        },
        "single_file_extensions": zf_service.get_zhongfanyi_single_file_extensions(),
        "double_file_extensions": zf_service.get_zhongfanyi_double_file_extensions(),
    }


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
    if progress and progress.get("status") in {"done", "failed"}:
        merged = _merge_queue_timestamps(progress, queue_task)
        if progress.get("stream_log"):
            merged["stream_log"] = progress.get("stream_log")
            if merged.get("result"):
                merged["result"]["stream_log"] = progress.get("stream_log")
        return merged
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


@router.get("/business-licence/config")
async def get_business_licence_config():
    all_routes = get_gemini_routes()
    return {
        "models": get_business_licence_models(),
        "default_model": BUSINESS_LICENCE_DEFAULT_MODEL,
        "routes": {
            BUSINESS_LICENCE_DEFAULT_ROUTE: all_routes.get(
                BUSINESS_LICENCE_DEFAULT_ROUTE,
                {"label": BUSINESS_LICENCE_DEFAULT_ROUTE, "description": ""},
            )
        },
        "default_route": BUSINESS_LICENCE_DEFAULT_ROUTE,
    }


def _validate_business_licence_file(file: UploadFile) -> None:
    allowed_ext = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
    if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported image format")


@router.post("/business-licence/company-name-preview")
async def preview_business_licence_company_name(
    file: UploadFile = File(...),
    model: str = Query(BUSINESS_LICENCE_DEFAULT_MODEL),
):
    _validate_business_licence_file(file)

    temp_dir = BASE_DIR / "uploads" / "_tmp_business_licence_preview"
    temp_dir.mkdir(parents=True, exist_ok=True)
    suffix = Path(file.filename or "business_licence.png").suffix or ".png"

    temp_path: Optional[Path] = None
    try:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=temp_dir) as temp_file:
                temp_file.write(await file.read())
                temp_path = Path(temp_file.name)

            parsed_data = await asyncio.to_thread(
                extract_business_licence_data,
                temp_path,
                model=model,
                gemini_route=BUSINESS_LICENCE_DEFAULT_ROUTE,
            )
        except Exception as exc:
            return {
                "requires_confirmation": False,
                "original_cn_name": "",
                "ai_translated_name": "",
                "parsed_data_json": "",
                "model": model,
                "gemini_route": BUSINESS_LICENCE_DEFAULT_ROUTE,
                "preview_error": str(exc),
            }
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink(missing_ok=True)

    company_name_info = get_business_licence_company_name(parsed_data) or {}
    original_cn_name = company_name_info.get("original_cn_name", "")
    ai_translated_name = company_name_info.get("ai_translated_name", "")

    return {
        "requires_confirmation": bool(original_cn_name and ai_translated_name),
        "original_cn_name": original_cn_name,
        "ai_translated_name": ai_translated_name,
        "parsed_data_json": json.dumps(parsed_data, ensure_ascii=False),
        "model": model,
        "gemini_route": BUSINESS_LICENCE_DEFAULT_ROUTE,
    }


@router.post("/business-licence")
async def submit_business_licence(
    file: UploadFile = File(...),
    parsed_data_json: Optional[str] = Form(None),
    company_name_override: Optional[str] = Form(None),
    model: str = Query(BUSINESS_LICENCE_DEFAULT_MODEL),
):
    _validate_business_licence_file(file)

    parsed_data = None
    if parsed_data_json:
        try:
            parsed_data = json.loads(parsed_data_json)
        except json.JSONDecodeError as exc:
            raise HTTPException(status_code=400, detail=f"Invalid parsed_data_json: {exc}") from exc

    task_id = await task_queue_service.submit_business_licence_task(
        file=file,
        model=model,
        gemini_route=BUSINESS_LICENCE_DEFAULT_ROUTE,
        parsed_data=parsed_data,
        company_name_override=company_name_override,
    )
    return {
        "status": "ACCEPTED",
        "task_id": task_id,
        "message": "Task submitted",
        "gemini_route": BUSINESS_LICENCE_DEFAULT_ROUTE,
    }


@router.get("/business-licence/status/{task_id}")
async def get_business_licence_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/doc-translate/config")
async def get_doc_translate_config():
    return {"models": get_doc_translate_models(), "default_model": DOC_TRANSLATE_DEFAULT_MODEL, "routes": get_gemini_routes(), "default_route": DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE, "languages": get_supported_languages(), "allowed_extensions": get_doc_translate_allowed_extensions()}


@router.post("/doc-translate")
async def submit_doc_translate(file: UploadFile = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), ocr_model: str = Query(DOC_TRANSLATE_DEFAULT_MODEL), gemini_route: str = Query(DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE)):
    allowed_ext = set(get_doc_translate_allowed_extensions())
    if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported file format")
    task_id = await task_queue_service.submit_doc_translate_task(file=file, source_lang=source_lang, target_langs=target_langs, ocr_model=ocr_model, gemini_route=gemini_route)
    return {"status": "ACCEPTED", "task_id": task_id, "message": "Task submitted"}


@router.post("/doc-translate/batch")
async def submit_doc_translate_batch(files: List[UploadFile] = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), ocr_model: str = Query(DOC_TRANSLATE_DEFAULT_MODEL), gemini_route: str = Query(DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE)):
    allowed_ext = set(get_doc_translate_allowed_extensions())
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
    return {"models": get_pdf2docx_models(), "default_model": PDF2DOCX_DEFAULT_MODEL, "routes": get_gemini_routes(), "default_route": PDF2DOCX_DEFAULT_GEMINI_ROUTE}


@router.post("/pdf2docx")
async def run_pdf2docx(file: UploadFile = File(...), model: str = Query(PDF2DOCX_DEFAULT_MODEL), gemini_route: str = Query(PDF2DOCX_DEFAULT_GEMINI_ROUTE)):
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
async def run_pdf2docx_batch(files: List[UploadFile] = File(...), model: str = Query(PDF2DOCX_DEFAULT_MODEL), gemini_route: str = Query(PDF2DOCX_DEFAULT_GEMINI_ROUTE)):
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
