import asyncio
import json
import os
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Literal, Optional

from fastapi import APIRouter, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.background import BackgroundTask

from app.core.config import settings
from app.db.session import SessionLocal
from app.core.file_naming import build_user_visible_filename
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
    DOC_TRANSLATE_DEFAULT_MODE,
    DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE,
    DOC_TRANSLATE_DEFAULT_MODEL,
    DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE,
    get_doc_translate_allowed_extensions,
    get_doc_translate_models,
    get_doc_translate_modes,
    get_doc_translate_translation_engines,
    get_supported_languages,
    normalize_doc_translate_mode,
    normalize_doc_translate_translation_rules,
    normalize_doc_translate_translation_engine,
)
from app.service.drivers_license_service import get_drivers_license_config
from app.service.gemini_service import get_gemini_routes
from app.service.msg_convert_service import (
    MSG_CONVERT_ALLOWED_EXTENSIONS,
    MSG_CONVERT_DEFAULT_OUTPUT_FORMAT,
    MSG_CONVERT_MAX_FILES,
    OLE_COMPOUND_FILE_SIGNATURE,
    get_msg_convert_config as build_msg_convert_config,
    normalize_msg_output_format,
)
from app.service.number_check_service import (
    ALIGNMENT_EXTENSIONS,
    DIRECT_SOURCE_EXTENSIONS,
    HEADER_FOOTER_EXTENSIONS,
    NUMBER_CHECK_MODE_ALIGNMENT,
    NUMBER_CHECK_MODE_DIRECT,
    TARGET_EXTENSIONS,
    _get_task_progress as get_number_check_progress,
    get_number_check_default_mode,
    get_number_check_models,
)
from app.service.pdf2docx_service import (
    PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    PDF2DOCX_DEFAULT_LAYOUT_MODE,
    PDF2DOCX_DEFAULT_MODEL,
    get_pdf2docx_layout_modes,
    get_pdf2docx_models,
    normalize_pdf2docx_layout_mode,
)
from app.service.task_queue_service import UploadSizeLimitError, task_queue_service
from app.service.word_count_service import get_word_count_config as build_word_count_config

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


class BatchCancelBody(BaseModel):
    task_ids: List[str]


class TaskFeedbackBody(BaseModel):
    marked: bool = True
    category: Optional[str] = None
    note: Optional[str] = None


class WordCountSubmitBody(BaseModel):
    directory_path: str
    recursive: bool = True
    include_hidden: bool = False
    extensions: Optional[List[str]] = None
    ocr_mode: Literal["auto", "on", "off"] = "auto"
    ocr_model: Optional[str] = None


def _upload_file_size(file: UploadFile) -> int:
    declared_size = getattr(file, "size", None)
    if isinstance(declared_size, int) and declared_size >= 0:
        return declared_size
    handle = file.file
    try:
        position = handle.tell()
        handle.seek(0, os.SEEK_END)
        size = handle.tell()
        handle.seek(position, os.SEEK_SET)
        return max(0, int(size))
    except Exception:
        return 0


async def _validate_msg_upload(file: UploadFile) -> int:
    filename = file.filename or ""
    if Path(filename).suffix.lower() not in MSG_CONVERT_ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail=f"仅支持 .msg 文件：{filename or '未命名文件'}")
    size = _upload_file_size(file)
    await file.seek(0)
    signature = await file.read(len(OLE_COMPOUND_FILE_SIGNATURE))
    await file.seek(0)
    if signature != OLE_COMPOUND_FILE_SIGNATURE:
        raise HTTPException(status_code=400, detail=f"文件不是有效的 Outlook MSG：{filename}")
    return size


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
        "client_ip": task.client_ip,
        "status": task.status,
        "progress": task.progress,
        "message": task.message or "",
        "error": task.error_message,
        "cancel_requested": bool(task.cancel_requested),
        "batch": {
            "id": task.batch_id,
            "name": task.batch_name,
            "index": task.batch_index,
            "total": task.batch_total,
        } if task.batch_id else None,
        "feedback": {
            "marked": bool(task.feedback_marked),
            "category": task.feedback_category,
            "note": task.feedback_note or "",
            "marked_at": task.feedback_marked_at.isoformat() if task.feedback_marked_at else None,
        },
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
    candidate = Path(display_name or "").name.strip() or "output"
    for prefix in (getattr(task, "display_no", None), getattr(task, "task_id", None)):
        if prefix and candidate.startswith(f"{prefix}_"):
            candidate = candidate[len(prefix) + 1 :].lstrip(" _-") or candidate

    stem = Path(candidate).stem or "output"
    suffix = Path(candidate).suffix
    index = 2
    while candidate.casefold() in used_names:
        candidate = f"{stem}_{index}{suffix}"
        index += 1
    used_names.add(candidate.casefold())
    return candidate


def _is_legacy_default_archive_name(download_name: Optional[str]) -> bool:
    return (download_name or "").strip().lower() == "batch_outputs.zip"


def _fallback_archive_download_name() -> str:
    return f"batch_download_{datetime.now():%Y%m%d_%H%M%S}.zip"


def _task_primary_input_filename(task) -> Optional[str]:
    try:
        input_files = json.loads(task.input_files_json or "{}")
    except (TypeError, json.JSONDecodeError):
        input_files = {}

    if isinstance(input_files, dict):
        for key in (
            "original_filename",
            "single_filename",
            "target_filename",
            "translated_filename",
            "source_filename",
            "alignment_filename",
        ):
            value = input_files.get(key)
            if value:
                return Path(str(value)).name

    if getattr(task, "filename", None):
        return str(task.filename)
    return None


def _build_batch_archive_download_name(tasks: list) -> str:
    completed_tasks = [task for task in tasks if getattr(task, "status", None) == "done"] or tasks
    first_filename = next((name for name in (_task_primary_input_filename(task) for task in completed_tasks) if name), None)
    if not first_filename:
        return _fallback_archive_download_name()

    suffix = "批量结果" if len(completed_tasks) <= 1 else f"等{len(completed_tasks)}个文件"
    return build_user_visible_filename(first_filename, suffix=suffix, ext=".zip")


def _build_archive_response(archive_sources: list[tuple[Path, str]], download_name: str) -> FileResponse:
    if not archive_sources:
        raise HTTPException(status_code=400, detail="No matching output files found")

    temp_dir = BASE_DIR / "outputs" / "_tmp_batch_downloads"
    temp_dir.mkdir(parents=True, exist_ok=True)
    normalized_name = Path((download_name or _fallback_archive_download_name()).strip()).name
    normalized_name = normalized_name or _fallback_archive_download_name()
    if not normalized_name.lower().endswith(".zip"):
        normalized_name = f"{normalized_name}.zip"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip", dir=temp_dir) as temp_file:
        archive_path = Path(temp_file.name)

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for source_path, archive_name in archive_sources:
            zip_file.write(source_path, arcname=archive_name)

    return FileResponse(
        str(archive_path),
        filename=normalized_name,
        media_type="application/zip",
        background=BackgroundTask(lambda path=archive_path: path.unlink(missing_ok=True)),
    )


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
async def list_tasks(status: Optional[str] = Query(None), task_type: Optional[str] = Query(None), keyword: Optional[str] = Query(None), feedback_marked: Optional[bool] = Query(None), page: int = Query(1, ge=1), page_size: int = Query(20, ge=1, le=100)):
    with SessionLocal() as db:
        tasks, total = task_repo.list_tasks(db, status=status, task_type=task_type, keyword=keyword, feedback_marked=feedback_marked, page=page, page_size=page_size)
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


@router.post("/{task_id}/feedback")
async def update_task_feedback(task_id: str, body: TaskFeedbackBody):
    with SessionLocal() as db:
        task = task_repo.update_task_feedback(
            db,
            task_id,
            marked=body.marked,
            category=body.category,
            note=body.note,
        )
        if not task:
            raise HTTPException(status_code=404, detail="Task not found")
        return {"status": "ok", "task": _task_to_dict(task)}


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

    requested_archive_name = None if _is_legacy_default_archive_name(body.archive_name) else body.archive_name
    return _build_archive_response(archive_sources, requested_archive_name or _build_batch_archive_download_name(tasks))


@router.post("/batch-cancel")
async def batch_cancel(body: BatchCancelBody):
    if not body.task_ids:
        raise HTTPException(status_code=400, detail="At least one task_id is required")

    cancelled = 0
    skipped = 0
    missing = 0
    seen_task_ids: set[str] = set()
    with SessionLocal() as db:
        for task_id in body.task_ids:
            if not task_id or task_id in seen_task_ids:
                continue
            seen_task_ids.add(task_id)
            task = task_repo.get_task_by_task_id(db, task_id)
            if not task:
                missing += 1
                continue
            if task.status in {"done", "failed", "cancelled"}:
                skipped += 1
                continue
            task_repo.cancel_task(db, task.task_id)
            cancelled += 1
    return {"status": "ok", "cancelled": cancelled, "skipped": skipped, "missing": missing}


@router.get("/batch/{batch_id}/download")
async def batch_group_download(batch_id: str, extensions: Optional[str] = Query(None), archive_name: Optional[str] = Query(None)):
    normalized_extensions = None
    if extensions:
        normalized_extensions = _normalize_extensions([item.strip() for item in extensions.split(",") if item.strip()])

    with SessionLocal() as db:
        tasks = task_repo.list_tasks_by_batch_id(db, batch_id)

    if not tasks:
        raise HTTPException(status_code=404, detail="Batch not found")

    archive_sources: list[tuple[Path, str]] = []
    used_names: set[str] = set()
    for task in tasks:
        if task.status != "done":
            continue
        for item in _iter_task_output_files(task, normalized_extensions):
            archive_sources.append((item["resolved"], _build_archive_name(task, item["display_name"], used_names)))

    requested_archive_name = None if _is_legacy_default_archive_name(archive_name) else archive_name
    safe_batch_name = requested_archive_name or _build_batch_archive_download_name(tasks)
    return _build_archive_response(archive_sources, safe_batch_name)


@router.post("/batch/{batch_id}/cancel")
async def batch_group_cancel(batch_id: str):
    cancelled = 0
    skipped = 0
    with SessionLocal() as db:
        tasks = task_repo.list_tasks_by_batch_id(db, batch_id)
        if not tasks:
            raise HTTPException(status_code=404, detail="Batch not found")
        for task in tasks:
            if task.status in {"done", "failed", "cancelled"}:
                skipped += 1
                continue
            task_repo.cancel_task(db, task.task_id)
            cancelled += 1
    return {"status": "ok", "batch_id": batch_id, "cancelled": cancelled, "skipped": skipped}


@router.post("/number-check")
async def run_number_check(
    alignment_file: Optional[UploadFile] = File(None),
    source_file: Optional[UploadFile] = File(None),
    target_file: Optional[UploadFile] = File(None),
    source_hf_file: Optional[UploadFile] = File(None),
    original_file: Optional[UploadFile] = File(None),
    translated_file: Optional[UploadFile] = File(None),
    single_file: Optional[UploadFile] = File(None),
    mode: Optional[str] = Query(None),
    gemini_route: str = Query("openrouter"),
    model_name: str = Query("gemini-3.1-pro-preview"),
):
    resolved_mode = (mode or get_number_check_default_mode()).strip().lower()
    if resolved_mode in {"excel", "alignment_excel", "single"}:
        resolved_mode = NUMBER_CHECK_MODE_ALIGNMENT
    elif resolved_mode == "double":
        resolved_mode = NUMBER_CHECK_MODE_DIRECT

    source_upload = source_file or original_file
    target_upload = target_file or translated_file
    alignment_upload = alignment_file or single_file

    if resolved_mode == NUMBER_CHECK_MODE_ALIGNMENT:
        if alignment_upload is None:
            raise HTTPException(status_code=400, detail="对照 Excel 模式需要上传 alignment_file")
    elif resolved_mode == NUMBER_CHECK_MODE_DIRECT:
        if source_upload is None or target_upload is None:
            raise HTTPException(status_code=400, detail="原文+译文模式需要同时上传 source_file 和 target_file")
    else:
        raise HTTPException(status_code=400, detail=f"不支持的数字专检模式: {mode}")

    submit_result = await task_queue_service.submit_number_check_task(
        mode=resolved_mode,
        alignment_file=alignment_upload,
        source_file=source_upload,
        target_file=target_upload,
        source_hf_file=source_hf_file,
        gemini_route=gemini_route,
        model_name=model_name,
    )
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


@router.get("/number-check/config")
async def get_number_check_config():
    alignment_file_extensions = sorted(ALIGNMENT_EXTENSIONS)
    direct_file_extensions = sorted(DIRECT_SOURCE_EXTENSIONS)
    target_file_extensions = sorted(TARGET_EXTENSIONS)
    source_hf_file_extensions = sorted(HEADER_FOOTER_EXTENSIONS)
    return {
        "models": get_number_check_models(),
        "default_model": "gemini-3.1-pro-preview",
        "routes": get_gemini_routes(),
        "default_route": "openrouter",
        "default_mode": get_number_check_default_mode(),
        "modes": {
            NUMBER_CHECK_MODE_ALIGNMENT: {
                "label": "双语对照 Excel 模式",
                "description": "上传一个已对齐的原文/译文双语 Excel；如需生成修订版，可选上传译文文件，支持 DOCX、XLSX、PPTX、PDF。",
            },
            NUMBER_CHECK_MODE_DIRECT: {
                "label": "原文+译文双文件模式",
                "description": "直接上传原文和译文双文件，由新版数检程序抽取内容并生成报告，支持 DOCX、XLSX、PPTX、PDF。",
            },
        },
        "alignment_file_extensions": alignment_file_extensions,
        "direct_file_extensions": direct_file_extensions,
        "target_file_extensions": target_file_extensions,
        "source_hf_file_extensions": source_hf_file_extensions,
        "single_file_extensions": alignment_file_extensions,
        "double_file_extensions": direct_file_extensions,
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

    submit_result = await task_queue_service.submit_zhongfanyi_task(
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
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


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
    return {"models": {name: {"description": info["description"], "id": info["id"], "max_output": info["max_output"], "max_output_display": info.get("max_output_display")} for name, info in ALIGNMENT_MODELS.items()}, "routes": get_gemini_routes(), "default_route": "openrouter", "languages": {k: v["description"] for k, v in SUPPORTED_LANGUAGES.items()}, "thresholds": THRESHOLD_MAP, "buffer_chars": BUFFER_CHARS}


@router.post("/alignment")
async def run_alignment(original_file: UploadFile = File(...), translated_file: UploadFile = File(...), source_lang: str = Query("zh"), target_lang: str = Query("en"), model_name: str = Query("Google gemini-3-flash-preview"), gemini_route: str = Query("openrouter"), enable_post_split: bool = Query(True), threshold_2: int = Query(25000), threshold_3: int = Query(50000), threshold_4: int = Query(75000), threshold_5: int = Query(100000), threshold_6: int = Query(125000), threshold_7: int = Query(150000), threshold_8: int = Query(175000), buffer_chars: int = Query(2000)):
    allowed_ext = {".docx", ".doc", ".pptx", ".xlsx", ".xls"}
    if os.path.splitext(original_file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported original file format")
    if os.path.splitext(translated_file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported translated file format")
    submit_result = await task_queue_service.submit_alignment_task(original_file=original_file, translated_file=translated_file, source_lang=source_lang, target_lang=target_lang, model_name=model_name, gemini_route=gemini_route, enable_post_split=enable_post_split, threshold_2=threshold_2, threshold_3=threshold_3, threshold_4=threshold_4, threshold_5=threshold_5, threshold_6=threshold_6, threshold_7=threshold_7, threshold_8=threshold_8, buffer_chars=buffer_chars)
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


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


@router.get("/word-count/config")
async def get_word_count_page_config():
    return build_word_count_config()


@router.post("/word-count")
async def submit_word_count(body: WordCountSubmitBody):
    try:
        submit_result = await task_queue_service.submit_word_count_task(
            directory_path=body.directory_path,
            recursive=body.recursive,
            include_hidden=body.include_hidden,
            extensions=body.extensions,
            ocr_mode=body.ocr_mode,
            ocr_model=body.ocr_model,
        )
    except (ValueError, FileNotFoundError, PermissionError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "status": "ACCEPTED",
        "task_id": submit_result.task_id,
        "message": "Task submitted",
        "deduped": submit_result.deduped,
    }


@router.post("/word-count/upload")
async def submit_word_count_upload(
    file: UploadFile = File(...),
    ocr_mode: Literal["auto", "on", "off"] = Form("auto"),
    ocr_model: Optional[str] = Form(None),
):
    try:
        submit_result = await task_queue_service.submit_word_count_upload_task(
            file=file,
            ocr_mode=ocr_mode,
            ocr_model=ocr_model,
        )
    except UploadSizeLimitError as exc:
        raise HTTPException(status_code=413, detail=str(exc)) from exc
    except (ValueError, FileNotFoundError, PermissionError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "status": "ACCEPTED",
        "task_id": submit_result.task_id,
        "message": "Task submitted",
        "deduped": submit_result.deduped,
    }


@router.get("/word-count/status/{task_id}")
async def get_word_count_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
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
    submit_result = await task_queue_service.submit_drivers_license_task(files=files, processing_mode=processing_mode)
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


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

    submit_result = await task_queue_service.submit_business_licence_task(
        file=file,
        model=model,
        gemini_route=BUSINESS_LICENCE_DEFAULT_ROUTE,
        parsed_data=parsed_data,
        company_name_override=company_name_override,
    )
    return {
        "status": "ACCEPTED",
        "task_id": submit_result.task_id,
        "message": "Task submitted",
        "gemini_route": BUSINESS_LICENCE_DEFAULT_ROUTE,
        "deduped": submit_result.deduped,
    }


@router.get("/business-licence/status/{task_id}")
async def get_business_licence_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/doc-translate/config")
async def get_doc_translate_config():
    return {
        "models": get_doc_translate_models(),
        "default_model": DOC_TRANSLATE_DEFAULT_MODEL,
        "routes": get_gemini_routes(),
        "default_route": DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE,
        "languages": get_supported_languages(),
        "translate_modes": get_doc_translate_modes(),
        "default_translate_mode": DOC_TRANSLATE_DEFAULT_MODE,
        "translation_engines": get_doc_translate_translation_engines(),
        "default_translation_engine": DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE,
        "allowed_extensions": get_doc_translate_allowed_extensions(),
    }


@router.post("/doc-translate")
async def submit_doc_translate(file: UploadFile = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), translate_mode: str = Query(DOC_TRANSLATE_DEFAULT_MODE), ocr_model: str = Query(DOC_TRANSLATE_DEFAULT_MODEL), gemini_route: str = Query(DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE), translation_engine: str = Query(DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE), translation_rules: str = Form("")):
    allowed_ext = set(get_doc_translate_allowed_extensions())
    if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported file format")
    try:
        translate_mode = normalize_doc_translate_mode(translate_mode)
        translation_engine = normalize_doc_translate_translation_engine(translation_engine)
        translation_rules = normalize_doc_translate_translation_rules(translation_rules)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    submit_result = await task_queue_service.submit_doc_translate_task(file=file, source_lang=source_lang, target_langs=target_langs, translate_mode=translate_mode, ocr_model=ocr_model, gemini_route=gemini_route, translation_engine=translation_engine, translation_rules=translation_rules)
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


@router.post("/doc-translate/batch")
async def submit_doc_translate_batch(files: List[UploadFile] = File(...), source_lang: str = Query("zh"), target_langs: str = Query("en"), translate_mode: str = Query(DOC_TRANSLATE_DEFAULT_MODE), ocr_model: str = Query(DOC_TRANSLATE_DEFAULT_MODEL), gemini_route: str = Query(DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE), translation_engine: str = Query(DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE), translation_rules: str = Form("")):
    allowed_ext = set(get_doc_translate_allowed_extensions())
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    try:
        translate_mode = normalize_doc_translate_mode(translate_mode)
        translation_engine = normalize_doc_translate_translation_engine(translation_engine)
        translation_rules = normalize_doc_translate_translation_rules(translation_rules)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    batch_id = str(uuid.uuid4())
    batch_name = f"通用证件批量翻译_{batch_id[:8]}.zip"
    batch_total = len(files)
    results = []
    for index, file in enumerate(files, start=1):
        if os.path.splitext(file.filename or "")[1].lower() not in allowed_ext:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": "Unsupported file format"})
            continue
        try:
            submit_result = await task_queue_service.submit_doc_translate_task(file=file, source_lang=source_lang, target_langs=target_langs, translate_mode=translate_mode, ocr_model=ocr_model, gemini_route=gemini_route, translation_engine=translation_engine, translation_rules=translation_rules, batch_id=batch_id, batch_name=batch_name, batch_index=index, batch_total=batch_total)
            results.append({"filename": file.filename, "task_id": submit_result.task_id, "status": "ACCEPTED", "deduped": submit_result.deduped, "batch_id": batch_id, "batch_index": index, "batch_total": batch_total})
        except Exception as exc:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": str(exc)})
    return {"status": "ACCEPTED", "tasks": results, "total": len(results), "batch_id": batch_id, "batch_name": batch_name}


@router.get("/doc-translate/status/{task_id}")
async def get_doc_translate_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/msg-convert/config")
async def get_msg_convert_config():
    return build_msg_convert_config()


@router.post("/msg-convert")
async def run_msg_convert(
    file: UploadFile = File(...),
    output_format: str = Query(MSG_CONVERT_DEFAULT_OUTPUT_FORMAT),
):
    try:
        normalized_format = normalize_msg_output_format(output_format)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    size = await _validate_msg_upload(file)
    max_bytes = max(1, int(settings.MSG_CONVERT_UPLOAD_MAX_MB or 95)) * 1024 * 1024
    if size > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"MSG 文件超过 {settings.MSG_CONVERT_UPLOAD_MAX_MB:g} MB 上传限制",
        )
    try:
        submit_result = await task_queue_service.submit_msg_convert_task(
            file=file,
            output_format=normalized_format,
        )
    except UploadSizeLimitError as exc:
        raise HTTPException(status_code=413, detail=str(exc)) from exc
    except (ValueError, FileNotFoundError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "status": "ACCEPTED",
        "task_id": submit_result.task_id,
        "message": "Task submitted",
        "deduped": submit_result.deduped,
    }


@router.post("/msg-convert/batch")
async def run_msg_convert_batch(
    files: List[UploadFile] = File(...),
    output_format: str = Query(MSG_CONVERT_DEFAULT_OUTPUT_FORMAT),
):
    if not files:
        raise HTTPException(status_code=400, detail="至少需要上传一个 MSG 文件")
    if len(files) > MSG_CONVERT_MAX_FILES:
        raise HTTPException(status_code=400, detail=f"单次最多上传 {MSG_CONVERT_MAX_FILES} 个 MSG 文件")
    try:
        normalized_format = normalize_msg_output_format(output_format)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    sizes = [await _validate_msg_upload(file) for file in files]
    total_size = sum(sizes)
    max_bytes = max(1, int(settings.MSG_CONVERT_UPLOAD_MAX_MB or 95)) * 1024 * 1024
    if total_size > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=(
                f"单次上传总大小不能超过 {settings.MSG_CONVERT_UPLOAD_MAX_MB:g} MB，"
                f"当前约 {total_size / (1024 * 1024):.1f} MB"
            ),
        )

    batch_id = str(uuid.uuid4())
    batch_name = f"MSG转文档批量结果_{batch_id[:8]}.zip"
    batch_total = len(files)
    results = []
    for index, file in enumerate(files, start=1):
        try:
            submit_result = await task_queue_service.submit_msg_convert_task(
                file=file,
                output_format=normalized_format,
                batch_id=batch_id,
                batch_name=batch_name,
                batch_index=index,
                batch_total=batch_total,
            )
            results.append(
                {
                    "filename": file.filename,
                    "task_id": submit_result.task_id,
                    "status": "ACCEPTED",
                    "deduped": submit_result.deduped,
                    "batch_id": batch_id,
                    "batch_index": index,
                    "batch_total": batch_total,
                }
            )
        except Exception as exc:
            results.append(
                {
                    "filename": file.filename,
                    "task_id": None,
                    "status": "FAILED",
                    "error": str(exc),
                }
            )
    return {
        "status": "ACCEPTED",
        "tasks": results,
        "total": len(results),
        "batch_id": batch_id,
        "batch_name": batch_name,
        "output_format": normalized_format,
    }


@router.get("/msg-convert/status/{task_id}")
async def get_msg_convert_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task


@router.get("/pdf2docx/config")
async def get_pdf2docx_config():
    return {"models": get_pdf2docx_models(), "default_model": PDF2DOCX_DEFAULT_MODEL, "routes": get_gemini_routes(), "default_route": PDF2DOCX_DEFAULT_GEMINI_ROUTE, "layout_modes": get_pdf2docx_layout_modes(), "default_layout_mode": PDF2DOCX_DEFAULT_LAYOUT_MODE}


@router.post("/pdf2docx")
async def run_pdf2docx(file: UploadFile = File(...), model: str = Query(PDF2DOCX_DEFAULT_MODEL), gemini_route: str = Query(PDF2DOCX_DEFAULT_GEMINI_ROUTE), layout_mode: str = Query(PDF2DOCX_DEFAULT_LAYOUT_MODE)):
    allowed_ext = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail="Unsupported file format")
    try:
        normalized_layout_mode = normalize_pdf2docx_layout_mode(layout_mode)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    submit_result = await task_queue_service.submit_pdf2docx_task(file=file, model=model, gemini_route=gemini_route, layout_mode=normalized_layout_mode)
    return {"status": "ACCEPTED", "task_id": submit_result.task_id, "message": "Task submitted", "deduped": submit_result.deduped}


@router.post("/pdf2docx/batch")
async def run_pdf2docx_batch(files: List[UploadFile] = File(...), model: str = Query(PDF2DOCX_DEFAULT_MODEL), gemini_route: str = Query(PDF2DOCX_DEFAULT_GEMINI_ROUTE), layout_mode: str = Query(PDF2DOCX_DEFAULT_LAYOUT_MODE)):
    allowed_ext = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
    if not files:
        raise HTTPException(status_code=400, detail="At least one file is required")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="Too many files (max 50)")
    try:
        normalized_layout_mode = normalize_pdf2docx_layout_mode(layout_mode)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    batch_id = str(uuid.uuid4())
    batch_name = f"不可编辑预处理批量结果_{batch_id[:8]}.zip"
    batch_total = len(files)
    results = []
    for index, file in enumerate(files, start=1):
        ext = os.path.splitext(file.filename or "")[1].lower()
        if ext not in allowed_ext:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": "Unsupported file format"})
            continue
        try:
            submit_result = await task_queue_service.submit_pdf2docx_task(file=file, model=model, gemini_route=gemini_route, layout_mode=normalized_layout_mode, batch_id=batch_id, batch_name=batch_name, batch_index=index, batch_total=batch_total)
            results.append({"filename": file.filename, "task_id": submit_result.task_id, "status": "ACCEPTED", "deduped": submit_result.deduped, "batch_id": batch_id, "batch_index": index, "batch_total": batch_total})
        except Exception as exc:
            results.append({"filename": file.filename, "task_id": None, "status": "FAILED", "error": str(exc)})
    return {"status": "ACCEPTED", "tasks": results, "total": len(results), "batch_id": batch_id, "batch_name": batch_name}


@router.get("/pdf2docx/status/{task_id}")
async def get_pdf2docx_status(task_id: str):
    queue_task = task_queue_service.get_task_status(task_id)
    if not queue_task:
        raise HTTPException(status_code=404, detail="Task not found")
    return queue_task
