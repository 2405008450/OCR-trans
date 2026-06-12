import asyncio
import contextlib
import importlib.util
import io
import json
import logging
import os
import re
import shutil
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import UploadFile

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename, ensure_unique_path
from app.service.gemini_service import (
    GEMINI_ROUTE_OPENROUTER,
    resolve_model_for_route,
)
from app.service.libreoffice_service import convert_doc_to_docx_via_libreoffice


logger = logging.getLogger("app.number_check")

_task_progress: Dict[str, Dict[str, Any]] = {}
_specialist_import_lock = threading.Lock()

REPO_ROOT = Path(__file__).resolve().parents[2]
NUMBER_CHECK_LATEST_ROOT = REPO_ROOT / "专检" / "数检_程序-AI"
NUMBER_CHECK_MAIN_FILE = NUMBER_CHECK_LATEST_ROOT / "main.py"

NUMBER_CHECK_MODE_ALIGNMENT = "alignment"
NUMBER_CHECK_MODE_DIRECT = "direct"

ALIGNMENT_EXTENSIONS = {".xlsx"}
DIRECT_SOURCE_EXTENSIONS = {".docx", ".doc", ".xlsx", ".pptx"}
TARGET_EXTENSIONS = {".docx", ".doc", ".xlsx", ".pptx", ".pdf"}
HEADER_FOOTER_EXTENSIONS = {".docx", ".doc"}

REPORT_FILE_KEYS = {
    "align_body.json": "body_json",
    "align_body_errors.json": "body_errors_json",
    "align_body_flat_errors.json": "body_flat_errors_json",
    "align_header.json": "header_json",
    "align_footer.json": "footer_json",
}

_CLIENT_LOG_SANITIZE_RULES: List[tuple] = [
    (re.compile(r"\[config\]\s+.*"), "[config] 任务已初始化"),
    (re.compile(r"(?i)(route=)\S+"), r"\1***"),
    (re.compile(r"(?i)(model=)\S+"), r"\1***"),
    (re.compile(r"(?i)gemini[-/][\w.\-]+"), "[AI模型]"),
    (re.compile(r"\bopenrouter\b", re.IGNORECASE), "[路线]"),
    (re.compile(r"\bgoogle\b", re.IGNORECASE), "[路线]"),
]


NUMBER_CHECK_MODELS: Dict[str, Dict[str, str]] = {
    "gemini-3-flash-preview": {
        "label": "快速版V2",
        "description": "速度更快，适合常规数字核对场景。",
    },
    "gemini-3.5-flash": {
        "label": "新模型",
        "description": "OpenRouter 新模型，适合常规数字核对场景。",
    },
    "gemini-3.1-pro-preview": {
        "label": "增强版V2",
        "description": "推理更强，适合复杂编号和上下文判断场景。",
    },
}

NUMBER_CHECK_MODEL_ALIASES: Dict[str, str] = {
    "gemini-3-flash-preview": "gemini-3-flash-preview",
    "google/gemini-3-flash-preview": "gemini-3-flash-preview",
    "gemini 3 flash preview": "gemini-3-flash-preview",
    "gemini-3.5-flash": "gemini-3.5-flash",
    "google/gemini-3.5-flash": "gemini-3.5-flash",
    "gemini 3.5 flash": "gemini-3.5-flash",
    "gemini-3.1-pro-preview": "gemini-3.1-pro-preview",
    "google/gemini-3.1-pro-preview": "gemini-3.1-pro-preview",
    "gemini 3.1 pro preview": "gemini-3.1-pro-preview",
}


def get_number_check_models() -> Dict[str, Dict[str, str]]:
    return NUMBER_CHECK_MODELS


def get_number_check_default_mode() -> str:
    return NUMBER_CHECK_MODE_ALIGNMENT


def normalize_number_check_model(model_name: Optional[str]) -> str:
    candidate = (model_name or "gemini-3.1-pro-preview").strip()
    key = NUMBER_CHECK_MODEL_ALIASES.get(candidate.lower(), candidate)
    if key not in NUMBER_CHECK_MODELS:
        raise ValueError(f"不支持的数字专检模型: {model_name}")
    return key


def _normalize_number_check_mode(mode: Optional[str]) -> str:
    candidate = (mode or NUMBER_CHECK_MODE_ALIGNMENT).strip().lower()
    aliases = {
        "excel": NUMBER_CHECK_MODE_ALIGNMENT,
        "alignment_excel": NUMBER_CHECK_MODE_ALIGNMENT,
        "single": NUMBER_CHECK_MODE_ALIGNMENT,
        "double": NUMBER_CHECK_MODE_DIRECT,
    }
    candidate = aliases.get(candidate, candidate)
    if candidate not in {NUMBER_CHECK_MODE_ALIGNMENT, NUMBER_CHECK_MODE_DIRECT}:
        raise ValueError(f"不支持的数字专检模式: {mode}")
    return candidate


def _sanitize_client_log(line: str) -> str:
    for pattern, replacement in _CLIENT_LOG_SANITIZE_RULES:
        line = pattern.sub(replacement, line)
    return line


def _append_stream_log(task_id: str, message: str) -> None:
    if task_id not in _task_progress:
        return
    line = _sanitize_client_log((message or "").rstrip())
    if not line:
        return
    current = _task_progress[task_id].get("stream_log", "")
    lines = current.splitlines() if current else []
    if lines and lines[-1] == line:
        return
    combined = f"{current}\n{line}" if current else line
    _task_progress[task_id]["stream_log"] = combined[-50000:]


def _emit_log(task_id: str, message: str, level: str = "info") -> None:
    _append_stream_log(task_id, message)
    log_line = f"[number-check][{task_id}] {message}"
    if level == "error":
        logger.error(log_line)
    elif level == "warning":
        logger.warning(log_line)
    else:
        logger.info(log_line)


def _init_task_progress(task_id: str, total_steps: int = 6) -> None:
    _task_progress[task_id] = {
        "status": "running",
        "current_step": 0,
        "total_steps": total_steps,
        "message": "初始化任务...",
        "progress": 0,
        "details": [],
        "stream_log": "",
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
    }


def _update_progress(
    task_id: str,
    current_step: int,
    total_steps: int,
    message: str,
    details: Optional[List[str]] = None,
) -> None:
    if task_id not in _task_progress:
        return
    progress = int((current_step / total_steps) * 100)
    detail_list = details or []
    _task_progress[task_id].update(
        {
            "current_step": current_step,
            "total_steps": total_steps,
            "message": message,
            "progress": progress,
            "details": detail_list,
            "updated_at": datetime.now().isoformat(),
        }
    )
    _emit_log(task_id, f"[{progress:>3}%] {message}")
    for item in detail_list:
        _emit_log(task_id, f"  - {item}")


def _complete_task(task_id: str, result: Optional[Dict[str, Any]] = None, error: Optional[str] = None) -> None:
    if task_id not in _task_progress:
        return
    _task_progress[task_id].update(
        {
            "status": "done" if not error else "failed",
            "progress": 100,
            "message": "处理完成" if not error else f"处理失败: {error}",
            "result": result,
            "error": error,
            "updated_at": datetime.now().isoformat(),
        }
    )
    _emit_log(task_id, "[done] 处理完成" if not error else f"[error] {error}", level="error" if error else "info")


def _get_task_progress(task_id: str) -> Optional[Dict[str, Any]]:
    return _task_progress.get(task_id)


def _cleanup_task(task_id: str) -> None:
    _task_progress.pop(task_id, None)


class _TaskLogWriter(io.TextIOBase):
    def __init__(self, task_id: str, original_stream) -> None:
        self.task_id = task_id
        self.original_stream = original_stream
        self._buffer = ""

    def writable(self) -> bool:
        return True

    def write(self, text: str) -> int:
        if not text:
            return 0
        self.original_stream.write(text)
        self.original_stream.flush()
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            _append_stream_log(self.task_id, line)
        return len(text)

    def flush(self) -> None:
        self.original_stream.flush()
        if self._buffer.strip():
            _append_stream_log(self.task_id, self._buffer)
        self._buffer = ""


def _validate_upload(file: UploadFile, label: str, allowed: set[str]) -> None:
    ext = Path(file.filename or "").suffix.lower()
    if ext not in allowed:
        allowed_text = " / ".join(sorted(allowed))
        raise ValueError(f"{label} 仅支持 {allowed_text}，当前为 {file.filename}")


async def _save_upload(file: UploadFile, target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "wb") as f:
        f.write(await file.read())
    return target


def _convert_doc_input_if_needed(source: Path, work_dir: Path, role: str) -> Path:
    if source.suffix.lower() != ".doc":
        return source
    work_dir.mkdir(parents=True, exist_ok=True)
    target = work_dir / f"{role}.docx"
    return Path(convert_doc_to_docx_via_libreoffice(source, target))


def _set_llm_env(model_name: str) -> str:
    api_key = (
        settings.OPENROUTER_API_KEY
        or os.getenv("API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or os.getenv("OPENROUTER_API_KEY")
    )
    if not api_key:
        raise ValueError("未配置 OPENROUTER_API_KEY、OPENAI_API_KEY 或 API_KEY，无法执行数字专检。")

    base_url = (
        os.getenv("BASE_URL")
        or settings.OPENROUTER_BASE_URL
        or os.getenv("OPENAI_BASE_URL")
        or os.getenv("OPENROUTER_BASE_URL")
    )
    resolved_model_name = resolve_model_for_route(model_name, GEMINI_ROUTE_OPENROUTER)

    os.environ["API_KEY"] = api_key
    os.environ["OPENAI_API_KEY"] = api_key
    os.environ["OPENROUTER_API_KEY"] = api_key
    if base_url:
        os.environ["BASE_URL"] = base_url
        os.environ["OPENAI_BASE_URL"] = base_url
        os.environ["OPENROUTER_BASE_URL"] = base_url
    os.environ["NUMBER_CHECK_MODEL"] = resolved_model_name
    return resolved_model_name


def _clear_specialist_module_cache() -> None:
    names = {
        "_number_check_ai_main",
        "main",
        "ai_check",
        "program_check",
        "extract_values",
        "report_generator",
        "normalizer",
        "normalizer_total",
        "full_content",
        "extract_any",
        "header_extractor",
        "footer_extractor",
        "body_extractor",
        "replace_revision",
        "replace_clean",
        "revision",
        "clean_replace_duplicates",
        "excel",
        "ppt",
        "pdf",
        "txt",
        "backup_copy",
        "num_checker",
    }
    for key in list(sys.modules.keys()):
        if key in names or any(key.startswith(f"{name}.") for name in names):
            del sys.modules[key]


def _load_latest_main_module():
    if not NUMBER_CHECK_MAIN_FILE.exists():
        raise FileNotFoundError(f"未找到新版数字专检主程序: {NUMBER_CHECK_MAIN_FILE}")

    root = str(NUMBER_CHECK_LATEST_ROOT)
    while root in sys.path:
        sys.path.remove(root)
    sys.path.insert(0, root)
    _clear_specialist_module_cache()

    spec = importlib.util.spec_from_file_location("_number_check_ai_main", NUMBER_CHECK_MAIN_FILE)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载数字专检主程序: {NUMBER_CHECK_MAIN_FILE}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["_number_check_ai_main"] = module
    spec.loader.exec_module(module)
    return module


def _relative_output_path(path: Path, output_dir: Path) -> str:
    return f"outputs/number_check/{output_dir.name}/{path.relative_to(output_dir).as_posix()}"


def _copy_if_needed(source: Path, target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        same_file = source.resolve() == target.resolve()
    except Exception:
        same_file = str(source) == str(target)
    if not same_file:
        shutil.copy2(source, target)
    return target


def _latest_matching_file(output_dir: Path, pattern: str) -> Optional[Path]:
    matches = [p for p in output_dir.glob(pattern) if p.is_file()]
    if not matches:
        return None
    return max(matches, key=lambda item: item.stat().st_mtime)


def _count_flat_errors(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    return len(payload) if isinstance(payload, list) else 0


def _count_ai_errors(rows: List[Dict[str, Any]]) -> int:
    total = 0
    for row in rows or []:
        try:
            total += int(row.get("AI错误数量", 0) or 0)
        except (TypeError, ValueError):
            continue
    return total


def _collect_outputs(
    output_dir: Path,
    revised_output_path: Optional[Path],
    body_rows: List[Dict[str, Any]],
    header_rows: List[Dict[str, Any]],
    footer_rows: List[Dict[str, Any]],
) -> tuple[Dict[str, str], Dict[str, int], Dict[str, str]]:
    reports: Dict[str, str] = {}
    files: Dict[str, str] = {}

    report_excel = _latest_matching_file(output_dir, "*_output.xlsx")
    if report_excel:
        reports["report_excel"] = _relative_output_path(report_excel, output_dir)

    alignment_excel = _latest_matching_file(output_dir, "*_alignment.xlsx")
    if alignment_excel:
        reports["alignment_excel"] = _relative_output_path(alignment_excel, output_dir)

    alignment_json = _latest_matching_file(output_dir, "*_alignment.json")
    if alignment_json:
        reports["alignment_json"] = _relative_output_path(alignment_json, output_dir)

    for filename, key in REPORT_FILE_KEYS.items():
        path = output_dir / filename
        if path.exists():
            reports[key] = _relative_output_path(path, output_dir)

    flat_errors_path = output_dir / "align_body_flat_errors.json"
    body_issues = _count_flat_errors(flat_errors_path) or _count_ai_errors(body_rows)
    header_issues = _count_ai_errors(header_rows)
    footer_issues = _count_ai_errors(footer_rows)
    report_counts = {
        "body_issues": body_issues,
        "header_issues": header_issues,
        "footer_issues": footer_issues,
        "total_issues": body_issues + header_issues + footer_issues,
    }

    if revised_output_path and revised_output_path.exists():
        files["revised_file"] = _relative_output_path(revised_output_path, output_dir)
        if revised_output_path.suffix.lower() == ".docx":
            files["corrected_docx"] = files["revised_file"]

    return reports, report_counts, files


def _run_latest_number_check_sync(
    *,
    task_id: str,
    mode: str,
    alignment_path: Optional[Path],
    source_path: Optional[Path],
    target_path: Optional[Path],
    source_hf_path: Optional[Path],
    output_dir: Path,
    gemini_route: str,
    model_name: str,
    alignment_filename: Optional[str],
    source_filename: Optional[str],
    target_filename: Optional[str],
) -> Dict[str, Any]:
    total_steps = 6
    effective_route = GEMINI_ROUTE_OPENROUTER
    if gemini_route != effective_route:
        _emit_log(task_id, f"[config] 数字专检固定使用 route={effective_route}，已忽略 route={gemini_route}", level="warning")

    _update_progress(task_id, 2, total_steps, "正在准备新版数字专检运行环境...")
    converted_dir = output_dir / "converted_inputs"
    if source_path:
        source_path = _convert_doc_input_if_needed(source_path, converted_dir, "source")
    if target_path:
        target_path = _convert_doc_input_if_needed(target_path, converted_dir, "target")
    if source_hf_path:
        source_hf_path = _convert_doc_input_if_needed(source_hf_path, converted_dir, "source_header_footer")

    resolved_model_name = _set_llm_env(model_name)
    _emit_log(task_id, f"[config] mode={mode}, route={effective_route}, model={resolved_model_name}")

    _update_progress(task_id, 3, total_steps, "正在加载新版数检主程序...")
    with _specialist_import_lock:
        main_module = _load_latest_main_module()

    revised_output_path: Optional[Path] = None
    if target_path:
        revised_output_path = ensure_unique_path(
            output_dir / build_user_visible_filename(
                target_filename or target_path.name,
                suffix="number_checked",
                ext=target_path.suffix.lower() or ".docx",
            )
        )

    run_kwargs = {
        "alignment_path": str(alignment_path) if alignment_path else None,
        "output_dir": str(output_dir),
        "src_docx_path": str(source_path) if source_path else None,
        "tgt_docx_path": str(target_path) if target_path else None,
        "src_hf_path": str(source_hf_path) if source_hf_path else None,
        "docx_path": str(target_path) if target_path else None,
        "revised_docx_path": str(revised_output_path) if revised_output_path else None,
        "revision_author": "数值检查",
        "use_total_normalizer": True,
        "ai_check_all": False,
    }

    _update_progress(
        task_id,
        4,
        total_steps,
        "正在执行规则检查和 AI 复核...",
        [
            "模式: 对照 Excel" if mode == NUMBER_CHECK_MODE_ALIGNMENT else "模式: 原文+译文",
            f"模型: {resolved_model_name}",
        ],
    )

    stdout_writer = _TaskLogWriter(task_id, sys.stdout)
    stderr_writer = _TaskLogWriter(task_id, sys.stderr)
    try:
        with contextlib.redirect_stdout(stdout_writer), contextlib.redirect_stderr(stderr_writer):
            body_rows, header_rows, footer_rows = main_module.run(**run_kwargs)
    finally:
        stdout_writer.flush()
        stderr_writer.flush()

    _update_progress(task_id, 5, total_steps, "正在整理输出文件...")
    reports, report_counts, files = _collect_outputs(
        output_dir,
        revised_output_path,
        body_rows or [],
        header_rows or [],
        footer_rows or [],
    )

    result: Dict[str, Any] = {
        "task_id": task_id,
        "mode": mode,
        "alignment_filename": alignment_filename,
        "source_filename": source_filename,
        "target_filename": target_filename,
        "model_name": resolved_model_name,
        "reports": reports,
        "report_counts": report_counts,
        "stats": {
            "total_issues": report_counts["total_issues"],
            "body_issues": report_counts["body_issues"],
            "header_issues": report_counts["header_issues"],
            "footer_issues": report_counts["footer_issues"],
        },
        "summary": "新版数字专检已完成，报告和修订文件已生成。",
    }
    result.update(files)

    _update_progress(task_id, 6, total_steps, "数字专检完成")
    _complete_task(task_id, result=result)
    return result


async def run_number_check_task(
    alignment_file: Optional[UploadFile] = None,
    source_file: Optional[UploadFile] = None,
    target_file: Optional[UploadFile] = None,
    source_hf_file: Optional[UploadFile] = None,
    original_file: Optional[UploadFile] = None,
    translated_file: Optional[UploadFile] = None,
    single_file: Optional[UploadFile] = None,
    mode: str = NUMBER_CHECK_MODE_ALIGNMENT,
    task_id: str = "",
    display_no: Optional[str] = None,
    gemini_route: str = "openrouter",
    model_name: str = "gemini-3.1-pro-preview",
) -> Dict[str, Any]:
    normalized_mode = _normalize_number_check_mode(mode)
    model_name = normalize_number_check_model(model_name)

    # 兼容旧调用名：original/translated 对应新版 source/target。
    source_file = source_file or original_file
    target_file = target_file or translated_file
    if single_file and not alignment_file:
        alignment_file = single_file

    if not task_id:
        task_id = str(uuid.uuid4())

    _init_task_progress(task_id, total_steps=6)
    _emit_log(task_id, f"[config] mode={normalized_mode}, route={gemini_route}, model={model_name}")

    try:
        if normalized_mode == NUMBER_CHECK_MODE_ALIGNMENT:
            if alignment_file is None:
                raise ValueError("对照 Excel 模式缺少 alignment_file")
            _validate_upload(alignment_file, "对照文件", ALIGNMENT_EXTENSIONS)
            if target_file:
                _validate_upload(target_file, "待修订译文文件", TARGET_EXTENSIONS)
            if source_hf_file:
                _validate_upload(source_hf_file, "页眉页脚原文文件", HEADER_FOOTER_EXTENSIONS)
        else:
            if source_file is None or target_file is None:
                raise ValueError("原文+译文模式缺少 source_file 或 target_file")
            _validate_upload(source_file, "原文文件", DIRECT_SOURCE_EXTENSIONS)
            _validate_upload(target_file, "译文文件", TARGET_EXTENSIONS - {".pdf"})

        folder_name = display_no or task_id
        upload_dir = Path(settings.UPLOAD_DIR) / "number_check" / folder_name
        output_dir = Path(settings.OUTPUT_DIR) / "number_check" / folder_name
        upload_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        _update_progress(task_id, 1, 6, "正在保存上传文件...")

        saved_alignment: Optional[Path] = None
        saved_source: Optional[Path] = None
        saved_target: Optional[Path] = None
        saved_source_hf: Optional[Path] = None

        if alignment_file:
            ext = Path(alignment_file.filename or "alignment.xlsx").suffix.lower() or ".xlsx"
            saved_alignment = await _save_upload(alignment_file, upload_dir / f"alignment{ext}")
        if source_file:
            ext = Path(source_file.filename or "source.docx").suffix.lower() or ".docx"
            saved_source = await _save_upload(source_file, upload_dir / f"source{ext}")
        if target_file:
            ext = Path(target_file.filename or "target.docx").suffix.lower() or ".docx"
            saved_target = await _save_upload(target_file, upload_dir / f"target{ext}")
        if source_hf_file:
            ext = Path(source_hf_file.filename or "source_hf.docx").suffix.lower() or ".docx"
            saved_source_hf = await _save_upload(source_hf_file, upload_dir / f"source_hf{ext}")

        return await asyncio.to_thread(
            _run_latest_number_check_sync,
            task_id=task_id,
            mode=normalized_mode,
            alignment_path=saved_alignment,
            source_path=saved_source,
            target_path=saved_target,
            source_hf_path=saved_source_hf,
            output_dir=output_dir,
            gemini_route=gemini_route,
            model_name=model_name,
            alignment_filename=alignment_file.filename if alignment_file else None,
            source_filename=source_file.filename if source_file else None,
            target_filename=target_file.filename if target_file else None,
        )
    except Exception as exc:
        _emit_log(task_id, f"[error] 任务失败: {type(exc).__name__}: {exc}", level="error")
        _complete_task(task_id, error=str(exc))
        raise
