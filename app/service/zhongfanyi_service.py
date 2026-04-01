# -*- coding: utf-8 -*-
"""
中翻译专检：对接 专检/中翻译/main.py，支持双文件模式、单文件双语模式与 Excel 报告导出。
"""
import importlib.machinery
import os
import shutil
import sys
import threading
import types
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from app.core.config import settings
from app.service.gemini_service import GEMINI_ROUTE_OPENROUTER, ensure_gemini_route_configured


ZHONGFANYI_DEFAULT_GEMINI_ROUTE = GEMINI_ROUTE_OPENROUTER
ZHONGFANYI_MODE_DOUBLE = "double"
ZHONGFANYI_MODE_SINGLE = "single"
ZHONGFANYI_SINGLE_FILE_EXTENSIONS = [".docx", ".doc", ".pdf", ".xlsx", ".xls", ".pptx"]
ZHONGFANYI_DOUBLE_FILE_EXTENSIONS = [".docx", ".doc", ".pdf", ".xlsx", ".xls", ".pptx"]
SECTION_DIR_NAMES = {"正文": "zhengwen", "页眉": "yemei", "页脚": "yejiao"}
SECTION_COUNT_KEYS = {"正文": "body_issues", "页眉": "header_issues", "页脚": "footer_issues"}

_task_progress: Dict[str, Dict[str, Any]] = {}
_specialist_import_lock = threading.Lock()

REPO_ROOT = Path(__file__).resolve().parents[2]
ZHONGFANYI_ROOT = REPO_ROOT / "专检" / "中翻译"
_ALIASED_NAMESPACE_PACKAGES = (
    "llm",
    "llm.llm_project",
    "zhongfanyi",
    "zhongfanyi.llm",
    "zhongfanyi.llm.llm_project",
)
_LOCAL_PACKAGE_NAMES = (
    "parsers",
    "replace",
    "backup_copy",
    "llm_check",
    "note",
    "utils",
    "divide",
    "match",
)


def get_zhongfanyi_default_mode() -> str:
    return ZHONGFANYI_MODE_DOUBLE


def get_zhongfanyi_single_file_extensions() -> list[str]:
    return ZHONGFANYI_SINGLE_FILE_EXTENSIONS


def get_zhongfanyi_double_file_extensions() -> list[str]:
    return ZHONGFANYI_DOUBLE_FILE_EXTENSIONS


def _register_namespace_package(name: str, package_root: Path) -> None:
    module = types.ModuleType(name)
    module.__path__ = [str(package_root)]
    module.__package__ = name
    module.__spec__ = importlib.machinery.ModuleSpec(name, loader=None, is_package=True)
    module.__spec__.submodule_search_locations = [str(package_root)]
    sys.modules[name] = module


def _clear_stale_namespace_modules(name: str, expected_root: Path) -> None:
    stale_keys = []
    expected_root = expected_root.resolve()
    for key, module in list(sys.modules.items()):
        if key != name and not key.startswith(f"{name}."):
            continue
        module_root = None
        try:
            spec = getattr(module, "__spec__", None)
            if spec and spec.submodule_search_locations:
                module_root = Path(list(spec.submodule_search_locations)[0])
            elif hasattr(module, "__path__"):
                module_root = Path(list(module.__path__)[0])
            elif hasattr(module, "__file__") and module.__file__:
                module_root = Path(module.__file__).resolve().parent
        except Exception:
            module_root = None
        is_expected = False
        if module_root is not None:
            try:
                module_root.resolve().relative_to(expected_root)
                is_expected = True
            except Exception:
                is_expected = False
        if not is_expected:
            stale_keys.append(key)
    for key in stale_keys:
        del sys.modules[key]


def _prepare_zhongfanyi_import_path() -> None:
    if not ZHONGFANYI_ROOT.exists():
        raise FileNotFoundError(f"未找到新版中翻译目录: {ZHONGFANYI_ROOT}")

    for module_name in _LOCAL_PACKAGE_NAMES:
        _clear_stale_namespace_modules(module_name, ZHONGFANYI_ROOT)
    for alias in _ALIASED_NAMESPACE_PACKAGES:
        _clear_stale_namespace_modules(alias, ZHONGFANYI_ROOT)
        _register_namespace_package(alias, ZHONGFANYI_ROOT)

    root_str = str(ZHONGFANYI_ROOT)
    parent_str = str(ZHONGFANYI_ROOT.parent)
    repo_root_str = str(REPO_ROOT)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)
    if parent_str not in sys.path:
        sys.path.insert(0, parent_str)
    if repo_root_str not in sys.path:
        sys.path.insert(0, repo_root_str)


def _init_task(task_id: str) -> None:
    _task_progress[task_id] = {
        "status": "running",
        "progress": 0,
        "message": "初始化...",
        "details": [],
        "stream_log": "[init] 任务已创建",
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
    }


def _append_log(task_id: str, message: str) -> None:
    if task_id not in _task_progress:
        return
    text = (message or "").strip()
    if not text:
        return
    current = _task_progress[task_id].get("stream_log", "")
    _task_progress[task_id]["stream_log"] = f"{current}\n{text}" if current else text


def _update_task(task_id: str, message: str, progress: int = 0, details: Optional[list] = None) -> None:
    if task_id not in _task_progress:
        return
    _append_log(task_id, f"[{min(100, progress):>3}%] {message}")
    _task_progress[task_id].update(
        message=message,
        progress=min(100, progress),
        details=details or [],
        updated_at=datetime.now().isoformat(),
    )


def _complete_task(task_id: str, result: Optional[Dict[str, Any]] = None, error: Optional[str] = None) -> None:
    if task_id not in _task_progress:
        return
    _append_log(task_id, "[done] 处理完成" if not error else f"[error] {error}")
    _task_progress[task_id].update(
        status="done" if not error else "failed",
        progress=100,
        message="处理完成" if not error else f"失败: {error}",
        result=result,
        error=error,
        updated_at=datetime.now().isoformat(),
    )


def get_task_progress(task_id: str) -> Optional[Dict[str, Any]]:
    return _task_progress.get(task_id)


def _output_web_path(file_path: str | Path) -> str:
    resolved = Path(file_path).resolve()
    output_root = Path(settings.OUTPUT_DIR).resolve()
    relative = resolved.relative_to(output_root)
    normalized_relative = str(relative).replace("\\", "/")
    return f"outputs/{normalized_relative}"


def _normalize_report_count(path: Optional[str]) -> int:
    if not path:
        return 0
    try:
        from parsers.json.clean_json import extract_and_parse

        return len(extract_and_parse(path))
    except Exception:
        return 0


def _inject_runtime_env(gemini_route: str) -> None:
    if settings.GOOGLE_API_KEY:
        os.environ["GOOGLE_API_KEY"] = settings.GOOGLE_API_KEY
    if settings.OPENROUTER_API_KEY:
        os.environ["OPENROUTER_API_KEY"] = settings.OPENROUTER_API_KEY
        os.environ["OPENAI_API_KEY"] = settings.OPENROUTER_API_KEY
        os.environ["API_KEY"] = settings.OPENROUTER_API_KEY
    if settings.OPENROUTER_BASE_URL:
        os.environ["OPENROUTER_BASE_URL"] = settings.OPENROUTER_BASE_URL
        os.environ["OPENAI_BASE_URL"] = settings.OPENROUTER_BASE_URL
        os.environ["BASE_URL"] = settings.OPENROUTER_BASE_URL
    os.environ["GEMINI_ROUTE"] = gemini_route


def _collect_reports(report_paths: Optional[Dict[str, str]], excel_paths: Optional[Dict[str, str]]) -> tuple[Dict[str, str], Dict[str, int]]:
    reports: Dict[str, str] = {}
    report_counts: Dict[str, int] = {value: 0 for value in SECTION_COUNT_KEYS.values()}

    for section_name, path in (report_paths or {}).items():
        if not path or not Path(path).exists():
            continue
        reports[f"{section_name}_json"] = _output_web_path(path)
        report_counts[SECTION_COUNT_KEYS.get(section_name, section_name)] = _normalize_report_count(path)

    for section_name, path in (excel_paths or {}).items():
        if not path or not Path(path).exists():
            continue
        reports[f"{section_name}_excel"] = _output_web_path(path)

    return reports, report_counts


def _copy_final_output(result_path: Optional[str], output_dir: Path) -> tuple[Optional[str], Dict[str, str]]:
    if not result_path:
        return None, {}

    source = Path(result_path)
    if not source.exists():
        return None, {}

    suffix = source.suffix.lower()
    target_name = {
        ".docx": "corrected.docx",
        ".doc": "corrected.doc",
        ".pdf": "annotated.pdf",
        ".xlsx": "annotated.xlsx",
        ".xls": "annotated.xls",
        ".pptx": "annotated.pptx",
    }.get(suffix, source.name)
    target = output_dir / target_name

    try:
        same_file = source.resolve() == target.resolve()
    except Exception:
        same_file = str(source) == str(target)
    if not same_file:
        shutil.copy2(source, target)

    web_path = _output_web_path(target)
    payload = {"output_file": web_path}
    if suffix in {".docx", ".doc"}:
        payload["corrected_docx"] = web_path
    elif suffix == ".pdf":
        payload["annotated_pdf"] = web_path
    elif suffix in {".xlsx", ".xls"}:
        payload["annotated_excel"] = web_path
    elif suffix == ".pptx":
        payload["annotated_pptx"] = web_path
    return web_path, payload


def run_zhongfanyi_task(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    mode: str = ZHONGFANYI_MODE_DOUBLE,
    original_path: Optional[str] = None,
    translated_path: Optional[str] = None,
    single_path: Optional[str] = None,
    use_ai_rule: bool = False,
    gemini_route: str = ZHONGFANYI_DEFAULT_GEMINI_ROUTE,
    ai_rule_file_path: Optional[str] = None,
    session_rule_text: Optional[str] = None,
) -> Dict[str, Any]:
    """
    执行新版中翻译专检流程。
    """
    _init_task(task_id)
    normalized_mode = (mode or ZHONGFANYI_MODE_DOUBLE).strip().lower()
    if normalized_mode not in {ZHONGFANYI_MODE_DOUBLE, ZHONGFANYI_MODE_SINGLE}:
        _complete_task(task_id, error=f"不支持的中翻译模式: {mode}")
        return _task_progress[task_id]

    gemini_route = ensure_gemini_route_configured(ZHONGFANYI_DEFAULT_GEMINI_ROUTE)
    folder_name = display_no or task_id
    output_dir = Path(settings.OUTPUT_DIR) / "zhongfanyi" / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)

    with _specialist_import_lock:
        _prepare_zhongfanyi_import_path()
        _inject_runtime_env(gemini_route)
        from zhongfanyi.main import run_full_pipeline

    input_path = single_path if normalized_mode == ZHONGFANYI_MODE_SINGLE else original_path
    if not input_path:
        _complete_task(task_id, error="缺少输入文件")
        return _task_progress[task_id]

    _update_task(task_id, "正在提取文本并执行中翻译专检...", 20)
    result_path, report_paths, excel_paths, stats = run_full_pipeline(
        input_path,
        translated_path or input_path,
        str(output_dir),
        use_ai_rule=use_ai_rule,
        bilingual=normalized_mode == ZHONGFANYI_MODE_SINGLE,
        ai_rule_file_path=ai_rule_file_path or None,
        session_rule_text=session_rule_text,
    )

    if report_paths is None or excel_paths is None or stats is None:
        _complete_task(task_id, error="对比或规则加载失败")
        return _task_progress[task_id]

    _update_task(task_id, "正在整理输出文件...", 90)
    output_web_path, output_payload = _copy_final_output(result_path, output_dir)
    reports, report_counts = _collect_reports(report_paths, excel_paths)
    total_issues = sum(report_counts.values())

    result = {
        "task_id": task_id,
        "mode": normalized_mode,
        "gemini_route": gemini_route,
        "reports": reports,
        "report_counts": report_counts,
        "stats": stats or {"success": 0, "failed": 0, "skipped": 0},
        "summary": "单文件模式已完成中翻译专检。" if normalized_mode == ZHONGFANYI_MODE_SINGLE else "双文件模式已完成中翻译专检。",
        "total_issues": total_issues,
    }
    if normalized_mode == ZHONGFANYI_MODE_SINGLE and single_path:
        result["single_filename"] = Path(single_path).name
    if normalized_mode == ZHONGFANYI_MODE_DOUBLE:
        if original_path:
            result["original_filename"] = Path(original_path).name
        if translated_path:
            result["translated_filename"] = Path(translated_path).name
    if output_web_path:
        result.update(output_payload)

    _complete_task(task_id, result=result)
    return result
