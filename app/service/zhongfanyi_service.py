# -*- coding: utf-8 -*-
"""
中翻译专检：对接 专检/zhongfanyi/main 的 run_full_pipeline，支持原文/译文/可选规则文件上传。
"""
import os
import sys
import shutil
import threading
import types
import importlib.machinery
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from app.core.config import settings
from app.service.gemini_service import ensure_gemini_route_configured


_task_progress: Dict[str, Dict[str, Any]] = {}
_specialist_import_lock = threading.Lock()

REPO_ROOT = Path(__file__).resolve().parents[2]
ZHONGFANYI_PARENT_ROOT = REPO_ROOT / "专检"
ZHONGFANYI_ROOT = ZHONGFANYI_PARENT_ROOT / "zhongfanyi"
ZHONGFANYI_LLM_ROOT = ZHONGFANYI_ROOT / "llm"


def _zhongfanyi_root() -> Path:
    """专检 目录（含 zhongfanyi 包）"""
    return ZHONGFANYI_PARENT_ROOT


def _zhongfanyi_cwd() -> Path:
    """专检/zhongfanyi 目录（含 llm 包，main.py 里 from llm.xxx 依赖此路径）"""
    return ZHONGFANYI_ROOT


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
        raise FileNotFoundError(f"未找到中翻译专检目录: {ZHONGFANYI_ROOT}")
    if not ZHONGFANYI_LLM_ROOT.exists():
        raise FileNotFoundError(f"未找到中翻译专检 llm 目录: {ZHONGFANYI_LLM_ROOT}")

    _clear_stale_namespace_modules("llm", ZHONGFANYI_ROOT)
    _clear_stale_namespace_modules("zhongfanyi", ZHONGFANYI_ROOT)
    _register_namespace_package("llm", ZHONGFANYI_LLM_ROOT)
    _register_namespace_package("zhongfanyi", ZHONGFANYI_ROOT)

    zhongfanyi_root_str = str(ZHONGFANYI_ROOT)
    parent_root_str = str(ZHONGFANYI_PARENT_ROOT)
    if zhongfanyi_root_str not in sys.path:
        sys.path.insert(0, zhongfanyi_root_str)
    if parent_root_str not in sys.path:
        sys.path.insert(0, parent_root_str)


def _prepare_zhongfanyi_path() -> None:
    """把 专检 和 专检/zhongfanyi 加入 sys.path，使可导入 zhongfanyi 且 main 内可导入 llm"""
    _prepare_zhongfanyi_import_path()


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


def _normalize_path(p: str) -> str:
    return p.replace("\\", "/") if p else p


def run_zhongfanyi_task(
    original_path: str,
    translated_path: str,
    task_id: str,
    display_no: Optional[str] = None,
    use_ai_rule: bool = False,
    gemini_route: str = "openrouter",
    ai_rule_file_path: Optional[str] = None,
    session_rule_text: Optional[str] = None,
) -> Dict[str, Any]:
    """
    在指定输出目录下执行中翻译专检完整流程（对比 + 修复），结果复制到 output_dir 供下载。
    """
    folder_name = display_no or task_id
    gemini_route = ensure_gemini_route_configured(gemini_route)
    output_dir = Path(settings.OUTPUT_DIR) / "zhongfanyi" / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    output_base = str(output_dir)

    with _specialist_import_lock:
        _prepare_zhongfanyi_path()
    # 注入 API 配置到环境变量，供 zhongfanyi 子模块（rule_generation/check）读取，避免后台线程中 .env 未生效
    if settings.GOOGLE_API_KEY:
        os.environ["GOOGLE_API_KEY"] = settings.GOOGLE_API_KEY
    if settings.OPENROUTER_API_KEY:
        os.environ["OPENROUTER_API_KEY"] = settings.OPENROUTER_API_KEY
        os.environ["OPENAI_API_KEY"] = settings.OPENROUTER_API_KEY
    if settings.OPENROUTER_BASE_URL:
        os.environ["OPENROUTER_BASE_URL"] = settings.OPENROUTER_BASE_URL
    os.environ["GEMINI_ROUTE"] = gemini_route
    with _specialist_import_lock:
        from zhongfanyi.main import run_full_pipeline

    _update_task(task_id, "正在提取与对比原文/译文...", 20)
    result_docx_path, report_paths, stats = run_full_pipeline(
        original_path,
        translated_path,
        output_base,
        use_ai_rule=use_ai_rule,
        ai_rule_file_path=ai_rule_file_path or None,
        session_rule_text=session_rule_text,
    )

    if result_docx_path is None:
        _complete_task(task_id, error="对比或规则加载失败")
        return _task_progress[task_id]

    _update_task(task_id, "正在保存结果...", 90)
    # 复制修复后的 docx 到统一输出目录，便于前端下载
    final_docx = output_dir / "corrected.docx"
    shutil.copy2(result_docx_path, final_docx)

    # 报告 JSON 在 output_dir 下的 zhengwen/yemei/yejiao/output_json 中，转为前端可下载的相对路径
    report_dir_names = {"正文": "zhengwen", "页眉": "yemei", "页脚": "yejiao"}
    reports = {}
    if report_paths:
        for name, path in report_paths.items():
            if path and os.path.exists(path):
                subdir = report_dir_names.get(name, name)
                fname = os.path.basename(path)
                reports[f"{name}_json"] = f"outputs/zhongfanyi/{folder_name}/{subdir}/output_json/{fname}"

    # 前端下载使用相对站点的路径，如 /outputs/zhongfanyi/{task_id}/corrected.docx
    corrected_web_path = f"outputs/zhongfanyi/{folder_name}/corrected.docx"
    result = {
        "task_id": task_id,
        "gemini_route": gemini_route,
        "corrected_docx": corrected_web_path,
        "reports": reports,
        "stats": stats or {"success": 0, "failed": 0, "skipped": 0},
    }
    _complete_task(task_id, result=result)
    return result
