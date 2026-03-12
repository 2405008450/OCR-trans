# -*- coding: utf-8 -*-
"""
中翻译专检：对接 专检/zhongfanyi/main 的 run_full_pipeline，支持原文/译文/可选规则文件上传。
"""
import os
import sys
import shutil
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from app.core.config import settings


_task_progress: Dict[str, Dict[str, Any]] = {}


def _zhongfanyi_root() -> Path:
    """专检 目录（含 zhongfanyi 包）"""
    repo_root = Path(__file__).resolve().parents[2]
    return repo_root / "专检"


def _zhongfanyi_cwd() -> Path:
    """专检/zhongfanyi 目录（含 llm 包，main.py 里 from llm.xxx 依赖此路径）"""
    return _zhongfanyi_root() / "zhongfanyi"


def _prepare_zhongfanyi_path() -> None:
    """把 专检 和 专检/zhongfanyi 加入 sys.path，使可导入 zhongfanyi 且 main 内可导入 llm"""
    root = _zhongfanyi_root()
    cwd = _zhongfanyi_cwd()
    for path in (cwd, root):  # 先加 zhongfanyi 目录，main 里 from llm.xxx 才能找到 llm
        path_str = str(path)
        if path.exists() and path_str not in sys.path:
            sys.path.insert(0, path_str)


def _init_task(task_id: str) -> None:
    _task_progress[task_id] = {
        "status": "running",
        "progress": 0,
        "message": "初始化...",
        "details": [],
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
    }


def _update_task(task_id: str, message: str, progress: int = 0, details: Optional[list] = None) -> None:
    if task_id not in _task_progress:
        return
    _task_progress[task_id].update(
        message=message,
        progress=min(100, progress),
        details=details or [],
        updated_at=datetime.now().isoformat(),
    )


def _complete_task(task_id: str, result: Optional[Dict[str, Any]] = None, error: Optional[str] = None) -> None:
    if task_id not in _task_progress:
        return
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
    use_ai_rule: bool = False,
    ai_rule_file_path: Optional[str] = None,
) -> Dict[str, Any]:
    """
    在指定输出目录下执行中翻译专检完整流程（对比 + 修复），结果复制到 output_dir 供下载。
    """
    output_dir = Path(settings.OUTPUT_DIR) / "zhongfanyi" / task_id
    output_dir.mkdir(parents=True, exist_ok=True)
    output_base = str(output_dir)

    _prepare_zhongfanyi_path()
    # 注入 API 配置到环境变量，供 zhongfanyi 子模块（rule_generation/check）读取，避免后台线程中 .env 未生效
    if settings.OPENROUTER_API_KEY:
        os.environ["OPENROUTER_API_KEY"] = settings.OPENROUTER_API_KEY
        os.environ["OPENAI_API_KEY"] = settings.OPENROUTER_API_KEY
    if settings.OPENROUTER_BASE_URL:
        os.environ["OPENROUTER_BASE_URL"] = settings.OPENROUTER_BASE_URL
    from zhongfanyi.main import run_full_pipeline

    _update_task(task_id, "正在提取与对比原文/译文...", 20)
    result_docx_path, report_paths, stats = run_full_pipeline(
        original_path,
        translated_path,
        output_base,
        use_ai_rule=use_ai_rule,
        ai_rule_file_path=ai_rule_file_path or None,
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
                reports[f"{name}_json"] = f"outputs/zhongfanyi/{task_id}/{subdir}/output_json/{fname}"

    # 前端下载使用相对站点的路径，如 /outputs/zhongfanyi/{task_id}/corrected.docx
    corrected_web_path = f"outputs/zhongfanyi/{task_id}/corrected.docx"
    result = {
        "task_id": task_id,
        "corrected_docx": corrected_web_path,
        "reports": reports,
        "stats": stats or {"success": 0, "failed": 0, "skipped": 0},
    }
    _complete_task(task_id, result=result)
    return result
