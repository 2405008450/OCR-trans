import os
import sys
import uuid
import shutil
from pathlib import Path
from typing import Dict, Any, List, Callable, Optional
from datetime import datetime

from fastapi import UploadFile
from docx import Document

from app.core.config import settings


# 任务进度存储
_task_progress: Dict[str, Dict[str, Any]] = {}


def _init_task_progress(task_id: str, total_steps: int = 6) -> None:
    """初始化任务进度"""
    _task_progress[task_id] = {
        "status": "running",
        "current_step": 0,
        "total_steps": total_steps,
        "message": "初始化任务...",
        "progress": 0,
        "details": [],
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
    """更新任务进度"""
    if task_id not in _task_progress:
        return

    progress = int((current_step / total_steps) * 100)
    _task_progress[task_id].update(
        {
            "current_step": current_step,
            "total_steps": total_steps,
            "message": message,
            "progress": progress,
            "details": details or [],
            "updated_at": datetime.now().isoformat(),
        }
    )


def _complete_task(task_id: str, result: Optional[Dict[str, Any]] = None, error: Optional[str] = None) -> None:
    """完成任务"""
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


def _get_task_progress(task_id: str) -> Optional[Dict[str, Any]]:
    """获取任务进度"""
    return _task_progress.get(task_id)


def _cleanup_task(task_id: str) -> None:
    """清理任务记录（可选，避免内存无限增长）"""
    if task_id in _task_progress:
        del _task_progress[task_id]


def _normalize_path(path: str) -> str:
    return path.replace("\\", "/") if path else path


def _validate_docx(file: UploadFile, label: str) -> None:
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext != ".docx":
        raise ValueError(f"{label} 必须是 .docx 文件，当前: {file.filename}")


def _prepare_specialized_import_path() -> None:
    """
    将 `专检/数值检查` 加入 sys.path，便于导入 llm.* 模块。
    """
    repo_root = Path(__file__).resolve().parents[2]
    specialized_root = repo_root / "专检" / "数值检查"
    specialized_root_str = str(specialized_root)
    if specialized_root.exists() and specialized_root_str not in sys.path:
        sys.path.insert(0, specialized_root_str)


def _apply_all_fixes(
    doc: Document,
    errors: List[Dict[str, Any]],
    label: str,
    replace_and_comment_in_docx,
    comment_manager,
    progress_callback: Optional[Callable[[int, int, str, List[str]], None]] = None,
) -> Dict[str, int]:
    if not errors:
        return {"success": 0, "failed": 0, "skipped": 0}

    success = 0
    failed = 0
    skipped = 0
    total = len(errors)
    print(f"\n>>> 正在修复 {label}，共 {total} 条")

    for idx, e in enumerate(errors, 1):
        old = (e.get("译文数值") or "").strip()
        new = (e.get("译文修改建议值") or "").strip()
        reason = (e.get("修改理由") or "").strip()
        context = e.get("译文上下文", "") or ""
        anchor = e.get("替换锚点", "") or ""

        if not old or not new:
            skipped += 1
            print(f"  [{idx}] 跳过: 缺少 old/new")
        else:
            ok, strategy = replace_and_comment_in_docx(
                doc,
                old,
                new,
                reason,
                comment_manager,
                context=context,
                anchor_text=anchor,
            )
            if ok:
                success += 1
                print(f"  [{idx}] 成功: '{old}' -> '{new}' ({strategy})")
            else:
                failed += 1
                print(f"  [{idx}] 失败: 未匹配 '{old}'")

        # 每10%进度或最后一项时报告进度
        if progress_callback and (idx % max(1, total // 10) == 0 or idx == total):
            details = [
                f"{label}: 成功 {success}, 失败 {failed}, 跳过 {skipped}",
                f"进度: {idx}/{total}"
            ]
            progress_callback(idx, total, f"正在修复{label} ({idx}/{total})", details)

    return {"success": success, "failed": failed, "skipped": skipped}


async def run_number_check_task(
    original_file: UploadFile,
    translated_file: UploadFile,
) -> Dict[str, Any]:
    """
    数值专检流程：
    1) 上传原文/译文 docx
    2) 提取正文/页眉/页脚并调用 Match 对比
    3) 生成 JSON 报告
    4) 对译文副本执行替换 + 批注
    5) 返回可下载结果
    """
    _validate_docx(original_file, "原文")
    _validate_docx(translated_file, "译文")

    task_id = str(uuid.uuid4())

    # 初始化进度跟踪（7个主要步骤）
    _init_task_progress(task_id, total_steps=7)

    upload_dir = Path(settings.UPLOAD_DIR) / "number_check"
    output_dir = Path(settings.OUTPUT_DIR) / "number_check" / task_id
    report_dir = output_dir / "reports"
    upload_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)

    original_path = upload_dir / f"{task_id}_original.docx"
    translated_path = upload_dir / f"{task_id}_translated.docx"

    _update_progress(task_id, 1, 7, "正在保存上传文件...")

    with open(original_path, "wb") as f:
        f.write(await original_file.read())
    with open(translated_path, "wb") as f:
        f.write(await translated_file.read())

    _prepare_specialized_import_path()

    _update_progress(task_id, 2, 7, "正在加载处理模块...")

    # 延迟导入，确保 sys.path 生效
    from llm.llm_project.llm_check.check import Match
    from llm.llm_project.parsers.body_extractor import extract_body_text
    from llm.llm_project.parsers.footer_extractor import extract_footers
    from llm.llm_project.parsers.header_extractor import extract_headers
    from llm.llm_project.replace.fix_replace_docx import ensure_backup_copy
    from llm.llm_project.replace.fix_replace_json import (
        replace_and_comment_in_docx,
        CommentManager,
    )
    from llm.utils.clean_json import load_json_file
    from llm.utils.json_files import write_json_with_timestamp

    # 阶段1：提取 & 对比
    _update_progress(task_id, 3, 7, "正在提取文档文本...")
    original_body = extract_body_text(str(original_path))
    translated_body = extract_body_text(str(translated_path))
    original_header = extract_headers(str(original_path))
    translated_header = extract_headers(str(translated_path))
    original_footer = extract_footers(str(original_path))
    translated_footer = extract_footers(str(translated_path))

    _update_progress(task_id, 4, 7, "正在对比数值差异...")

    matcher = Match()
    parts = [
        ("正文", original_body, translated_body, report_dir / "body"),
        ("页眉", original_header, translated_header, report_dir / "header"),
        ("页脚", original_footer, translated_footer, report_dir / "footer"),
    ]

    report_paths: Dict[str, str] = {}
    for idx, (name, orig_txt, tran_txt, out_dir) in enumerate(parts, 1):
        out_dir.mkdir(parents=True, exist_ok=True)
        _update_progress(task_id, 4, 7, f"正在对比{name} ({idx}/3)...")
        if orig_txt and tran_txt:
            result = matcher.compare_texts(orig_txt, tran_txt)
        else:
            result = []
        _, json_path = write_json_with_timestamp(result, str(out_dir))
        report_paths[name] = json_path

    # 阶段2：加载报告并修复
    _update_progress(task_id, 5, 7, "正在加载检查报告...")
    body_errors = load_json_file(report_paths.get("正文"))
    header_errors = load_json_file(report_paths.get("页眉"))
    footer_errors = load_json_file(report_paths.get("页脚"))

    backup_copy_path = ensure_backup_copy(str(translated_path))
    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    # 创建进度回调函数
    def progress_callback(current: int, total: int, msg: str, details: List[str]) -> None:
        _update_progress(task_id, 5, 7, msg, details)

    body_stat = _apply_all_fixes(doc, body_errors, "正文", replace_and_comment_in_docx, comment_manager, progress_callback)
    header_stat = _apply_all_fixes(doc, header_errors, "页眉", replace_and_comment_in_docx, comment_manager, progress_callback)
    footer_stat = _apply_all_fixes(doc, footer_errors, "页脚", replace_and_comment_in_docx, comment_manager, progress_callback)

    _update_progress(task_id, 6, 7, "正在保存修复后的文档...")

    doc.save(backup_copy_path)

    final_doc_path = output_dir / f"{task_id}_corrected.docx"
    shutil.copy2(backup_copy_path, final_doc_path)

    total_success = body_stat["success"] + header_stat["success"] + footer_stat["success"]
    total_failed = body_stat["failed"] + header_stat["failed"] + footer_stat["failed"]
    total_skipped = body_stat["skipped"] + header_stat["skipped"] + footer_stat["skipped"]

    result = {
        "task_id": task_id,
        "original_filename": original_file.filename,
        "translated_filename": translated_file.filename,
        "corrected_docx": _normalize_path(str(final_doc_path)),
        "reports": {
            "body_json": _normalize_path(report_paths.get("正文")),
            "header_json": _normalize_path(report_paths.get("页眉")),
            "footer_json": _normalize_path(report_paths.get("页脚")),
        },
        "stats": {
            "success": total_success,
            "failed": total_failed,
            "skipped": total_skipped,
        },
    }

    _complete_task(task_id, result=result)
    return result
