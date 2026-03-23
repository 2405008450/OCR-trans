import asyncio
import logging
import os
import shutil
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from docx import Document
from fastapi import UploadFile

from app.core.config import settings
from app.service.gemini_service import ensure_gemini_route_configured


_task_progress: Dict[str, Dict[str, Any]] = {}
logger = logging.getLogger("app.number_check")

# 专检模块导入锁：防止并发任务修改 sys.path / sys.modules 时互相干扰
_specialist_import_lock = threading.Lock()

NUMBER_CHECK_MODELS: Dict[str, Dict[str, str]] = {
    "gemini-3-flash-preview": {
        "label": "快速版V2",
        "description": "速度更快，适合常规数字核对场景。",
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
    "gemini-3.1-pro-preview": "gemini-3.1-pro-preview",
    "google/gemini-3.1-pro-preview": "gemini-3.1-pro-preview",
    "gemini 3.1 pro preview": "gemini-3.1-pro-preview",
}


def get_number_check_models() -> Dict[str, Dict[str, str]]:
    return NUMBER_CHECK_MODELS


def normalize_number_check_model(model_name: Optional[str]) -> str:
    candidate = (model_name or "gemini-3-flash-preview").strip()
    key = NUMBER_CHECK_MODEL_ALIASES.get(candidate.lower(), candidate)
    if key not in NUMBER_CHECK_MODELS:
        raise ValueError(f"不支持的数字专检模型: {model_name}")
    return key


def _append_stream_log(task_id: str, message: str) -> None:
    if task_id not in _task_progress:
        return
    line = (message or "").strip()
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
    if task_id in _task_progress:
        del _task_progress[task_id]


def _normalize_path(path: str) -> str:
    return path.replace("\\", "/") if path else path


def _validate_docx(file: UploadFile, label: str) -> None:
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext != ".docx":
        raise ValueError(f"{label} 必须是 .docx 文件，当前为 {file.filename}")


def _prepare_specialized_import_path() -> None:
    """
    将 专检/数值检查/ 目录加入 sys.path，并清理可能由其他专检模块（如 zhongfanyi）
    遗留的同名 'llm.*' 缓存，避免 sys.modules 命名空间污染。
    """
    repo_root = Path(__file__).resolve().parents[2]
    # 硬编码精确路径，避免 rglob 误匹配 zhongfanyi 同名 llm_project
    specialized_root = repo_root / "专检" / "数值检查"
    if not specialized_root.exists():
        # 兼容：如果目录名编码不同（Windows 中文路径），也用 rglob 兜底
        for candidate in repo_root.glob("专检/*/llm/llm_project"):
            check_file = candidate / "llm_check" / "check.py"
            if check_file.exists():
                # 只接受数值检查目录（check.py 导入 gemini_service 而不是 zhongfanyi）
                content = check_file.read_text(encoding="utf-8", errors="ignore")
                if "zhongfanyi" not in content:
                    specialized_root = candidate.parent.parent
                    break
    if not specialized_root.exists():
        raise FileNotFoundError(f"未找到数字专检依赖目录: {specialized_root}")

    specialized_root_str = str(specialized_root)

    # 检查 sys.modules 中的 'llm' 是否来自正确目录，否则清除污染缓存
    llm_mod = sys.modules.get("llm")
    if llm_mod is not None:
        llm_path_obj: Optional[Path] = None
        try:
            spec = getattr(llm_mod, "__spec__", None)
            if spec and spec.submodule_search_locations:
                llm_path_obj = Path(list(spec.submodule_search_locations)[0])
            elif hasattr(llm_mod, "__path__"):
                llm_path_obj = Path(list(llm_mod.__path__)[0])
            elif hasattr(llm_mod, "__file__") and llm_mod.__file__:
                llm_path_obj = Path(llm_mod.__file__).parent
        except Exception:
            pass
        # 使用 Path 比对（大小写不敏感），判断缓存是否来自正确目录
        is_correct = False
        if llm_path_obj is not None:
            try:
                # 如果 llm_path_obj 是 specialized_root 的子路径则正确
                llm_path_obj.relative_to(specialized_root)
                is_correct = True
            except ValueError:
                is_correct = False
        if not is_correct:
            stale = [k for k in list(sys.modules.keys()) if k == "llm" or k.startswith("llm.")]
            for k in stale:
                del sys.modules[k]
            if stale:
                logger.info(f"[import] 清除了 {len(stale)} 个来自其他目录的 llm.* 污染缓存")

    if specialized_root_str not in sys.path:
        sys.path.insert(0, specialized_root_str)


def _apply_all_fixes(
    doc: Document,
    errors: List[Dict[str, Any]],
    label: str,
    replace_and_comment_in_docx,
    comment_manager,
    task_id: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str, List[str]], None]] = None,
) -> Dict[str, int]:
    if not errors:
        if task_id:
            _emit_log(task_id, f"[fix] {label} 无需修复")
        return {"success": 0, "failed": 0, "skipped": 0}

    success = 0
    failed = 0
    skipped = 0
    total = len(errors)
    if task_id:
        _emit_log(task_id, f"[fix] >>> 正在修复 {label}（共 {total} 条）...")

    for idx, error in enumerate(errors, 1):
        old = (error.get("译文数值") or "").strip()
        new = (error.get("译文修改建议值") or "").strip()
        reason = (error.get("修改理由") or "").strip()
        context = error.get("译文上下文", "") or ""
        anchor = error.get("替换锚点", "") or ""

        if not old or not new:
            skipped += 1
            if task_id:
                _emit_log(task_id, f"[fix]   [{idx}/{total}] 跳过: 缺少【译文数值】或【译文修改建议值】", level="warning")
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
                if task_id:
                    _emit_log(task_id, f"[fix]   [{idx}/{total}] ✓ '{old}' → '{new}'  策略={strategy}  理由={reason}")
            else:
                failed += 1
                if task_id:
                    _emit_log(task_id, f"[fix]   [{idx}/{total}] ✗ 未匹配到 '{old}'", level="warning")

        if progress_callback and (idx % max(1, total // 10) == 0 or idx == total):
            details = [
                f"{label}: 成功 {success}，失败 {failed}，跳过 {skipped}",
                f"进度: {idx}/{total}",
            ]
            progress_callback(idx, total, f"正在修复{label} ({idx}/{total})", details)

    rate = success / (success + failed) if (success + failed) else 0
    if task_id:
        _emit_log(
            task_id,
            f"[fix] --- {label} 修复统计: 成功={success} 失败={failed} 跳过={skipped} 成功率={rate:.0%} ---",
        )
    return {"success": success, "failed": failed, "skipped": skipped}


def _run_number_check_sync(
    task_id: str,
    original_path: Path,
    translated_path: Path,
    output_dir: Path,
    report_dir: Path,
    gemini_route: str,
    model_name: str,
    original_filename: str,
    translated_filename: str,
) -> Dict[str, Any]:
    """
    所有同步阻塞操作（模块导入、文本提取、LLM调用、文档修复）都在此函数中执行。
    由调用方通过 asyncio.to_thread 放进线程池，避免阻塞事件循环。
    """
    import time as _time

    # ── 路径准备 & 模块导入（用锁防止并发任务污染 sys.modules）─────────
    with _specialist_import_lock:
        _prepare_specialized_import_path()
        if settings.GOOGLE_API_KEY:
            os.environ["GOOGLE_API_KEY"] = settings.GOOGLE_API_KEY
        if settings.OPENROUTER_API_KEY:
            os.environ["OPENROUTER_API_KEY"] = settings.OPENROUTER_API_KEY
            os.environ["OPENAI_API_KEY"] = settings.OPENROUTER_API_KEY
        if settings.OPENROUTER_BASE_URL:
            os.environ["OPENROUTER_BASE_URL"] = settings.OPENROUTER_BASE_URL
        os.environ["GEMINI_ROUTE"] = gemini_route

        _update_progress(task_id, 2, 7, "正在加载处理模块...")
        from llm.llm_project.llm_check.check import Match
        from llm.llm_project.parsers.body_extractor import extract_body_text
        from llm.llm_project.parsers.footer_extractor import extract_footers
        from llm.llm_project.parsers.header_extractor import extract_headers
        from llm.llm_project.replace.fix_replace_docx import ensure_backup_copy
        from llm.llm_project.replace.fix_replace_json import CommentManager, replace_and_comment_in_docx
        from llm.utils.clean_json import load_json_file
        from llm.utils.json_files import write_json_with_timestamp

    # ── 阶段 1：逐步提取文本 ─────────────────────────────────────
    _update_progress(task_id, 3, 7, "正在提取文档文本...")

    _emit_log(task_id, "[extract] 正在提取原文正文...")
    original_body = extract_body_text(str(original_path))
    _emit_log(task_id, f"[extract]   原文正文长度: {len(original_body)} 字符")

    _emit_log(task_id, "[extract] 正在提取译文正文...")
    translated_body = extract_body_text(str(translated_path))
    _emit_log(task_id, f"[extract]   译文正文长度: {len(translated_body)} 字符")

    _emit_log(task_id, "[extract] 正在提取页眉...")
    original_header = extract_headers(str(original_path))
    translated_header = extract_headers(str(translated_path))
    _emit_log(task_id, f"[extract]   原文页眉: {len(original_header)} 字符 / 译文页眉: {len(translated_header)} 字符")

    _emit_log(task_id, "[extract] 正在提取页脚...")
    original_footer = extract_footers(str(original_path))
    translated_footer = extract_footers(str(translated_path))
    _emit_log(task_id, f"[extract]   原文页脚: {len(original_footer)} 字符 / 译文页脚: {len(translated_footer)} 字符")

    def _preview(text: str, limit: int = 200) -> str:
        s = str(text)[:limit]
        return s + ("..." if len(str(text)) > limit else "")

    _emit_log(task_id, "[preview] ==== 原文内容预览 ====")
    _emit_log(task_id, f"[preview] [页眉] {_preview(original_header) or '(空)'}")
    _emit_log(task_id, f"[preview] [正文] {_preview(original_body, 300) or '(空)'}")
    _emit_log(task_id, f"[preview] [页脚] {_preview(original_footer) or '(空)'}")
    _emit_log(task_id, "[preview] ==== 译文内容预览 ====")
    _emit_log(task_id, f"[preview] [页眉] {_preview(translated_header) or '(空)'}")
    _emit_log(task_id, f"[preview] [正文] {_preview(translated_body, 300) or '(空)'}")
    _emit_log(task_id, f"[preview] [页脚] {_preview(translated_footer) or '(空)'}")

    # ── 阶段 2：LLM 对比（同步调用，因为本函数已在线程中）────────
    _update_progress(task_id, 4, 7, "正在对比数值差异...")
    _emit_log(task_id, f"[compare] === LLM 对比  路线={gemini_route}  模型={model_name} ===")

    matcher = Match(model_name=model_name, task_logger=lambda message: _emit_log(task_id, message))
    parts = [
        ("正文", original_body, translated_body, report_dir / "body"),
        ("页眉", original_header, translated_header, report_dir / "header"),
        ("页脚", original_footer, translated_footer, report_dir / "footer"),
    ]

    report_paths: Dict[str, str] = {}
    for idx, (name, original_text, translated_text, output_subdir) in enumerate(parts, 1):
        output_subdir.mkdir(parents=True, exist_ok=True)
        _update_progress(task_id, 4, 7, f"正在对比{name} ({idx}/3)...")
        _emit_log(
            task_id,
            f"[compare] ====== 正在检查{name} "
            f"(原文 {len(original_text)} 字符 / 译文 {len(translated_text)} 字符) ======",
        )
        if original_text and translated_text:
            t0 = _time.time()
            result = matcher.compare_texts(original_text, translated_text)
            elapsed = _time.time() - t0
            error_count = len(result) if isinstance(result, list) else "?"
            _emit_log(task_id, f"[compare] {name}检查完成，耗时 {elapsed:.1f}s，发现 {error_count} 条问题")
        else:
            result = []
            _emit_log(task_id, f"[skip] {name} 文本为空，跳过模型比对", level="warning")
        _, json_path = write_json_with_timestamp(result, str(output_subdir))
        report_paths[name] = json_path
        _emit_log(task_id, f"[report] {name} 报告已写入: {json_path}")

    # ── 阶段 3：加载报告 ─────────────────────────────────────────
    _update_progress(task_id, 5, 7, "正在加载检查报告...")
    body_errors = load_json_file(report_paths.get("正文"))
    header_errors = load_json_file(report_paths.get("页眉"))
    footer_errors = load_json_file(report_paths.get("页脚"))

    def _log_errors_preview(label: str, errors: list) -> None:
        _emit_log(task_id, f"[report] 已加载 {label} 报告: {len(errors)} 条错误")
        for i, item in enumerate(errors[:3], 1):
            _emit_log(
                task_id,
                f"[report]   [{i}] 类型={item.get('错误类型', '')}  "
                f"译文值={item.get('译文数值', '')}  "
                f"建议={item.get('译文修改建议值', '')}",
            )
        if len(errors) > 3:
            _emit_log(task_id, f"[report]   ... 共 {len(errors)} 条（仅预览前 3 条）")

    _log_errors_preview("正文", body_errors)
    _log_errors_preview("页眉", header_errors)
    _log_errors_preview("页脚", footer_errors)

    # ── 阶段 4：自动替换与批注 ────────────────────────────────────
    _emit_log(task_id, "[fix] === 阶段 3：自动替换与批注 ===")
    _emit_log(task_id, "[fix] 正在创建译文备份...")
    backup_copy_path = ensure_backup_copy(str(translated_path))
    _emit_log(task_id, f"[fix] 备份路径: {backup_copy_path}")
    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    def progress_callback(current: int, total: int, message: str, details: List[str]) -> None:
        _update_progress(task_id, 5, 7, message, details)

    body_stat = _apply_all_fixes(doc, body_errors, "正文", replace_and_comment_in_docx, comment_manager,
                                  task_id=task_id, progress_callback=progress_callback)
    header_stat = _apply_all_fixes(doc, header_errors, "页眉", replace_and_comment_in_docx, comment_manager,
                                    task_id=task_id, progress_callback=progress_callback)
    footer_stat = _apply_all_fixes(doc, footer_errors, "页脚", replace_and_comment_in_docx, comment_manager,
                                    task_id=task_id, progress_callback=progress_callback)

    _update_progress(task_id, 6, 7, "正在保存修复后的文档...")
    doc.save(backup_copy_path)

    final_doc_path = output_dir / "corrected.docx"
    shutil.copy2(backup_copy_path, final_doc_path)
    _emit_log(task_id, f"[output] 修复文档已写出: {final_doc_path}")

    total_success = body_stat["success"] + header_stat["success"] + footer_stat["success"]
    total_failed = body_stat["failed"] + header_stat["failed"] + footer_stat["failed"]
    total_skipped = body_stat["skipped"] + header_stat["skipped"] + footer_stat["skipped"]

    result = {
        "task_id": task_id,
        "original_filename": original_filename,
        "translated_filename": translated_filename,
        "model_name": model_name,
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

    overall_rate = total_success / (total_success + total_failed) if (total_success + total_failed) else 0
    _emit_log(task_id, "=" * 50)
    _emit_log(task_id, "[summary] 全部流程处理完成！")
    _emit_log(
        task_id,
        f"[summary] 成功={total_success}  失败={total_failed}  跳过={total_skipped}  "
        f"总计={total_success + total_failed + total_skipped}  成功率={overall_rate:.0%}",
    )
    _emit_log(task_id, f"[summary] 最终文件: {final_doc_path}")
    _emit_log(task_id, "=" * 50)
    _complete_task(task_id, result=result)
    return result


async def run_number_check_task(
    original_file: UploadFile,
    translated_file: UploadFile,
    task_id: str = "",
    display_no: Optional[str] = None,
    gemini_route: str = "google",
    model_name: str = "gemini-3-flash-preview",
) -> Dict[str, Any]:
    _validate_docx(original_file, "原文")
    _validate_docx(translated_file, "译文")
    gemini_route = ensure_gemini_route_configured(gemini_route)
    model_name = normalize_number_check_model(model_name)

    if not task_id:
        task_id = str(uuid.uuid4())

    _init_task_progress(task_id, total_steps=7)
    _emit_log(task_id, f"[config] route={gemini_route}, model={model_name}")

    folder_name = display_no or task_id
    upload_dir = Path(settings.UPLOAD_DIR) / "number_check" / folder_name
    output_dir = Path(settings.OUTPUT_DIR) / "number_check" / folder_name
    report_dir = output_dir / "reports"
    upload_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)

    original_path = upload_dir / "original.docx"
    translated_path = upload_dir / "translated.docx"

    # 保存上传文件（需要 await，必须在 async 中）
    _update_progress(task_id, 1, 7, "正在保存上传文件...")
    with open(original_path, "wb") as f:
        f.write(await original_file.read())
    with open(translated_path, "wb") as f:
        f.write(await translated_file.read())

    # 将所有阻塞操作放入线程池，彻底释放事件循环
    try:
        return await asyncio.to_thread(
            _run_number_check_sync,
            task_id,
            original_path,
            translated_path,
            output_dir,
            report_dir,
            gemini_route,
            model_name,
            original_file.filename or "original.docx",
            translated_file.filename or "translated.docx",
        )
    except Exception as exc:
        # 保证内存中的进度状态也更新为 failed，防止前端无限轮询
        _emit_log(task_id, f"[error] 任务失败: {type(exc).__name__}: {exc}", level="error")
        _complete_task(task_id, error=str(exc))
        raise
