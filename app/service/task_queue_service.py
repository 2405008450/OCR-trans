import asyncio
import io
import json
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from fastapi import UploadFile

from app.core.config import settings
from app.core.file_naming import build_storage_filename
from app.db.session import SessionLocal
from app.repository import task_repo
from app.service.doc_translate_service import execute_doc_translate_task
from app.service import zhongfanyi_service as zf_service
from app.service.llm_service import execute_ocr_task_from_path
from app.service.number_check_service import _get_task_progress as get_number_check_progress
from app.service.number_check_service import run_number_check_task
from app.service.pdf2docx_service import execute_pdf2docx_task_from_path


class TaskQueueService:
    def __init__(self):
        self._worker_task: Optional[asyncio.Task] = None
        self._stop_event = asyncio.Event()
        self._task_executor: Optional[ThreadPoolExecutor] = None
        self._task_logs: Dict[str, str] = {}
        self._last_log_line: Dict[str, str] = {}
        self._max_log_chars = 50000

    async def start(self):
        if self._worker_task and not self._worker_task.done():
            return
        self._stop_event = asyncio.Event()
        if self._task_executor is None:
            self._task_executor = ThreadPoolExecutor(max_workers=4, thread_name_prefix="task-queue")
        self._requeue_interrupted_tasks()
        self._worker_task = asyncio.create_task(self._worker_loop(), name="task-queue-worker")

    async def stop(self):
        if not self._worker_task:
            return
        self._stop_event.set()
        self._worker_task.cancel()
        try:
            await self._worker_task
        except asyncio.CancelledError:
            pass
        finally:
            self._worker_task = None
            if self._task_executor is not None:
                self._task_executor.shutdown(wait=False, cancel_futures=False)
                self._task_executor = None

    async def submit_ocr_task(self, **kwargs) -> str:
        file: UploadFile = kwargs.pop("file")
        reserved_task = self._create_db_task(
            task_type="ocr",
            filename=file.filename or "input.bin",
            params=kwargs,
            input_files={},
        )
        try:
            input_path, original_filename = await self._save_single_upload(
                file,
                "ocr",
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {"input_path": input_path, "original_filename": original_filename},
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_number_check_task(
        self,
        *,
        original_file: UploadFile,
        translated_file: UploadFile,
        gemini_route: str,
        model_name: str,
    ) -> str:
        reserved_task = self._create_db_task(
            task_type="number_check",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={"gemini_route": gemini_route, "model_name": model_name},
            input_files={},
        )
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / "number_check" / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)

            original_ext = Path(original_file.filename or "original.docx").suffix or ".docx"
            translated_ext = Path(translated_file.filename or "translated.docx").suffix or ".docx"
            original_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                original_file.filename,
                reserved_task.task_id,
                role="original",
                ext=original_ext,
            )
            translated_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                translated_file.filename,
                reserved_task.task_id,
                role="translated",
                ext=translated_ext,
            )
            original_path.write_bytes(await original_file.read())
            translated_path.write_bytes(await translated_file.read())

            self._update_task_input_files(
                reserved_task.task_id,
                {
                    "original_path": str(original_path).replace("\\", "/"),
                    "translated_path": str(translated_path).replace("\\", "/"),
                    "original_filename": original_file.filename,
                    "translated_filename": translated_file.filename,
                },
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_zhongfanyi_task(
        self,
        *,
        original_file: UploadFile,
        translated_file: UploadFile,
        use_ai_rule: bool,
        gemini_route: str,
        rule_file: Optional[UploadFile],
        session_rule_content: Optional[str],
    ) -> str:
        reserved_task = self._create_db_task(
            task_type="zhongfanyi",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={
                "use_ai_rule": use_ai_rule,
                "gemini_route": gemini_route,
                "ai_rule_file_path": None,
                "session_rule_text": (session_rule_content.strip() or None) if session_rule_content else None,
            },
            input_files={},
        )
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / "zhongfanyi" / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)

            ext_orig = Path(original_file.filename or "original.docx").suffix.lower()
            ext_tran = Path(translated_file.filename or "translated.docx").suffix.lower()
            original_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                original_file.filename,
                reserved_task.task_id,
                role="original",
                ext=ext_orig,
            )
            translated_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                translated_file.filename,
                reserved_task.task_id,
                role="translated",
                ext=ext_tran,
            )
            original_path.write_bytes(await original_file.read())
            translated_path.write_bytes(await translated_file.read())

            ai_rule_file_path = None
            if rule_file and use_ai_rule:
                ext_rule = Path(rule_file.filename or "rule.txt").suffix.lower()
                ai_rule_path = upload_dir / build_storage_filename(
                    reserved_task.display_no,
                    rule_file.filename,
                    reserved_task.task_id,
                    role="rule",
                    ext=ext_rule,
                )
                ai_rule_path.write_bytes(await rule_file.read())
                ai_rule_file_path = str(ai_rule_path).replace("\\", "/")

            with SessionLocal() as db:
                task = task_repo.get_task_by_task_id(db, reserved_task.task_id)
                if task:
                    params = json.loads(task.params_json or "{}")
                    params["ai_rule_file_path"] = ai_rule_file_path
                    task.params_json = json.dumps(params, ensure_ascii=False)
                    db.commit()

            self._update_task_input_files(
                reserved_task.task_id,
                {
                    "original_path": str(original_path).replace("\\", "/"),
                    "translated_path": str(translated_path).replace("\\", "/"),
                },
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_alignment_task(
        self,
        *,
        original_file: UploadFile,
        translated_file: UploadFile,
        source_lang: str,
        target_lang: str,
        model_name: str,
        gemini_route: str,
        enable_post_split: bool,
        threshold_2: int,
        threshold_3: int,
        threshold_4: int,
        threshold_5: int,
        threshold_6: int,
        threshold_7: int,
        threshold_8: int,
        buffer_chars: int,
    ) -> str:
        reserved_task = self._create_db_task(
            task_type="alignment",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={
                "source_lang": source_lang,
                "target_lang": target_lang,
                "model_name": model_name,
                "gemini_route": gemini_route,
                "enable_post_split": enable_post_split,
                "threshold_2": threshold_2,
                "threshold_3": threshold_3,
                "threshold_4": threshold_4,
                "threshold_5": threshold_5,
                "threshold_6": threshold_6,
                "threshold_7": threshold_7,
                "threshold_8": threshold_8,
                "buffer_chars": buffer_chars,
            },
            input_files={},
        )
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / "alignment" / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)

            orig_ext = Path(original_file.filename or "original.docx").suffix.lower()
            trans_ext = Path(translated_file.filename or "translated.docx").suffix.lower()
            original_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                original_file.filename,
                reserved_task.task_id,
                role="original",
                ext=orig_ext,
            )
            translated_path = upload_dir / build_storage_filename(
                reserved_task.display_no,
                translated_file.filename,
                reserved_task.task_id,
                role="translated",
                ext=trans_ext,
            )
            original_path.write_bytes(await original_file.read())
            translated_path.write_bytes(await translated_file.read())

            self._update_task_input_files(
                reserved_task.task_id,
                {
                    "original_path": str(original_path).replace("\\", "/"),
                    "translated_path": str(translated_path).replace("\\", "/"),
                },
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_doc_translate_task(
        self,
        *,
        file: UploadFile,
        source_lang: str,
        target_langs: str,
        ocr_model: str,
        gemini_route: str,
    ) -> str:
        reserved_task = self._create_db_task(
            task_type="doc_translate",
            filename=file.filename or "input.bin",
            params={"source_lang": source_lang, "target_langs": target_langs, "ocr_model": ocr_model, "gemini_route": gemini_route},
            input_files={},
        )
        try:
            input_path, original_filename = await self._save_single_upload(
                file,
                "doc_translate",
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {"input_path": input_path, "original_filename": original_filename},
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_pdf2docx_task(
        self,
        *,
        file: UploadFile,
        model: str,
        gemini_route: str,
    ) -> str:
        reserved_task = self._create_db_task(
            task_type="pdf2docx",
            filename=file.filename or "input.bin",
            params={"model": model, "gemini_route": gemini_route},
            input_files={},
        )
        try:
            input_path, original_filename = await self._save_single_upload(
                file,
                "pdf2docx",
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {"input_path": input_path, "original_filename": original_filename},
            )
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    def get_task_status(self, task_id: str) -> Optional[Dict[str, Any]]:
        with SessionLocal() as db:
            task = task_repo.get_task_by_task_id(db, task_id)
            if not task:
                return None
            result = json.loads(task.result_json) if task.result_json else None
            tasks_ahead = task_repo.count_tasks_ahead(db, task) if task.status == "queued" else 0
            status_map = {"queued": "queued", "running": "processing", "done": "done", "failed": "failed"}
            payload: Dict[str, Any] = {
                "display_no": task.display_no,
                "status": status_map.get(task.status, task.status),
                "progress": task.progress,
                "message": task.message or "",
                "details": [],
                "result": result,
                "error": task.error_message,
                "stream_log": self._task_logs.get(task_id, ""),
            }
            if task.status == "queued":
                payload["queue_position"] = tasks_ahead + 1
                payload["tasks_ahead"] = tasks_ahead
                payload["message"] = task.message or f"任务排队中，前方还有 {tasks_ahead} 个任务"
            return payload

    def _create_db_task(self, *, task_type: str, filename: str, params: Dict[str, Any], input_files: Dict[str, Any]):
        task_id = str(uuid.uuid4())
        self._set_task_log(task_id, f"[queue] 任务已提交，等待处理: {filename}")
        with SessionLocal() as db:
            task = task_repo.create_task(
                db,
                task_id=task_id,
                task_type=task_type,
                filename=filename,
                status="queued",
                progress=0,
                message="任务已提交，正在排队等待处理",
                params_json=json.dumps(params, ensure_ascii=False),
                input_files_json=json.dumps(input_files, ensure_ascii=False),
            )
        return task

    def _update_task_input_files(self, task_id: str, input_files: Dict[str, Any]) -> None:
        with SessionLocal() as db:
            task_repo.update_task_input_files(db, task_id, json.dumps(input_files, ensure_ascii=False))

    def _fail_reserved_task(self, task_id: str, exc: Exception) -> None:
        self._append_task_log(task_id, f"[error] 保存上传文件失败: {exc}")
        with SessionLocal() as db:
            task_repo.fail_task(db, task_id, f"保存上传文件失败: {exc}")

    def _trim_task_log(self, text: str) -> str:
        if len(text) <= self._max_log_chars:
            return text
        return text[-self._max_log_chars :]

    def _set_task_log(self, task_id: str, text: str) -> None:
        normalized = self._trim_task_log(text or "")
        self._task_logs[task_id] = normalized
        lines = [line for line in normalized.splitlines() if line.strip()]
        if lines:
            self._last_log_line[task_id] = lines[-1]

    def _append_task_log(self, task_id: str, message: str) -> None:
        line = (message or "").strip()
        if not line:
            return
        if self._last_log_line.get(task_id) == line:
            return
        current = self._task_logs.get(task_id, "")
        combined = f"{current}\n{line}" if current else line
        self._task_logs[task_id] = self._trim_task_log(combined)
        self._last_log_line[task_id] = line

    async def _save_single_upload(self, file: UploadFile, folder: str, display_no: str, task_id: str):
        upload_dir = Path(settings.UPLOAD_DIR) / folder / display_no
        upload_dir.mkdir(parents=True, exist_ok=True)
        file_ext = Path(file.filename or "input.bin").suffix or ".bin"
        input_path = upload_dir / build_storage_filename(display_no, file.filename, task_id, ext=file_ext)
        input_path.write_bytes(await file.read())
        return str(input_path).replace("\\", "/"), file.filename or input_path.name

    def _requeue_interrupted_tasks(self):
        with SessionLocal() as db:
            task_repo.requeue_running_tasks(db)

    async def _worker_loop(self):
        while not self._stop_event.is_set():
            with SessionLocal() as db:
                task = task_repo.claim_next_queued_task(db)
            if not task:
                await asyncio.sleep(1)
                continue
            await self._execute_task(task.task_id)

    async def _execute_task(self, task_id: str):
        with SessionLocal() as db:
            task = task_repo.get_task_by_task_id(db, task_id)
            if not task:
                return
            params = json.loads(task.params_json or "{}")
            input_files = json.loads(task.input_files_json or "{}")
            task_type = task.task_type
            filename = task.filename
            display_no = task.display_no

        async def update(progress: int, message: str):
            self._append_task_log(task_id, f"[{progress:>3}%] {message}")
            with SessionLocal() as db:
                task_repo.update_task_progress(db, task_id, progress=progress, message=message, status="running")

        try:
            self._append_task_log(task_id, f"[start] 开始处理任务类型: {task_type}")
            if task_type == "ocr":
                result = await execute_ocr_task_from_path(
                    task_id=task_id,
                    display_no=display_no,
                    input_path=input_files["input_path"],
                    original_filename=input_files.get("original_filename") or filename,
                    progress_callback=update,
                    executor=self._task_executor,
                    **params,
                )
                output_path = result.get("results", [{}])[0].get("translated_image") if result.get("results") else None
            elif task_type == "number_check":
                result = await self._execute_number_check(task_id, display_no, input_files, params, update)
                output_path = result.get("corrected_docx")
            elif task_type == "zhongfanyi":
                result = await self._execute_zhongfanyi(task_id, display_no, input_files, params, update)
                output_path = result.get("corrected_docx")
            elif task_type == "alignment":
                result = await self._execute_alignment(task_id, display_no, input_files, params, update)
                output_path = result.get("output_excel") if result else None
            elif task_type == "doc_translate":
                result = await self._execute_doc_translate(task_id, display_no, input_files, params, update)
                output_path = result.get("raw_output_txt") if result else None
            elif task_type == "pdf2docx":
                result = await self._execute_pdf2docx(task_id, display_no, input_files, params, update)
                output_path = result.get("output_docx") if result else None
            else:
                raise ValueError(f"不支持的任务类型: {task_type}")

            output_files = self._extract_output_files(task_type, result, output_path, filename)
            with SessionLocal() as db:
                task_repo.complete_task(
                    db,
                    task_id,
                    result_json=json.dumps(result, ensure_ascii=False) if result is not None else None,
                    output_path=output_path,
                    output_files_json=json.dumps(output_files, ensure_ascii=False) if output_files else None,
                )
            self._append_task_log(task_id, "[done] 任务处理完成")
        except Exception as exc:
            self._append_task_log(task_id, f"[error] {type(exc).__name__}: {exc}")
            brief_tb = traceback.format_exc(limit=5)
            if brief_tb:
                self._append_task_log(task_id, brief_tb.rstrip())
            with SessionLocal() as db:
                task_repo.fail_task(db, task_id, str(exc))

    async def _execute_number_check(
        self,
        task_id: str,
        display_no: str,
        input_files: Dict[str, Any],
        params: Dict[str, Any],
        update: Callable[[int, str], Any],
    ) -> Dict[str, Any]:
        await update(5, "开始执行数字专检")
        original_bytes = Path(input_files["original_path"]).read_bytes()
        translated_bytes = Path(input_files["translated_path"]).read_bytes()
        original_upload = UploadFile(
            filename=input_files.get("original_filename") or "original.docx",
            file=io.BytesIO(original_bytes),
        )
        translated_upload = UploadFile(
            filename=input_files.get("translated_filename") or "translated.docx",
            file=io.BytesIO(translated_bytes),
        )
        job = asyncio.create_task(
            run_number_check_task(
                original_upload,
                translated_upload,
                task_id=task_id,
                display_no=display_no,
                gemini_route=params.get("gemini_route", "google"),
                model_name=params.get("model_name", "gemini-3-flash-preview"),
            )
        )
        await self._mirror_progress(task_id, job, lambda: get_number_check_progress(task_id), update)
        return await job

    async def _execute_zhongfanyi(
        self,
        task_id: str,
        display_no: str,
        input_files: Dict[str, Any],
        params: Dict[str, Any],
        update: Callable[[int, str], Any],
    ) -> Dict[str, Any]:
        await update(5, "开始执行中翻译专检")
        loop = asyncio.get_running_loop()
        job = loop.run_in_executor(
            self._task_executor,
            lambda: zf_service.run_zhongfanyi_task(
                input_files["original_path"],
                input_files["translated_path"],
                task_id,
                display_no=display_no,
                use_ai_rule=params.get("use_ai_rule", False),
                gemini_route=params.get("gemini_route", "google"),
                ai_rule_file_path=params.get("ai_rule_file_path"),
                session_rule_text=params.get("session_rule_text"),
            ),
        )
        await self._mirror_progress(task_id, job, lambda: zf_service.get_task_progress(task_id), update)
        return await job

    async def _execute_alignment(
        self,
        task_id: str,
        display_no: str,
        input_files: Dict[str, Any],
        params: Dict[str, Any],
        update: Callable[[int, str], Any],
    ) -> Dict[str, Any]:
        await update(5, "开始执行多语对照记忆处理")
        from app.service import alignment_service

        job = asyncio.create_task(
            alignment_service.run_alignment_task(
                original_path=input_files["original_path"],
                translated_path=input_files["translated_path"],
                task_id=task_id,
                display_no=display_no,
                executor=self._task_executor,
                **params,
            )
        )
        await self._mirror_progress(task_id, job, lambda: alignment_service.get_alignment_progress(task_id), update)
        status = alignment_service.get_alignment_progress(task_id) or {}
        return status.get("result") or {}

    async def _execute_doc_translate(
        self,
        task_id: str,
        display_no: str,
        input_files: Dict[str, Any],
        params: Dict[str, Any],
        update: Callable[[int, str], Any],
    ) -> Dict[str, Any]:
        await update(5, "开始执行文档翻译")
        target_langs_str = params.get("target_langs", "en")
        target_langs = [lang.strip() for lang in target_langs_str.split(",") if lang.strip()]
        return await execute_doc_translate_task(
            task_id=task_id,
            display_no=display_no,
            input_path=input_files["input_path"],
            original_filename=input_files.get("original_filename") or "input.pdf",
            source_lang=params.get("source_lang", "zh"),
            target_langs=target_langs,
            ocr_model=params.get("ocr_model", "google/gemini-3-flash-preview"),
            gemini_route=params.get("gemini_route", "google"),
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_pdf2docx(
        self,
        task_id: str,
        display_no: str,
        input_files: Dict[str, Any],
        params: Dict[str, Any],
        update: Callable[[int, str], Any],
    ) -> Dict[str, Any]:
        await update(5, "开始执行 PDF 转 Word")
        return await execute_pdf2docx_task_from_path(
            task_id=task_id,
            display_no=display_no,
            input_path=input_files["input_path"],
            original_filename=input_files.get("original_filename") or "input.pdf",
            model=params.get("model", "google/gemini-3-flash-preview"),
            gemini_route=params.get("gemini_route", "google"),
            progress_callback=update,
            executor=self._task_executor,
        )

    @staticmethod
    def _extract_output_files(
        task_type: str,
        result: Optional[Dict[str, Any]],
        output_path: Optional[str],
        original_filename: Optional[str] = None,
    ) -> list:
        from datetime import datetime as _dt

        timestamp = _dt.now().strftime("%Y%m%d_%H%M%S")
        stem = Path(original_filename or "file").stem if original_filename else "file"
        if " | " in stem:
            stem = stem.split(" | ")[0]
        stem = stem[:60]

        TASK_SUFFIX = {
            "ocr": "OCR翻译",
            "number_check": "数字专检",
            "zhongfanyi": "中翻译专检",
            "alignment": "对照记忆",
            "doc_translate": "证件翻译",
            "pdf2docx": "文档预处理",
        }
        suffix = TASK_SUFFIX.get(task_type, task_type)

        def _friendly(label: str, real_path: str) -> str:
            ext = Path(real_path).suffix
            return f"{stem}_{timestamp}_{suffix}_{label}{ext}"

        files: list = []
        if not result:
            if output_path:
                files.append({"name": _friendly("输出", output_path), "path": output_path, "type": "output"})
            return files

        def _add(key: str, label: str, ftype: str = "output"):
            val = result.get(key)
            if val and isinstance(val, str):
                files.append({"name": _friendly(label, val), "path": val, "type": ftype})

        if task_type == "ocr":
            for r in result.get("results", []):
                if isinstance(r, dict):
                    if r.get("translated_image"):
                        files.append({"name": _friendly("翻译结果", r["translated_image"]), "path": r["translated_image"], "type": "output"})
                    if r.get("visualization_image"):
                        files.append({"name": _friendly("OCR可视化", r["visualization_image"]), "path": r["visualization_image"], "type": "output"})
        elif task_type == "number_check":
            _add("corrected_docx", "修正译文")
            for item in result.get("json_results", []):
                if isinstance(item, dict) and item.get("path"):
                    files.append({"name": _friendly("检查报告", item["path"]), "path": item["path"], "type": "report"})
        elif task_type == "zhongfanyi":
            _add("corrected_docx", "修正译文")
            _add("report_path", "检查报告")
        elif task_type == "alignment":
            _add("output_excel", "对齐结果")
        elif task_type == "doc_translate":
            _add("raw_output_txt", "OCR原文")
            for lang_result in result.get("translations", []):
                if isinstance(lang_result, dict) and lang_result.get("output_docx"):
                    lang = lang_result.get("lang", "")
                    files.append({"name": _friendly(f"翻译_{lang}", lang_result["output_docx"]), "path": lang_result["output_docx"], "type": "output"})
        elif task_type == "pdf2docx":
            _add("output_docx", "Word文档")

        if not files and output_path:
            files.append({"name": _friendly("输出", output_path), "path": output_path, "type": "output"})
        return files

    async def _mirror_progress(
        self,
        task_id: str,
        job: asyncio.Task,
        getter: Callable[[], Optional[Dict[str, Any]]],
        update: Callable[[int, str], Any],
    ):
        while not job.done():
            snapshot = getter()
            if snapshot:
                if snapshot.get("stream_log"):
                    self._set_task_log(task_id, snapshot.get("stream_log", ""))
                progress = snapshot.get("progress", 0)
                message = snapshot.get("message") or snapshot.get("status") or "正在处理"
                await update(progress, message)
            await asyncio.sleep(1)


task_queue_service = TaskQueueService()
ocr_task_queue = task_queue_service
