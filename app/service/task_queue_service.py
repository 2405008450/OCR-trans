import asyncio
import io
import json
import uuid
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from fastapi import UploadFile

from app.core.config import settings
from app.db.session import SessionLocal
from app.repository import task_repo
from app.service import business_licence_service as bl_service
from app.service import zhongfanyi_service as zf_service
from app.service.llm_service import execute_ocr_task_from_path
from app.service.number_check_service import _get_task_progress as get_number_check_progress
from app.service.number_check_service import run_number_check_task


class TaskQueueService:
    def __init__(self):
        self._worker_task: Optional[asyncio.Task] = None
        self._stop_event = asyncio.Event()

    async def start(self):
        if self._worker_task and not self._worker_task.done():
            return
        self._stop_event = asyncio.Event()
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

    async def submit_ocr_task(self, **kwargs) -> str:
        file: UploadFile = kwargs.pop("file")
        task_id, input_path, original_filename = await self._save_single_upload(file, "ocr")
        return self._create_db_task(
            task_id=task_id,
            task_type="ocr",
            filename=original_filename,
            params=kwargs,
            input_files={"input_path": input_path, "original_filename": original_filename},
        )

    async def submit_number_check_task(self, *, original_file: UploadFile, translated_file: UploadFile) -> str:
        task_id = str(uuid.uuid4())
        upload_dir = Path(settings.UPLOAD_DIR) / "number_check"
        upload_dir.mkdir(parents=True, exist_ok=True)
        original_ext = Path(original_file.filename or "original.docx").suffix or ".docx"
        translated_ext = Path(translated_file.filename or "translated.docx").suffix or ".docx"
        original_path = upload_dir / f"{task_id}_original{original_ext}"
        translated_path = upload_dir / f"{task_id}_translated{translated_ext}"
        original_path.write_bytes(await original_file.read())
        translated_path.write_bytes(await translated_file.read())
        return self._create_db_task(
            task_id=task_id,
            task_type="number_check",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={},
            input_files={
                "original_path": str(original_path).replace("\\", "/"),
                "translated_path": str(translated_path).replace("\\", "/"),
                "original_filename": original_file.filename,
                "translated_filename": translated_file.filename,
            },
        )

    async def submit_zhongfanyi_task(
        self,
        *,
        original_file: UploadFile,
        translated_file: UploadFile,
        use_ai_rule: bool,
        rule_file: Optional[UploadFile],
        session_rule_content: Optional[str],
    ) -> str:
        task_id = str(uuid.uuid4())
        upload_dir = Path(settings.UPLOAD_DIR) / "zhongfanyi" / task_id
        upload_dir.mkdir(parents=True, exist_ok=True)
        ext_orig = Path(original_file.filename or "original.docx").suffix.lower()
        ext_tran = Path(translated_file.filename or "translated.docx").suffix.lower()
        original_path = upload_dir / f"original{ext_orig}"
        translated_path = upload_dir / f"translated{ext_tran}"
        original_path.write_bytes(await original_file.read())
        translated_path.write_bytes(await translated_file.read())

        ai_rule_file_path = None
        if rule_file and use_ai_rule:
            ext_rule = Path(rule_file.filename or "rule.txt").suffix.lower()
            ai_rule_path = upload_dir / f"rule{ext_rule}"
            ai_rule_path.write_bytes(await rule_file.read())
            ai_rule_file_path = str(ai_rule_path).replace("\\", "/")

        return self._create_db_task(
            task_id=task_id,
            task_type="zhongfanyi",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={
                "use_ai_rule": use_ai_rule,
                "ai_rule_file_path": ai_rule_file_path,
                "session_rule_text": (session_rule_content.strip() or None) if session_rule_content else None,
            },
            input_files={
                "original_path": str(original_path).replace("\\", "/"),
                "translated_path": str(translated_path).replace("\\", "/"),
            },
        )

    async def submit_alignment_task(
        self,
        *,
        original_file: UploadFile,
        translated_file: UploadFile,
        source_lang: str,
        target_lang: str,
        model_name: str,
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
        task_id = str(uuid.uuid4())
        upload_dir = Path(settings.UPLOAD_DIR) / "alignment"
        upload_dir.mkdir(parents=True, exist_ok=True)
        orig_ext = Path(original_file.filename or "original.docx").suffix.lower()
        trans_ext = Path(translated_file.filename or "translated.docx").suffix.lower()
        original_path = upload_dir / f"{task_id}_original{orig_ext}"
        translated_path = upload_dir / f"{task_id}_translated{trans_ext}"
        original_path.write_bytes(await original_file.read())
        translated_path.write_bytes(await translated_file.read())
        return self._create_db_task(
            task_id=task_id,
            task_type="alignment",
            filename=f"{original_file.filename} | {translated_file.filename}",
            params={
                "source_lang": source_lang,
                "target_lang": target_lang,
                "model_name": model_name,
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
            input_files={
                "original_path": str(original_path).replace("\\", "/"),
                "translated_path": str(translated_path).replace("\\", "/"),
            },
        )

    async def submit_business_licence_task(
        self,
        *,
        file: UploadFile,
        source_lang: str,
        target_lang: str,
        config_file: str,
    ) -> str:
        task_id, input_path, original_filename = await self._save_single_upload(file, "business_licence")
        await bl_service.prepare_business_licence_task(task_id, original_filename)
        return self._create_db_task(
            task_id=task_id,
            task_type="business_licence",
            filename=original_filename,
            params={"source_lang": source_lang, "target_lang": target_lang, "config_file": config_file},
            input_files={"input_path": input_path, "original_filename": original_filename},
        )

    def get_task_status(self, task_id: str) -> Optional[Dict[str, Any]]:
        with SessionLocal() as db:
            task = task_repo.get_task_by_task_id(db, task_id)
            if not task:
                return None
            result = json.loads(task.result_json) if task.result_json else None
            tasks_ahead = task_repo.count_tasks_ahead(db, task) if task.status == "queued" else 0
            status_map = {"queued": "queued", "running": "processing", "done": "done", "failed": "failed"}
            payload: Dict[str, Any] = {
                "status": status_map.get(task.status, task.status),
                "progress": task.progress,
                "message": task.message or "",
                "details": [],
                "result": result,
                "error": task.error_message,
            }
            if task.status == "queued":
                payload["queue_position"] = tasks_ahead + 1
                payload["tasks_ahead"] = tasks_ahead
                payload["message"] = task.message or f"任务排队中，前方还有 {tasks_ahead} 个任务"
            return payload

    def _create_db_task(self, *, task_id: str, task_type: str, filename: str, params: Dict[str, Any], input_files: Dict[str, Any]) -> str:
        with SessionLocal() as db:
            task_repo.create_task(
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
        return task_id

    async def _save_single_upload(self, file: UploadFile, folder: str):
        task_id = str(uuid.uuid4())
        upload_dir = Path(settings.UPLOAD_DIR) / folder
        upload_dir.mkdir(parents=True, exist_ok=True)
        file_ext = Path(file.filename or "input.bin").suffix or ".bin"
        input_path = upload_dir / f"{task_id}{file_ext}"
        input_path.write_bytes(await file.read())
        return task_id, str(input_path).replace("\\", "/"), file.filename or input_path.name

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

        async def update(progress: int, message: str):
            with SessionLocal() as db:
                task_repo.update_task_progress(db, task_id, progress=progress, message=message, status="running")

        try:
            if task_type == "ocr":
                result = await execute_ocr_task_from_path(
                    task_id=task_id,
                    input_path=input_files["input_path"],
                    original_filename=input_files.get("original_filename") or filename,
                    progress_callback=update,
                    **params,
                )
                output_path = result.get("results", [{}])[0].get("translated_image") if result.get("results") else None
            elif task_type == "number_check":
                result = await self._execute_number_check(task_id, input_files, update)
                output_path = result.get("corrected_docx")
            elif task_type == "zhongfanyi":
                result = await self._execute_zhongfanyi(task_id, input_files, params, update)
                output_path = result.get("corrected_docx")
            elif task_type == "alignment":
                result = await self._execute_alignment(task_id, input_files, params, update)
                output_path = result.get("output_excel") if result else None
            elif task_type == "business_licence":
                result = await self._execute_business_licence(task_id, input_files, params, update)
                output_path = result.get("output_path") if result else None
            else:
                raise ValueError(f"不支持的任务类型: {task_type}")

            with SessionLocal() as db:
                task_repo.complete_task(
                    db,
                    task_id,
                    result_json=json.dumps(result, ensure_ascii=False) if result is not None else None,
                    output_path=output_path,
                )
        except Exception as exc:
            with SessionLocal() as db:
                task_repo.fail_task(db, task_id, str(exc))

    async def _execute_number_check(self, task_id: str, input_files: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, "开始执行数字专检")
        original_bytes = Path(input_files["original_path"]).read_bytes()
        translated_bytes = Path(input_files["translated_path"]).read_bytes()
        original_upload = UploadFile(filename=input_files.get("original_filename") or "original.docx", file=io.BytesIO(original_bytes))
        translated_upload = UploadFile(filename=input_files.get("translated_filename") or "translated.docx", file=io.BytesIO(translated_bytes))
        job = asyncio.create_task(run_number_check_task(original_upload, translated_upload, task_id=task_id))
        await self._mirror_progress(job, lambda: get_number_check_progress(task_id), update)
        return await job

    async def _execute_zhongfanyi(self, task_id: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, "开始执行中翻译专检")
        job = asyncio.create_task(asyncio.to_thread(
            zf_service.run_zhongfanyi_task,
            input_files["original_path"],
            input_files["translated_path"],
            task_id,
            params.get("use_ai_rule", False),
            params.get("ai_rule_file_path"),
            params.get("session_rule_text"),
        ))
        await self._mirror_progress(job, lambda: zf_service.get_task_progress(task_id), update)
        return await job

    async def _execute_alignment(self, task_id: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, "开始执行多语对照记忆处理")
        from app.service import alignment_service

        job = asyncio.create_task(alignment_service.run_alignment_task(
            original_path=input_files["original_path"],
            translated_path=input_files["translated_path"],
            task_id=task_id,
            **params,
        ))
        await self._mirror_progress(job, lambda: alignment_service.get_alignment_progress(task_id), update)
        status = alignment_service.get_alignment_progress(task_id) or {}
        return status.get("result") or {}

    async def _execute_business_licence(self, task_id: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, "开始执行营业执照翻译")
        await bl_service.start_prepared_business_licence_task(
            task_id=task_id,
            input_path=input_files["input_path"],
            source_lang=params["source_lang"],
            target_lang=params["target_lang"],
            config_file=params["config_file"],
        )
        while True:
            snapshot = bl_service.get_task_snapshot(task_id)
            if snapshot:
                await update(snapshot.get("progress", 0), snapshot.get("message", "????"))
                if snapshot.get("status") == "done":
                    return {
                        "task_id": task_id,
                        "output_path": snapshot.get("output_path"),
                        "input_filename": snapshot.get("input_filename"),
                    }
                if snapshot.get("status") == "error":
                    raise RuntimeError(snapshot.get("error") or "??????????")
            await asyncio.sleep(1)

    async def _mirror_progress(self, job: asyncio.Task, getter: Callable[[], Optional[Dict[str, Any]]], update: Callable[[int, str], Any]):
        while not job.done():
            snapshot = getter()
            if snapshot:
                progress = snapshot.get("progress", 0)
                message = snapshot.get("message") or snapshot.get("status") or "正在处理"
                await update(progress, message)
            await asyncio.sleep(1)


task_queue_service = TaskQueueService()
ocr_task_queue = task_queue_service
