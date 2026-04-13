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
from app.core.file_naming import build_storage_filename, build_user_visible_filename
from app.db.session import SessionLocal
from app.repository import task_repo
from app.service import zhongfanyi_service as zf_service
from app.service.business_licence_service import (
    BUSINESS_LICENCE_DEFAULT_MODEL,
    BUSINESS_LICENCE_DEFAULT_ROUTE,
    execute_business_licence_task,
)
from app.service.doc_translate_service import execute_doc_translate_task
from app.service.drivers_license_service import execute_drivers_license_task
from app.service.llm_service import execute_ocr_task_from_path
from app.service.number_check_service import _get_task_progress as get_number_check_progress, run_number_check_task
from app.service.pdf2docx_service import (
    PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    PDF2DOCX_DEFAULT_MODEL,
    execute_pdf2docx_task_from_path,
)


class TaskCancelledError(Exception):
    pass


class TaskQueueService:
    MAX_AUTO_REQUEUE_ATTEMPTS = 1
    DEFAULT_TASK_TYPE_LIMITS: Dict[str, int] = {
        'ocr': 1,
        'pdf2docx': 1,
        'doc_translate': 1,
        'alignment': 1,
        'drivers_license': 1,
        'business_licence': 2,
        'number_check': 2,
        'zhongfanyi': 2,
    }
    SHARED_TASK_GROUPS: Dict[str, str] = {
        'number_check': 'specialist_text',
        'zhongfanyi': 'specialist_text',
    }
    SHARED_GROUP_LIMITS: Dict[str, int] = {
        'specialist_text': 1,
    }
    INPUT_FILES_WAIT_SECONDS = 5.0
    INPUT_FILES_POLL_INTERVAL_SECONDS = 0.2

    def __init__(self):
        self._worker_task: Optional[asyncio.Task] = None
        self._stop_event = asyncio.Event()
        self._dispatch_event = asyncio.Event()
        self._task_executor: Optional[ThreadPoolExecutor] = None
        self._task_logs: Dict[str, str] = {}
        self._last_log_line: Dict[str, str] = {}
        self._max_log_chars = 50000
        self._running_tasks: Dict[str, asyncio.Task] = {}
        self._running_task_types: Dict[str, str] = {}
        self._running_counts: Dict[str, int] = {}
        self._running_group_counts: Dict[str, int] = {}
        self._max_concurrent_tasks = max(1, settings.TASK_QUEUE_MAX_CONCURRENT_TASKS)
        self._poll_interval_seconds = max(0.1, settings.TASK_QUEUE_POLL_INTERVAL_SECONDS)
        self._candidate_batch_size = max(1, settings.TASK_QUEUE_CANDIDATE_BATCH_SIZE)
        self._task_type_limits = self._build_task_type_limits()

    @classmethod
    def _build_task_type_limits(cls) -> Dict[str, int]:
        limits = dict(cls.DEFAULT_TASK_TYPE_LIMITS)
        for task_type, limit in settings.TASK_QUEUE_TYPE_LIMITS.items():
            limits[task_type] = limit
        return limits

    async def start(self):
        if self._worker_task and not self._worker_task.done():
            return
        self._stop_event = asyncio.Event()
        self._dispatch_event = asyncio.Event()
        self._running_tasks = {}
        self._running_task_types = {}
        self._running_counts = {}
        self._running_group_counts = {}
        if self._task_executor is None:
            executor_workers = max(
                1,
                settings.TASK_QUEUE_EXECUTOR_MAX_WORKERS,
                self._max_concurrent_tasks,
            )
            self._task_executor = ThreadPoolExecutor(max_workers=executor_workers, thread_name_prefix='task-queue')
        self._requeue_interrupted_tasks()
        self._worker_task = asyncio.create_task(self._worker_loop(), name='task-queue-dispatcher')

    async def stop(self):
        if not self._worker_task:
            return
        self._stop_event.set()
        self._dispatch_event.set()
        self._worker_task.cancel()
        running_tasks = list(self._running_tasks.values())
        for running_task in running_tasks:
            running_task.cancel()
        try:
            await self._worker_task
            if running_tasks:
                await asyncio.gather(*running_tasks, return_exceptions=True)
        except asyncio.CancelledError:
            pass
        finally:
            self._worker_task = None
            self._running_tasks = {}
            self._running_task_types = {}
            self._running_counts = {}
            self._running_group_counts = {}
            if self._task_executor is not None:
                self._task_executor.shutdown(wait=False, cancel_futures=True)
                self._task_executor = None

    async def submit_ocr_task(self, **kwargs) -> str:
        file: UploadFile = kwargs.pop('file')
        reserved_task = self._create_db_task('ocr', file.filename or 'input.bin', kwargs, {})
        try:
            input_path, original_filename = await self._save_single_upload(file, 'ocr', reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': original_filename})
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_number_check_task(
        self,
        *,
        mode: str,
        original_file: Optional[UploadFile],
        translated_file: Optional[UploadFile],
        single_file: Optional[UploadFile],
        gemini_route: str,
        model_name: str,
    ) -> str:
        if mode == 'single' and single_file:
            display_name = single_file.filename or 'single.docx'
        else:
            original_name = original_file.filename if original_file else 'original.docx'
            translated_name = translated_file.filename if translated_file else 'translated.docx'
            display_name = f'{original_name} | {translated_name}'
        reserved_task = self._create_db_task(
            'number_check',
            display_name,
            {'mode': mode, 'gemini_route': gemini_route, 'model_name': model_name},
            {},
        )
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / 'number_check' / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)
            if mode == 'single':
                if single_file is None:
                    raise ValueError('single_file is required for single mode')
                single_ext = Path(single_file.filename or 'single.docx').suffix or '.docx'
                single_path = upload_dir / build_storage_filename(
                    reserved_task.display_no,
                    single_file.filename,
                    reserved_task.task_id,
                    role='single',
                    ext=single_ext,
                )
                single_path.write_bytes(await single_file.read())
                self._update_task_input_files(
                    reserved_task.task_id,
                    {
                        'single_path': str(single_path).replace('\\', '/'),
                        'single_filename': single_file.filename,
                    },
                )
                self._notify_dispatcher()
                return reserved_task.task_id

            if original_file is None or translated_file is None:
                raise ValueError('original_file and translated_file are required for double mode')

            original_ext = Path(original_file.filename or 'original.docx').suffix or '.docx'
            translated_ext = Path(translated_file.filename or 'translated.docx').suffix or '.docx'
            original_path = upload_dir / build_storage_filename(reserved_task.display_no, original_file.filename, reserved_task.task_id, role='original', ext=original_ext)
            translated_path = upload_dir / build_storage_filename(reserved_task.display_no, translated_file.filename, reserved_task.task_id, role='translated', ext=translated_ext)
            original_path.write_bytes(await original_file.read())
            translated_path.write_bytes(await translated_file.read())
            self._update_task_input_files(reserved_task.task_id, {'original_path': str(original_path).replace('\\', '/'), 'translated_path': str(translated_path).replace('\\', '/'), 'original_filename': original_file.filename, 'translated_filename': translated_file.filename})
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_zhongfanyi_task(self, *, mode: str, original_file: Optional[UploadFile], translated_file: Optional[UploadFile], single_file: Optional[UploadFile], use_ai_rule: bool, gemini_route: str, model_name: str, rule_file: Optional[UploadFile], session_rule_content: Optional[str]) -> str:
        if mode == zf_service.ZHONGFANYI_MODE_SINGLE:
            display_name = single_file.filename if single_file else 'single.docx'
        else:
            original_name = original_file.filename if original_file else 'original.docx'
            translated_name = translated_file.filename if translated_file else 'translated.docx'
            display_name = f'{original_name} | {translated_name}'
        reserved_task = self._create_db_task('zhongfanyi', display_name, {'mode': mode, 'use_ai_rule': use_ai_rule, 'gemini_route': gemini_route, 'model_name': model_name, 'ai_rule_file_path': None, 'session_rule_text': (session_rule_content.strip() or None) if session_rule_content else None}, {})
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / 'zhongfanyi' / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)
            input_files = {}
            if mode == zf_service.ZHONGFANYI_MODE_SINGLE:
                ext_single = Path(single_file.filename or 'single.docx').suffix.lower()
                single_path = upload_dir / build_storage_filename(reserved_task.display_no, single_file.filename, reserved_task.task_id, role='single', ext=ext_single)
                single_path.write_bytes(await single_file.read())
                input_files.update({'single_path': str(single_path).replace('\\', '/'), 'single_filename': single_file.filename})
            else:
                ext_orig = Path(original_file.filename or 'original.docx').suffix.lower()
                ext_tran = Path(translated_file.filename or 'translated.docx').suffix.lower()
                original_path = upload_dir / build_storage_filename(reserved_task.display_no, original_file.filename, reserved_task.task_id, role='original', ext=ext_orig)
                translated_path = upload_dir / build_storage_filename(reserved_task.display_no, translated_file.filename, reserved_task.task_id, role='translated', ext=ext_tran)
                original_path.write_bytes(await original_file.read())
                translated_path.write_bytes(await translated_file.read())
                input_files.update({
                    'original_path': str(original_path).replace('\\', '/'),
                    'translated_path': str(translated_path).replace('\\', '/'),
                    'original_filename': original_file.filename,
                    'translated_filename': translated_file.filename,
                })
            ai_rule_file_path = None
            if rule_file and use_ai_rule:
                ext_rule = Path(rule_file.filename or 'rule.txt').suffix.lower()
                ai_rule_path = upload_dir / build_storage_filename(reserved_task.display_no, rule_file.filename, reserved_task.task_id, role='rule', ext=ext_rule)
                ai_rule_path.write_bytes(await rule_file.read())
                ai_rule_file_path = str(ai_rule_path).replace('\\', '/')
            with SessionLocal() as db:
                task = task_repo.get_task_by_task_id(db, reserved_task.task_id)
                if task:
                    params = json.loads(task.params_json or '{}')
                    params['ai_rule_file_path'] = ai_rule_file_path
                    task.params_json = json.dumps(params, ensure_ascii=False)
                    db.commit()
            self._update_task_input_files(reserved_task.task_id, input_files)
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_alignment_task(self, *, original_file: UploadFile, translated_file: UploadFile, source_lang: str, target_lang: str, model_name: str, gemini_route: str, enable_post_split: bool, threshold_2: int, threshold_3: int, threshold_4: int, threshold_5: int, threshold_6: int, threshold_7: int, threshold_8: int, buffer_chars: int) -> str:
        reserved_task = self._create_db_task('alignment', f'{original_file.filename} | {translated_file.filename}', {'source_lang': source_lang, 'target_lang': target_lang, 'model_name': model_name, 'gemini_route': gemini_route, 'enable_post_split': enable_post_split, 'threshold_2': threshold_2, 'threshold_3': threshold_3, 'threshold_4': threshold_4, 'threshold_5': threshold_5, 'threshold_6': threshold_6, 'threshold_7': threshold_7, 'threshold_8': threshold_8, 'buffer_chars': buffer_chars}, {})
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / 'alignment' / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)
            orig_ext = Path(original_file.filename or 'original.docx').suffix.lower()
            trans_ext = Path(translated_file.filename or 'translated.docx').suffix.lower()
            original_path = upload_dir / build_storage_filename(reserved_task.display_no, original_file.filename, reserved_task.task_id, role='original', ext=orig_ext)
            translated_path = upload_dir / build_storage_filename(reserved_task.display_no, translated_file.filename, reserved_task.task_id, role='translated', ext=trans_ext)
            original_path.write_bytes(await original_file.read())
            translated_path.write_bytes(await translated_file.read())
            self._update_task_input_files(
                reserved_task.task_id,
                {
                    'original_path': str(original_path).replace('\\', '/'),
                    'translated_path': str(translated_path).replace('\\', '/'),
                    'original_filename': original_file.filename,
                    'translated_filename': translated_file.filename,
                },
            )
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_doc_translate_task(self, *, file: UploadFile, source_lang: str, target_langs: str, ocr_model: str, gemini_route: str) -> str:
        reserved_task = self._create_db_task('doc_translate', file.filename or 'input.bin', {'source_lang': source_lang, 'target_langs': target_langs, 'ocr_model': ocr_model, 'gemini_route': gemini_route}, {})
        try:
            input_path, original_filename = await self._save_single_upload(file, 'doc_translate', reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': original_filename})
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_business_licence_task(
        self,
        *,
        file: UploadFile,
        model: str,
        gemini_route: str,
        parsed_data: Optional[Dict[str, Any]] = None,
        company_name_override: Optional[str] = None,
    ) -> str:
        reserved_task = self._create_db_task(
            'business_licence',
            file.filename or 'input.bin',
            {
                'model': model,
                'gemini_route': gemini_route,
                'parsed_data': parsed_data,
                'company_name_override': company_name_override,
            },
            {},
        )
        try:
            input_path, original_filename = await self._save_single_upload(file, 'business_licence', reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': original_filename})
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_pdf2docx_task(self, *, file: UploadFile, model: str, gemini_route: str) -> str:
        reserved_task = self._create_db_task('pdf2docx', file.filename or 'input.bin', {'model': model, 'gemini_route': gemini_route}, {})
        try:
            input_path, original_filename = await self._save_single_upload(file, 'pdf2docx', reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': original_filename})
            self._notify_dispatcher()
            return reserved_task.task_id
        except Exception as exc:
            self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_drivers_license_task(self, *, files: list[UploadFile], processing_mode: str) -> str:
        reserved_task = self._create_db_task('drivers_license', ' | '.join([(file.filename or 'input.bin') for file in files]), {'processing_mode': processing_mode}, {})
        try:
            upload_dir = Path(settings.UPLOAD_DIR) / 'drivers_license' / reserved_task.display_no
            upload_dir.mkdir(parents=True, exist_ok=True)
            saved_files = []
            for index, file in enumerate(files, start=1):
                file_ext = Path(file.filename or f'image_{index}.bin').suffix or '.bin'
                input_path = upload_dir / build_storage_filename(reserved_task.display_no, file.filename, reserved_task.task_id, role=f'image{index:02d}', ext=file_ext)
                input_path.write_bytes(await file.read())
                saved_files.append({'index': index, 'path': str(input_path).replace('\\', '/'), 'original_filename': file.filename or input_path.name})
            self._update_task_input_files(reserved_task.task_id, {'files': saved_files})
            self._notify_dispatcher()
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
            tasks_ahead = task_repo.count_tasks_ahead(db, task) if task.status == 'queued' else 0
            payload = {'display_no': task.display_no, 'task_id': task.task_id, 'status': {'queued': 'queued', 'running': 'processing', 'done': 'done', 'failed': 'failed', 'cancelled': 'cancelled'}.get(task.status, task.status), 'progress': task.progress, 'message': task.message or '', 'details': [], 'result': result, 'error': task.error_message, 'stream_log': self._task_logs.get(task_id, ''), 'created_at': task.created_at.isoformat() if task.created_at else None, 'started_at': task.started_at.isoformat() if task.started_at else None, 'finished_at': task.finished_at.isoformat() if task.finished_at else None}
            if task.status == 'queued':
                payload['queue_position'] = tasks_ahead + 1
                payload['tasks_ahead'] = tasks_ahead
                payload['message'] = task.message or f'Queued, {tasks_ahead} task(s) ahead'
            return payload

    def _create_db_task(self, task_type: str, filename: str, params: Dict[str, Any], input_files: Dict[str, Any]):
        task_id = str(uuid.uuid4())
        self._set_task_log(task_id, f'[queue] queued: {filename}')
        with SessionLocal() as db:
            return task_repo.create_task(db, task_id=task_id, task_type=task_type, filename=filename, status='queued', progress=0, message='Queued', params_json=json.dumps(params, ensure_ascii=False), input_files_json=json.dumps(input_files, ensure_ascii=False))

    def _update_task_input_files(self, task_id: str, input_files: Dict[str, Any]) -> None:
        with SessionLocal() as db:
            task_repo.update_task_input_files(db, task_id, json.dumps(input_files, ensure_ascii=False))

    def _fail_reserved_task(self, task_id: str, exc: Exception) -> None:
        self._append_task_log(task_id, f'[error] upload save failed: {exc}')
        with SessionLocal() as db:
            task_repo.fail_task(db, task_id, f'upload save failed: {exc}')

    def _trim_task_log(self, text: str) -> str:
        return text if len(text) <= self._max_log_chars else text[-self._max_log_chars:]

    def _set_task_log(self, task_id: str, text: str) -> None:
        normalized = self._trim_task_log(text or '')
        self._task_logs[task_id] = normalized
        lines = [line for line in normalized.splitlines() if line.strip()]
        if lines:
            self._last_log_line[task_id] = lines[-1]

    def _append_task_log(self, task_id: str, message: str) -> None:
        line = (message or '').strip()
        if not line or self._last_log_line.get(task_id) == line:
            return
        current = self._task_logs.get(task_id, '')
        self._task_logs[task_id] = self._trim_task_log(f'{current}\n{line}' if current else line)
        self._last_log_line[task_id] = line

    async def _save_single_upload(self, file: UploadFile, folder: str, display_no: str, task_id: str):
        upload_dir = Path(settings.UPLOAD_DIR) / folder / display_no
        upload_dir.mkdir(parents=True, exist_ok=True)
        file_ext = Path(file.filename or 'input.bin').suffix or '.bin'
        input_path = upload_dir / build_storage_filename(display_no, file.filename, task_id, ext=file_ext)
        input_path.write_bytes(await file.read())
        return str(input_path).replace('\\', '/'), file.filename or input_path.name

    @staticmethod
    def _get_input_value(input_files: Dict[str, Any], *candidates: str):
        for key in candidates:
            value = input_files.get(key)
            if value:
                return value
        return None

    def _get_missing_input_fields(self, task_type: str, params: Dict[str, Any], input_files: Dict[str, Any]) -> list[str]:
        if task_type in {'ocr', 'doc_translate', 'business_licence', 'pdf2docx'}:
            return [] if self._get_input_value(input_files, 'input_path') else ['input_path']

        if task_type == 'drivers_license':
            return [] if (input_files.get('files') or []) else ['files']

        if task_type == 'alignment':
            missing = []
            if not self._get_input_value(input_files, 'original_path', 'original'):
                missing.append('original_path')
            if not self._get_input_value(input_files, 'translated_path', 'translated'):
                missing.append('translated_path')
            return missing

        if task_type == 'number_check':
            mode = params.get('mode', 'double')
            if mode == 'single':
                return [] if self._get_input_value(input_files, 'single_path') else ['single_path']
            missing = []
            if not self._get_input_value(input_files, 'original_path'):
                missing.append('original_path')
            if not self._get_input_value(input_files, 'translated_path'):
                missing.append('translated_path')
            return missing

        if task_type == 'zhongfanyi':
            mode = params.get('mode', zf_service.ZHONGFANYI_MODE_DOUBLE)
            if mode == zf_service.ZHONGFANYI_MODE_SINGLE:
                return [] if self._get_input_value(input_files, 'single_path') else ['single_path']
            missing = []
            if not self._get_input_value(input_files, 'original_path'):
                missing.append('original_path')
            if not self._get_input_value(input_files, 'translated_path'):
                missing.append('translated_path')
            return missing

        return []

    async def _ensure_task_input_files(self, task_id: str, task_type: str, params: Dict[str, Any], input_files: Dict[str, Any]) -> Dict[str, Any]:
        missing = self._get_missing_input_fields(task_type, params, input_files)
        if not missing:
            return input_files

        self._append_task_log(task_id, f"[wait] waiting for input files: {', '.join(missing)}")
        latest_input_files = dict(input_files)
        loop = asyncio.get_running_loop()
        deadline = loop.time() + self.INPUT_FILES_WAIT_SECONDS

        while missing and loop.time() < deadline:
            await asyncio.sleep(self.INPUT_FILES_POLL_INTERVAL_SECONDS)
            with SessionLocal() as db:
                task = task_repo.get_task_by_task_id(db, task_id)
                latest_input_files = json.loads(task.input_files_json or '{}') if task else {}
            missing = self._get_missing_input_fields(task_type, params, latest_input_files)

        if missing:
            raise ValueError(f"task input files not ready: {', '.join(missing)}")

        return latest_input_files

    def _requeue_interrupted_tasks(self):
        with SessionLocal() as db:
            task_repo.requeue_running_tasks(
                db,
                max_retry_count=self.MAX_AUTO_REQUEUE_ATTEMPTS,
            )

    def _notify_dispatcher(self) -> None:
        if not self._dispatch_event.is_set():
            self._dispatch_event.set()

    def _can_start_task_type(self, task_type: str) -> bool:
        task_limit = self._task_type_limits.get(task_type, self._max_concurrent_tasks)
        if self._running_counts.get(task_type, 0) >= task_limit:
            return False

        group_name = self.SHARED_TASK_GROUPS.get(task_type)
        if not group_name:
            return True

        group_limit = self.SHARED_GROUP_LIMITS.get(group_name, task_limit)
        return self._running_group_counts.get(group_name, 0) < group_limit

    def _reserve_task_slot(self, task_id: str, task_type: str) -> None:
        self._running_task_types[task_id] = task_type
        self._running_counts[task_type] = self._running_counts.get(task_type, 0) + 1

        group_name = self.SHARED_TASK_GROUPS.get(task_type)
        if group_name:
            self._running_group_counts[group_name] = self._running_group_counts.get(group_name, 0) + 1

    def _release_task_slot(self, task_id: str) -> None:
        task_type = self._running_task_types.pop(task_id, None)
        if not task_type:
            return

        remaining = self._running_counts.get(task_type, 0) - 1
        if remaining > 0:
            self._running_counts[task_type] = remaining
        else:
            self._running_counts.pop(task_type, None)

        group_name = self.SHARED_TASK_GROUPS.get(task_type)
        if group_name:
            group_remaining = self._running_group_counts.get(group_name, 0) - 1
            if group_remaining > 0:
                self._running_group_counts[group_name] = group_remaining
            else:
                self._running_group_counts.pop(group_name, None)

    def _claim_next_dispatchable_task(self):
        with SessionLocal() as db:
            queued_tasks = task_repo.list_queued_tasks(db, limit=self._candidate_batch_size)
            for queued_task in queued_tasks:
                if not self._can_start_task_type(queued_task.task_type):
                    continue
                claimed_task = task_repo.claim_queued_task_by_task_id(db, queued_task.task_id)
                if claimed_task:
                    return claimed_task
        return None

    def _start_claimed_task(self, task_id: str, task_type: str) -> None:
        self._reserve_task_slot(task_id, task_type)
        runner = asyncio.create_task(self._execute_task(task_id), name=f'task-{task_type}-{task_id[:8]}')
        self._running_tasks[task_id] = runner

        def _on_done(done_task: asyncio.Task, *, claimed_task_id: str):
            self._running_tasks.pop(claimed_task_id, None)
            self._release_task_slot(claimed_task_id)
            try:
                done_task.result()
            except asyncio.CancelledError:
                pass
            except Exception as exc:
                self._append_task_log(claimed_task_id, f'[error] runner crashed: {exc}')
            self._notify_dispatcher()

        runner.add_done_callback(
            lambda done_task, claimed_task_id=task_id: _on_done(done_task, claimed_task_id=claimed_task_id)
        )

    async def _dispatch_ready_tasks(self) -> bool:
        dispatched_any = False
        while (
            not self._stop_event.is_set()
            and len(self._running_tasks) < self._max_concurrent_tasks
        ):
            claimed_task = self._claim_next_dispatchable_task()
            if not claimed_task:
                break
            dispatched_any = True
            self._start_claimed_task(claimed_task.task_id, claimed_task.task_type)
        return dispatched_any

    async def _worker_loop(self):
        while not self._stop_event.is_set():
            dispatched_any = await self._dispatch_ready_tasks()
            if dispatched_any:
                continue
            self._dispatch_event.clear()
            try:
                await asyncio.wait_for(self._dispatch_event.wait(), timeout=self._poll_interval_seconds)
            except asyncio.TimeoutError:
                continue

    async def _execute_task(self, task_id: str):
        with SessionLocal() as db:
            task = task_repo.get_task_by_task_id(db, task_id)
            if not task:
                return
            params = json.loads(task.params_json or '{}')
            input_files = json.loads(task.input_files_json or '{}')
            task_type = task.task_type
            filename = task.filename
            display_no = task.display_no

        input_files = await self._ensure_task_input_files(task_id, task_type, params, input_files)

        async def update(progress: int, message: str):
            with SessionLocal() as db:
                if task_repo.is_cancel_requested(db, task_id):
                    task_repo.mark_cancelled(db, task_id)
                    self._append_task_log(task_id, '[cancel] user cancelled')
                    raise TaskCancelledError('user cancelled')
            self._append_task_log(task_id, f'[{progress:>3}%] {message}')
            with SessionLocal() as db:
                task_repo.update_task_progress(db, task_id, progress=progress, message=message, status='running')

        try:
            self._append_task_log(task_id, f'[start] {task_type}')
            if task_type == 'ocr':
                result = await execute_ocr_task_from_path(task_id=task_id, display_no=display_no, input_path=input_files['input_path'], original_filename=input_files.get('original_filename') or filename, progress_callback=update, executor=self._task_executor, **params)
                output_path = result.get('results', [{}])[0].get('translated_image') if result.get('results') else None
            elif task_type == 'number_check':
                result = await self._execute_number_check(task_id, display_no, input_files, params, update)
                output_path = result.get('corrected_docx')
            elif task_type == 'zhongfanyi':
                result = await self._execute_zhongfanyi(task_id, display_no, input_files, params, update)
                output_path = result.get('corrected_docx')
            elif task_type == 'alignment':
                result = await self._execute_alignment(task_id, display_no, input_files, params, update)
                output_path = result.get('output_excel') if result else None
            elif task_type == 'drivers_license':
                result = await self._execute_drivers_license(task_id, display_no, input_files, params, update)
                output_path = result.get('output_docx') if result else None
            elif task_type == 'doc_translate':
                result = await self._execute_doc_translate(task_id, display_no, input_files, params, update)
                output_path = result.get('raw_output_txt') if result else None
            elif task_type == 'business_licence':
                result = await self._execute_business_licence(task_id, display_no, input_files, params, update)
                output_path = result.get('output_docx') if result else None
            elif task_type == 'pdf2docx':
                result = await self._execute_pdf2docx(task_id, display_no, input_files, params, update)
                output_path = result.get('output_docx') if result else None
            else:
                raise ValueError(f'unsupported task type: {task_type}')
            output_files = self._extract_output_files(task_type, result, output_path, filename)
            with SessionLocal() as db:
                if task_repo.is_cancel_requested(db, task_id):
                    task_repo.mark_cancelled(db, task_id)
                    self._append_task_log(task_id, '[cancel] user cancelled before completion')
                    raise TaskCancelledError('user cancelled')
                task_repo.complete_task(db, task_id, result_json=json.dumps(result, ensure_ascii=False) if result is not None else None, output_path=output_path, output_files_json=json.dumps(output_files, ensure_ascii=False) if output_files else None)
            self._append_task_log(task_id, '[done] completed')
        except TaskCancelledError:
            self._append_task_log(task_id, '[cancel] completed cancel')
        except Exception as exc:
            self._append_task_log(task_id, f'[error] {type(exc).__name__}: {exc}')
            brief_tb = traceback.format_exc(limit=5)
            if brief_tb:
                self._append_task_log(task_id, brief_tb.rstrip())
            with SessionLocal() as db:
                task_repo.fail_task(db, task_id, str(exc))

    async def _execute_number_check(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'number check started')
        mode = params.get('mode', 'double')
        if mode == 'single':
            single_upload = UploadFile(
                filename=input_files.get('single_filename') or 'single.docx',
                file=io.BytesIO(Path(input_files['single_path']).read_bytes()),
            )
            job = asyncio.create_task(
                run_number_check_task(
                    single_file=single_upload,
                    mode=mode,
                    task_id=task_id,
                    display_no=display_no,
                    gemini_route=params.get('gemini_route', 'openrouter'),
                    model_name=params.get('model_name', 'gemini-3.1-pro-preview'),
                )
            )
        else:
            original_upload = UploadFile(filename=input_files.get('original_filename') or 'original.docx', file=io.BytesIO(Path(input_files['original_path']).read_bytes()))
            translated_upload = UploadFile(filename=input_files.get('translated_filename') or 'translated.docx', file=io.BytesIO(Path(input_files['translated_path']).read_bytes()))
            job = asyncio.create_task(
                run_number_check_task(
                    original_file=original_upload,
                    translated_file=translated_upload,
                    mode=mode,
                    task_id=task_id,
                    display_no=display_no,
                    gemini_route=params.get('gemini_route', 'openrouter'),
                    model_name=params.get('model_name', 'gemini-3.1-pro-preview'),
                )
            )
        await self._mirror_progress(task_id, job, lambda: get_number_check_progress(task_id), update)
        return await job

    async def _execute_zhongfanyi(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'zhongfanyi started')
        loop = asyncio.get_running_loop()
        job = loop.run_in_executor(
            self._task_executor,
            lambda: zf_service.run_zhongfanyi_task(
                task_id=task_id,
                display_no=display_no,
                mode=params.get('mode', zf_service.ZHONGFANYI_MODE_DOUBLE),
                original_path=input_files.get('original_path'),
                translated_path=input_files.get('translated_path'),
                single_path=input_files.get('single_path'),
                original_filename=input_files.get('original_filename'),
                translated_filename=input_files.get('translated_filename'),
                single_filename=input_files.get('single_filename'),
                use_ai_rule=params.get('use_ai_rule', False),
                gemini_route=params.get('gemini_route', 'openrouter'),
                model_name=params.get('model_name', zf_service.ZHONGFANYI_DEFAULT_MODEL),
                ai_rule_file_path=params.get('ai_rule_file_path'),
                session_rule_text=params.get('session_rule_text'),
            )
        )
        await self._mirror_progress(task_id, job, lambda: zf_service.get_task_progress(task_id), update)
        return await job

    async def _execute_alignment(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'alignment started')
        from app.service import alignment_service

        original_path = self._get_input_value(input_files, 'original_path', 'original')
        translated_path = self._get_input_value(input_files, 'translated_path', 'translated')
        if not original_path or not translated_path:
            raise ValueError('alignment task missing original/trans paths')

        job = asyncio.create_task(
            alignment_service.run_alignment_task(
                original_path=original_path,
                translated_path=translated_path,
                original_filename=input_files.get('original_filename'),
                translated_filename=input_files.get('translated_filename'),
                task_id=task_id,
                display_no=display_no,
                executor=self._task_executor,
                **params,
            )
        )
        await self._mirror_progress(task_id, job, lambda: alignment_service.get_alignment_progress(task_id), update)
        await job
        status = alignment_service.get_alignment_progress(task_id) or {}
        if status.get('stream_log'):
            self._set_task_log(task_id, status.get('stream_log', ''))

        final_status = status.get('status')
        if final_status == 'failed':
            raise RuntimeError(status.get('error') or status.get('message') or 'alignment failed')
        if final_status != 'done':
            raise RuntimeError(f"alignment finished with unexpected status: {final_status or 'unknown'}")

        result = status.get('result')
        if not result or not result.get('output_excel'):
            raise RuntimeError('alignment finished without output_excel')
        return result

    async def _execute_drivers_license(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'drivers license started')
        file_items = input_files.get('files') or []
        input_paths = [item['path'] for item in file_items if item.get('path')]
        original_filenames = [item.get('original_filename') or Path(item['path']).name for item in file_items if item.get('path')]
        return await execute_drivers_license_task(task_id=task_id, display_no=display_no, input_paths=input_paths, original_filenames=original_filenames, processing_mode=params.get('processing_mode', 'merge'), progress_callback=update, executor=self._task_executor)

    async def _execute_doc_translate(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'doc translate started')
        target_langs = [lang.strip() for lang in params.get('target_langs', 'en').split(',') if lang.strip()]
        return await execute_doc_translate_task(task_id=task_id, display_no=display_no, input_path=input_files['input_path'], original_filename=input_files.get('original_filename') or 'input.pdf', source_lang=params.get('source_lang', 'zh'), target_langs=target_langs, ocr_model=params.get('ocr_model', 'google/gemini-3-flash-preview'), gemini_route=params.get('gemini_route', 'openrouter'), progress_callback=update, executor=self._task_executor)

    async def _execute_business_licence(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'business licence started')
        return await execute_business_licence_task(
            task_id=task_id,
            display_no=display_no,
            input_path=input_files['input_path'],
            original_filename=input_files.get('original_filename') or 'business_licence.png',
            model=params.get('model', BUSINESS_LICENCE_DEFAULT_MODEL),
            gemini_route=params.get('gemini_route', BUSINESS_LICENCE_DEFAULT_ROUTE),
            parsed_data=params.get('parsed_data'),
            company_name_override=params.get('company_name_override'),
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_pdf2docx(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'pdf2docx started')
        return await execute_pdf2docx_task_from_path(task_id=task_id, display_no=display_no, input_path=input_files['input_path'], original_filename=input_files.get('original_filename') or 'input.pdf', model=params.get('model', PDF2DOCX_DEFAULT_MODEL), gemini_route=params.get('gemini_route', PDF2DOCX_DEFAULT_GEMINI_ROUTE), progress_callback=update, executor=self._task_executor)

    @staticmethod
    def _extract_output_files(task_type: str, result: Optional[Dict[str, Any]], output_path: Optional[str], original_filename: Optional[str] = None) -> list:
        def friendly(real_path: str, fallback_name: Optional[str] = None, fallback_suffix: Optional[str] = None) -> str:
            basename = Path(real_path or '').name
            if basename:
                return basename
            ext = Path(real_path or '').suffix or '.bin'
            return build_user_visible_filename(fallback_name, suffix=fallback_suffix, ext=ext)

        files = []
        if not result:
            if output_path:
                files.append({'name': friendly(output_path, original_filename, 'output'), 'path': output_path, 'type': 'output'})
            return files

        def add_result(key: str, ftype: str = 'output', fallback_name: Optional[str] = None, fallback_suffix: Optional[str] = None):
            value = result.get(key)
            if value and isinstance(value, str):
                files.append(
                    {
                        'name': friendly(value, fallback_name or original_filename, fallback_suffix or key),
                        'path': value,
                        'type': ftype,
                    }
                )

        if task_type == 'ocr':
            for item in result.get('results', []):
                if isinstance(item, dict):
                    if item.get('translated_image'):
                        files.append({'name': friendly(item['translated_image'], original_filename, 'translated'), 'path': item['translated_image'], 'type': 'output'})
                    if item.get('visualization_image'):
                        files.append({'name': friendly(item['visualization_image'], original_filename, 'visualization'), 'path': item['visualization_image'], 'type': 'output'})
        elif task_type == 'number_check':
            add_result('corrected_docx')
            reports = result.get('reports', {})
            if isinstance(reports, dict):
                for label, path in reports.items():
                    if isinstance(path, str) and path:
                        files.append({'name': friendly(path, original_filename, label), 'path': path, 'type': 'report'})
        elif task_type == 'zhongfanyi':
            add_result('corrected_docx')
            add_result('annotated_pdf')
            add_result('annotated_excel')
            add_result('annotated_pptx')
            add_result('output_file')
            reports = result.get('reports', {})
            if isinstance(reports, dict):
                for label, path in reports.items():
                    if isinstance(path, str) and path:
                        files.append({'name': friendly(path, original_filename, label), 'path': path, 'type': 'report'})
        elif task_type == 'alignment':
            add_result('output_excel')
        elif task_type == 'drivers_license':
            if result.get('processing_mode') == 'batch':
                for item in result.get('items', []):
                    if isinstance(item, dict) and item.get('output_docx'):
                        files.append({'name': friendly(item['output_docx'], item.get('input_filename'), 'translation'), 'path': item['output_docx'], 'type': 'output'})
            else:
                add_result('output_docx')
        elif task_type == 'doc_translate':
            add_result('raw_output_txt')
            translations = result.get('translations', {})
            values = translations.values() if isinstance(translations, dict) else translations
            for item in values:
                if isinstance(item, dict) and item.get('output_docx'):
                    files.append({'name': friendly(item['output_docx'], original_filename, item.get('lang_code') or 'translation'), 'path': item['output_docx'], 'type': 'output'})
        elif task_type == 'business_licence':
            add_result('output_docx')
        elif task_type == 'pdf2docx':
            add_result('output_docx')
        if not files and output_path:
            files.append({'name': friendly(output_path, original_filename, 'output'), 'path': output_path, 'type': 'output'})

        deduped_files = []
        seen = set()
        for item in files:
            key = (item.get('path'), item.get('type'))
            if key in seen:
                continue
            seen.add(key)
            deduped_files.append(item)
        return deduped_files

    async def _mirror_progress(self, task_id: str, job, getter: Callable[[], Optional[Dict[str, Any]]], update: Callable[[int, str], Any]):
        while not job.done():
            snapshot = getter()
            if snapshot:
                if snapshot.get('stream_log'):
                    self._set_task_log(task_id, snapshot.get('stream_log', ''))
                await update(snapshot.get('progress', 0), snapshot.get('message') or snapshot.get('status') or 'processing')
            await asyncio.sleep(1)


task_queue_service = TaskQueueService()
ocr_task_queue = task_queue_service

