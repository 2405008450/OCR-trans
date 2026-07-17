import asyncio
import hashlib
import io
import json
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from fastapi import UploadFile
from sqlalchemy.exc import IntegrityError

from app.core.config import settings
from app.core.file_naming import build_storage_filename, build_user_visible_filename
from app.core.request_context import get_client_ip
from app.db.session import SessionLocal
from app.repository import task_repo
from app.service import zhongfanyi_service as zf_service
from app.service.business_licence_service import (
    BUSINESS_LICENCE_DEFAULT_MODEL,
    BUSINESS_LICENCE_DEFAULT_ROUTE,
    execute_business_licence_task,
)
from app.service.doc_translate_service import DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE, execute_doc_translate_task
from app.service.english_variant_service import get_converter
from app.service.drivers_license_service import execute_drivers_license_task
from app.service.file_rename_service import (
    execute_file_rename_copy_task,
    prepare_file_rename_request,
)
from app.service.msg_convert_service import (
    MSG_CONVERT_DEFAULT_OUTPUT_FORMAT,
    execute_msg_convert_task,
    normalize_msg_output_format,
    validate_msg_file,
)
from app.service.office_text_transform_service import execute_english_variant_task
from app.service.number_check_service import (
    NUMBER_CHECK_MODE_ALIGNMENT,
    _get_task_progress as get_number_check_progress,
    run_number_check_task,
)
from app.service.pdf2docx_service import (
    PDF2DOCX_DEFAULT_LAYOUT_MODE,
    PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    PDF2DOCX_DEFAULT_MODEL,
    execute_pdf2docx_task_from_path,
)
from app.service.pdf_merge_service import execute_pdf_merge_task, prepare_pdf_merge_request
from app.service.pdf_tools_service import execute_pdf_tools_task, prepare_pdf_tools_request
from app.service.word_count_service import (
    execute_word_count_task,
    prepare_word_count_request,
    prepare_word_count_upload_request,
)


class TaskCancelledError(Exception):
    pass


class UploadSizeLimitError(ValueError):
    pass


@dataclass(frozen=True)
class TaskSubmitResult:
    task_id: str
    deduped: bool = False


@dataclass
class StagedUpload:
    role: str
    original_filename: str
    fallback_name: str
    temp_path: Path
    size: int
    sha256: str
    ext: str


class TaskQueueService:
    UPLOAD_CHUNK_SIZE = 1024 * 1024
    MAX_AUTO_REQUEUE_ATTEMPTS = 1
    DEFAULT_TASK_TYPE_LIMITS: Dict[str, int] = {
        'pdf2docx': 1,
        'doc_translate': 1,
        'alignment': 1,
        'drivers_license': 1,
        'business_licence': 2,
        'number_check': 2,
        'zhongfanyi': 2,
        'word_count': 1,
        'pdf_merge': 1,
        'pdf_tools': 1,
        'msg_convert': 1,
        'file_rename': 1,
        'english_variant': 1,
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

    @staticmethod
    def _normalize_for_fingerprint(value: Any) -> Any:
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, dict):
            return {
                str(key): TaskQueueService._normalize_for_fingerprint(item)
                for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
            }
        if isinstance(value, (list, tuple)):
            return [TaskQueueService._normalize_for_fingerprint(item) for item in value]
        return str(value)

    @classmethod
    def _build_file_fingerprints(cls, staged_uploads: list[StagedUpload]) -> list[Dict[str, Any]]:
        return [
            {
                'role': item.role,
                'filename': item.original_filename,
                'size': item.size,
                'sha256': item.sha256,
            }
            for item in staged_uploads
        ]

    @classmethod
    def build_request_fingerprint(cls, task_type: str, params: Dict[str, Any], file_fingerprints: list[Dict[str, Any]]) -> str:
        payload = {
            'task_type': task_type,
            'params': cls._normalize_for_fingerprint(params),
            'files': cls._normalize_for_fingerprint(file_fingerprints),
        }
        encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')
        return hashlib.sha256(encoded).hexdigest()

    async def _stage_upload(
        self,
        task_type: str,
        role: str,
        upload: UploadFile,
        fallback_name: str,
        max_bytes: Optional[int] = None,
    ) -> StagedUpload:
        original_filename = upload.filename or fallback_name
        ext = Path(original_filename).suffix or Path(fallback_name).suffix or '.bin'
        temp_dir = Path(settings.UPLOAD_DIR) / '_tmp_uploads' / task_type
        temp_dir.mkdir(parents=True, exist_ok=True)
        temp_path = temp_dir / f'{uuid.uuid4().hex}{ext}'
        digest = hashlib.sha256()
        size = 0
        try:
            with temp_path.open('wb') as target:
                while True:
                    chunk = await upload.read(self.UPLOAD_CHUNK_SIZE)
                    if not chunk:
                        break
                    size += len(chunk)
                    if max_bytes is not None and size > max_bytes:
                        limit_mb = max_bytes / (1024 * 1024)
                        raise UploadSizeLimitError(f"文件超过上传限制 {limit_mb:g} MB，请改用共享路径统计")
                    target.write(chunk)
                    digest.update(chunk)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise
        return StagedUpload(
            role=role,
            original_filename=original_filename,
            fallback_name=fallback_name,
            temp_path=temp_path,
            size=size,
            sha256=digest.hexdigest(),
            ext=ext,
        )

    async def _stage_uploads(
        self,
        task_type: str,
        uploads: list[tuple[str, Optional[UploadFile], str]],
        max_bytes: Optional[int] = None,
    ) -> list[StagedUpload]:
        staged_uploads: list[StagedUpload] = []
        try:
            for role, upload, fallback_name in uploads:
                if upload is None:
                    continue
                staged_uploads.append(
                    await self._stage_upload(task_type, role, upload, fallback_name, max_bytes=max_bytes)
                )
        except Exception:
            self._cleanup_staged_uploads(staged_uploads)
            raise
        return staged_uploads

    @staticmethod
    def _cleanup_staged_uploads(staged_uploads: list[StagedUpload]) -> None:
        for item in staged_uploads:
            item.temp_path.unlink(missing_ok=True)

    @staticmethod
    def _move_staged_upload(staged_upload: StagedUpload, upload_dir: Path, display_no: str, task_id: str) -> str:
        upload_dir.mkdir(parents=True, exist_ok=True)
        input_path = upload_dir / build_storage_filename(
            display_no,
            staged_upload.original_filename,
            task_id,
            role=staged_upload.role,
            ext=staged_upload.ext,
        )
        staged_upload.temp_path.replace(input_path)
        return str(input_path).replace('\\', '/')

    def _reserve_task_submission(
        self,
        *,
        task_type: str,
        filename: str,
        params: Dict[str, Any],
        staged_uploads: list[StagedUpload],
        batch_id: Optional[str] = None,
        batch_name: Optional[str] = None,
        batch_index: Optional[int] = None,
        batch_total: Optional[int] = None,
    ):
        file_fingerprints = self._build_file_fingerprints(staged_uploads)
        request_fingerprint = self.build_request_fingerprint(task_type, params, file_fingerprints)
        file_fingerprints_json = json.dumps(file_fingerprints, ensure_ascii=False)

        with SessionLocal() as db:
            existing_task = task_repo.get_active_task_by_request_fingerprint(db, request_fingerprint)
            if existing_task:
                self._append_task_log(existing_task.task_id, '[dedupe] duplicate submission reused this task')
                return TaskSubmitResult(existing_task.task_id, deduped=True), None

            task_id = str(uuid.uuid4())
            self._set_task_log(task_id, f'[queue] queued: {filename}')
            try:
                task = task_repo.create_task(
                    db,
                    task_id=task_id,
                    task_type=task_type,
                    filename=filename,
                    client_ip=get_client_ip(),
                    status='queued',
                    progress=0,
                    message='Queued',
                    params_json=json.dumps(params, ensure_ascii=False),
                    input_files_json='{}',
                    request_fingerprint=request_fingerprint,
                    file_fingerprints_json=file_fingerprints_json,
                    batch_id=batch_id,
                    batch_name=batch_name,
                    batch_index=batch_index,
                    batch_total=batch_total,
                )
            except IntegrityError:
                db.rollback()
                existing_task = task_repo.get_active_task_by_request_fingerprint(db, request_fingerprint)
                if existing_task:
                    self._append_task_log(existing_task.task_id, '[dedupe] duplicate submission reused this task')
                    return TaskSubmitResult(existing_task.task_id, deduped=True), None
                raise

        return TaskSubmitResult(task.task_id), task

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

    async def submit_number_check_task(
        self,
        *,
        mode: str,
        alignment_file: Optional[UploadFile],
        source_file: Optional[UploadFile],
        target_file: Optional[UploadFile],
        source_hf_file: Optional[UploadFile],
        gemini_route: str,
        model_name: str,
    ) -> TaskSubmitResult:
        if mode == NUMBER_CHECK_MODE_ALIGNMENT:
            alignment_name = alignment_file.filename if alignment_file else 'alignment.xlsx'
            target_name = target_file.filename if target_file else None
            display_name = f'{alignment_name} | {target_name}' if target_name else alignment_name
            if alignment_file is None:
                raise ValueError('alignment_file is required for alignment mode')
            upload_specs = [
                ('alignment', alignment_file, 'alignment.xlsx'),
                ('target', target_file, 'target.docx'),
                ('source_hf', source_hf_file, 'source_hf.docx'),
            ]
        else:
            source_name = source_file.filename if source_file else 'source.docx'
            target_name = target_file.filename if target_file else 'target.docx'
            display_name = f'{source_name} | {target_name}'
            if source_file is None or target_file is None:
                raise ValueError('source_file and target_file are required for direct mode')
            upload_specs = [
                ('source', source_file, 'source.docx'),
                ('target', target_file, 'target.docx'),
            ]

        params = {'mode': mode, 'gemini_route': gemini_route, 'model_name': model_name}
        staged_uploads = await self._stage_uploads('number_check', upload_specs)
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='number_check',
                filename=display_name,
                params=params,
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            upload_dir = Path(settings.UPLOAD_DIR) / 'number_check' / reserved_task.display_no
            input_files: Dict[str, Any] = {}
            for item in staged_uploads:
                input_files[f'{item.role}_path'] = self._move_staged_upload(item, upload_dir, reserved_task.display_no, reserved_task.task_id)
                input_files[f'{item.role}_filename'] = item.original_filename

            self._update_task_input_files(reserved_task.task_id, input_files)
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_zhongfanyi_task(self, *, mode: str, original_file: Optional[UploadFile], translated_file: Optional[UploadFile], single_file: Optional[UploadFile], use_ai_rule: bool, gemini_route: str, model_name: str, rule_file: Optional[UploadFile], session_rule_content: Optional[str]) -> TaskSubmitResult:
        session_rule_text = (session_rule_content.strip() or None) if session_rule_content else None
        if mode == zf_service.ZHONGFANYI_MODE_SINGLE:
            display_name = single_file.filename if single_file else 'single.docx'
            if single_file is None:
                raise ValueError('single_file is required for single mode')
            upload_specs = [('single', single_file, 'single.docx')]
        else:
            original_name = original_file.filename if original_file else 'original.docx'
            translated_name = translated_file.filename if translated_file else 'translated.docx'
            display_name = f'{original_name} | {translated_name}'
            if original_file is None or translated_file is None:
                raise ValueError('original_file and translated_file are required for double mode')
            upload_specs = [('original', original_file, 'original.docx'), ('translated', translated_file, 'translated.docx')]
        if rule_file and use_ai_rule:
            upload_specs.append(('rule', rule_file, 'rule.txt'))

        params = {
            'mode': mode,
            'use_ai_rule': use_ai_rule,
            'gemini_route': gemini_route,
            'model_name': model_name,
            'ai_rule_file_path': None,
            'session_rule_text': session_rule_text,
        }
        staged_uploads = await self._stage_uploads('zhongfanyi', upload_specs)
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='zhongfanyi',
                filename=display_name,
                params=params,
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            upload_dir = Path(settings.UPLOAD_DIR) / 'zhongfanyi' / reserved_task.display_no
            input_files = {}
            ai_rule_file_path = None
            for item in staged_uploads:
                moved_path = self._move_staged_upload(item, upload_dir, reserved_task.display_no, reserved_task.task_id)
                if item.role == 'rule':
                    ai_rule_file_path = moved_path
                else:
                    input_files[f'{item.role}_path'] = moved_path
                    input_files[f'{item.role}_filename'] = item.original_filename

            final_params = dict(params)
            final_params['ai_rule_file_path'] = ai_rule_file_path
            self._update_task_params(reserved_task.task_id, final_params)
            self._update_task_input_files(reserved_task.task_id, input_files)
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_alignment_task(self, *, original_file: UploadFile, translated_file: UploadFile, source_lang: str, target_lang: str, model_name: str, gemini_route: str, enable_post_split: bool, threshold_2: int, threshold_3: int, threshold_4: int, threshold_5: int, threshold_6: int, threshold_7: int, threshold_8: int, buffer_chars: int) -> TaskSubmitResult:
        display_name = f'{original_file.filename} | {translated_file.filename}'
        params = {'source_lang': source_lang, 'target_lang': target_lang, 'model_name': model_name, 'gemini_route': gemini_route, 'enable_post_split': enable_post_split, 'threshold_2': threshold_2, 'threshold_3': threshold_3, 'threshold_4': threshold_4, 'threshold_5': threshold_5, 'threshold_6': threshold_6, 'threshold_7': threshold_7, 'threshold_8': threshold_8, 'buffer_chars': buffer_chars}
        staged_uploads = await self._stage_uploads(
            'alignment',
            [('original', original_file, 'original.docx'), ('translated', translated_file, 'translated.docx')],
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='alignment',
                filename=display_name,
                params=params,
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            upload_dir = Path(settings.UPLOAD_DIR) / 'alignment' / reserved_task.display_no
            input_files = {}
            for item in staged_uploads:
                input_files[f'{item.role}_path'] = self._move_staged_upload(item, upload_dir, reserved_task.display_no, reserved_task.task_id)
                input_files[f'{item.role}_filename'] = item.original_filename
            self._update_task_input_files(reserved_task.task_id, input_files)
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_doc_translate_task(self, *, file: UploadFile, source_lang: str, target_langs: str, translate_mode: str, ocr_model: str, gemini_route: str, translation_engine: str, translation_rules: str = "", batch_id: Optional[str] = None, batch_name: Optional[str] = None, batch_index: Optional[int] = None, batch_total: Optional[int] = None) -> TaskSubmitResult:
        params = {'source_lang': source_lang, 'target_langs': target_langs, 'translate_mode': translate_mode, 'ocr_model': ocr_model, 'gemini_route': gemini_route, 'translation_engine': translation_engine, 'translation_rules': translation_rules}
        staged_uploads = await self._stage_uploads('doc_translate', [('input', file, 'input.bin')])
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='doc_translate',
                filename=file.filename or 'input.bin',
                params=params,
                staged_uploads=staged_uploads,
                batch_id=batch_id,
                batch_name=batch_name,
                batch_index=batch_index,
                batch_total=batch_total,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            item = staged_uploads[0]
            input_path = self._move_staged_upload(item, Path(settings.UPLOAD_DIR) / 'doc_translate' / reserved_task.display_no, reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': item.original_filename})
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_english_variant_task(
        self,
        *,
        file: UploadFile,
        target_style: str,
        batch_id: Optional[str] = None,
        batch_name: Optional[str] = None,
        batch_index: Optional[int] = None,
        batch_total: Optional[int] = None,
    ) -> TaskSubmitResult:
        converter = get_converter()
        params = {
            'target_style': target_style,
            'dictionary_version': converter.dictionary_version,
            'dictionary_sha256': converter.source_sha256,
        }
        max_bytes = max(1, int(settings.ENGLISH_VARIANT_UPLOAD_MAX_MB or 95)) * 1024 * 1024
        staged_uploads = await self._stage_uploads(
            'english_variant',
            [('input', file, 'input.docx')],
            max_bytes=max_bytes,
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='english_variant',
                filename=file.filename or 'input.docx',
                params=params,
                staged_uploads=staged_uploads,
                batch_id=batch_id,
                batch_name=batch_name,
                batch_index=batch_index,
                batch_total=batch_total,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            item = staged_uploads[0]
            input_path = self._move_staged_upload(
                item,
                Path(settings.UPLOAD_DIR) / 'english_variant' / reserved_task.display_no,
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {'input_path': input_path, 'original_filename': item.original_filename},
            )
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
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
    ) -> TaskSubmitResult:
        params = {
            'model': model,
            'gemini_route': gemini_route,
            'parsed_data': parsed_data,
            'company_name_override': company_name_override,
        }
        staged_uploads = await self._stage_uploads('business_licence', [('input', file, 'input.bin')])
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='business_licence',
                filename=file.filename or 'input.bin',
                params=params,
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            item = staged_uploads[0]
            input_path = self._move_staged_upload(item, Path(settings.UPLOAD_DIR) / 'business_licence' / reserved_task.display_no, reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': item.original_filename})
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_pdf2docx_task(self, *, file: UploadFile, model: str, gemini_route: str, layout_mode: str = PDF2DOCX_DEFAULT_LAYOUT_MODE, batch_id: Optional[str] = None, batch_name: Optional[str] = None, batch_index: Optional[int] = None, batch_total: Optional[int] = None) -> TaskSubmitResult:
        params = {'model': model, 'gemini_route': gemini_route, 'layout_mode': layout_mode}
        staged_uploads = await self._stage_uploads('pdf2docx', [('input', file, 'input.bin')])
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='pdf2docx',
                filename=file.filename or 'input.bin',
                params=params,
                staged_uploads=staged_uploads,
                batch_id=batch_id,
                batch_name=batch_name,
                batch_index=batch_index,
                batch_total=batch_total,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            item = staged_uploads[0]
            input_path = self._move_staged_upload(item, Path(settings.UPLOAD_DIR) / 'pdf2docx' / reserved_task.display_no, reserved_task.display_no, reserved_task.task_id)
            self._update_task_input_files(reserved_task.task_id, {'input_path': input_path, 'original_filename': item.original_filename})
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_msg_convert_task(
        self,
        *,
        file: UploadFile,
        output_format: str = MSG_CONVERT_DEFAULT_OUTPUT_FORMAT,
        batch_id: Optional[str] = None,
        batch_name: Optional[str] = None,
        batch_index: Optional[int] = None,
        batch_total: Optional[int] = None,
    ) -> TaskSubmitResult:
        normalized_format = normalize_msg_output_format(output_format)
        max_bytes = max(1, int(settings.MSG_CONVERT_UPLOAD_MAX_MB or 95)) * 1024 * 1024
        try:
            staged_uploads = await self._stage_uploads(
                'msg_convert',
                [('input', file, 'input.msg')],
                max_bytes=max_bytes,
            )
        except UploadSizeLimitError as exc:
            raise UploadSizeLimitError(
                f"MSG 文件超过 {settings.MSG_CONVERT_UPLOAD_MAX_MB:g} MB 上传限制"
            ) from exc

        reserved_task = None
        try:
            item = staged_uploads[0]
            validate_msg_file(item.temp_path, item.original_filename, max_bytes=max_bytes)
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='msg_convert',
                filename=item.original_filename,
                params={'output_format': normalized_format},
                staged_uploads=staged_uploads,
                batch_id=batch_id,
                batch_name=batch_name,
                batch_index=batch_index,
                batch_total=batch_total,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            input_path = self._move_staged_upload(
                item,
                Path(settings.UPLOAD_DIR) / 'msg_convert' / reserved_task.display_no,
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {'input_path': input_path, 'original_filename': item.original_filename},
            )
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_word_count_task(
        self,
        *,
        directory_path: str,
        recursive: bool = True,
        include_hidden: bool = False,
        extensions: Optional[list[str]] = None,
        ocr_mode: str = 'auto',
        ocr_model: Optional[str] = None,
        relative_paths: Optional[list[str]] = None,
    ) -> TaskSubmitResult:
        prepared = prepare_word_count_request(
            directory_path=directory_path,
            recursive=recursive,
            include_hidden=include_hidden,
            extensions=extensions,
            ocr_mode=ocr_mode,
            ocr_model=ocr_model,
            relative_paths=relative_paths,
        )
        params = prepared['params']
        input_files = prepared['input_files']
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='word_count',
                filename=prepared['filename'],
                params=params,
                staged_uploads=[],
            )
            if submit_result.deduped:
                return submit_result
            self._update_task_input_files(reserved_task.task_id, input_files)
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_word_count_upload_task(
        self,
        *,
        file: UploadFile,
        ocr_mode: str = 'auto',
        ocr_model: Optional[str] = None,
    ) -> TaskSubmitResult:
        prepared = prepare_word_count_upload_request(
            filename=file.filename or '',
            ocr_mode=ocr_mode,
            ocr_model=ocr_model,
        )
        max_bytes = max(1, int(settings.WORD_COUNT_UPLOAD_MAX_MB or 50)) * 1024 * 1024
        staged_uploads = await self._stage_uploads(
            'word_count',
            [('input', file, f"input{prepared['extension']}")],
            max_bytes=max_bytes,
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='word_count',
                filename=prepared['filename'],
                params=prepared['params'],
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            item = staged_uploads[0]
            input_path = self._move_staged_upload(
                item,
                Path(settings.UPLOAD_DIR) / 'word_count' / reserved_task.display_no,
                reserved_task.display_no,
                reserved_task.task_id,
            )
            self._update_task_input_files(
                reserved_task.task_id,
                {
                    'directory_path': input_path,
                    'input_kind': 'file',
                    'input_source': 'upload',
                    'original_filename': item.original_filename,
                },
            )
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_pdf_merge_task(
        self,
        *,
        directory_path: str,
        relative_paths: list[str],
        output_filename: str,
    ) -> TaskSubmitResult:
        prepared = prepare_pdf_merge_request(
            directory_path=directory_path,
            relative_paths=relative_paths,
            output_filename=output_filename,
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='pdf_merge',
                filename=prepared['filename'],
                params=prepared['params'],
                staged_uploads=[],
            )
            if submit_result.deduped:
                return submit_result
            self._update_task_input_files(reserved_task.task_id, prepared['input_files'])
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_file_rename_task(
        self,
        *,
        directory_path: str,
        relative_paths: list[str],
        mode: str,
        recursive: bool = True,
        include_hidden: bool = False,
        regex_pattern: str = '',
        replacement: str = '',
        ignore_case: bool = False,
        cleanup_remove_leading_number: bool = True,
        cleanup_leading_number_max_digits: int = 6,
        cleanup_leading_number_space: bool = True,
        cleanup_leading_number_underscore: bool = True,
        cleanup_remove_datetime: bool = True,
        cleanup_datetime_compact: bool = True,
        cleanup_datetime_dotted: bool = True,
        cleanup_remove_translated: bool = True,
        cleanup_translated_suffix: str = '_translated',
    ) -> TaskSubmitResult:
        prepared = prepare_file_rename_request(
            directory_path=directory_path,
            relative_paths=relative_paths,
            mode=mode,
            recursive=recursive,
            include_hidden=include_hidden,
            regex_pattern=regex_pattern,
            replacement=replacement,
            ignore_case=ignore_case,
            cleanup_remove_leading_number=cleanup_remove_leading_number,
            cleanup_leading_number_max_digits=cleanup_leading_number_max_digits,
            cleanup_leading_number_space=cleanup_leading_number_space,
            cleanup_leading_number_underscore=cleanup_leading_number_underscore,
            cleanup_remove_datetime=cleanup_remove_datetime,
            cleanup_datetime_compact=cleanup_datetime_compact,
            cleanup_datetime_dotted=cleanup_datetime_dotted,
            cleanup_remove_translated=cleanup_remove_translated,
            cleanup_translated_suffix=cleanup_translated_suffix,
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='file_rename',
                filename=prepared['filename'],
                params=prepared['params'],
                staged_uploads=[],
            )
            if submit_result.deduped:
                return submit_result
            self._update_task_input_files(reserved_task.task_id, prepared['input_files'])
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_pdf_tools_task(
        self,
        *,
        directory_path: str,
        relative_path: str,
        operation: str,
        options: Optional[dict[str, Any]] = None,
    ) -> TaskSubmitResult:
        prepared = prepare_pdf_tools_request(
            directory_path=directory_path,
            relative_path=relative_path,
            operation=operation,
            options=options or {},
        )
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='pdf_tools',
                filename=prepared['filename'],
                params=prepared['params'],
                staged_uploads=[],
            )
            if submit_result.deduped:
                return submit_result
            self._update_task_input_files(reserved_task.task_id, prepared['input_files'])
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            if reserved_task is not None:
                self._fail_reserved_task(reserved_task.task_id, exc)
            raise

    async def submit_drivers_license_task(self, *, files: list[UploadFile], processing_mode: str) -> TaskSubmitResult:
        display_name = ' | '.join([(file.filename or 'input.bin') for file in files])
        upload_specs = [(f'image{index:02d}', file, f'image_{index}.bin') for index, file in enumerate(files, start=1)]
        staged_uploads = await self._stage_uploads('drivers_license', upload_specs)
        reserved_task = None
        try:
            submit_result, reserved_task = self._reserve_task_submission(
                task_type='drivers_license',
                filename=display_name,
                params={'processing_mode': processing_mode},
                staged_uploads=staged_uploads,
            )
            if submit_result.deduped:
                self._cleanup_staged_uploads(staged_uploads)
                return submit_result

            upload_dir = Path(settings.UPLOAD_DIR) / 'drivers_license' / reserved_task.display_no
            saved_files = []
            for index, item in enumerate(staged_uploads, start=1):
                input_path = self._move_staged_upload(item, upload_dir, reserved_task.display_no, reserved_task.task_id)
                saved_files.append({'index': index, 'path': input_path, 'original_filename': item.original_filename})
            self._update_task_input_files(reserved_task.task_id, {'files': saved_files})
            self._notify_dispatcher()
            return submit_result
        except Exception as exc:
            self._cleanup_staged_uploads(staged_uploads)
            if reserved_task is not None:
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
            return task_repo.create_task(db, task_id=task_id, task_type=task_type, filename=filename, client_ip=get_client_ip(), status='queued', progress=0, message='Queued', params_json=json.dumps(params, ensure_ascii=False), input_files_json=json.dumps(input_files, ensure_ascii=False))

    def _update_task_input_files(self, task_id: str, input_files: Dict[str, Any]) -> None:
        with SessionLocal() as db:
            task_repo.update_task_input_files(db, task_id, json.dumps(input_files, ensure_ascii=False))

    def _update_task_params(self, task_id: str, params: Dict[str, Any]) -> None:
        with SessionLocal() as db:
            task_repo.update_task_params(db, task_id, json.dumps(params, ensure_ascii=False))

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
        if task_type in {'doc_translate', 'business_licence', 'pdf2docx', 'msg_convert', 'english_variant'}:
            return [] if self._get_input_value(input_files, 'input_path') else ['input_path']

        if task_type == 'word_count':
            return [] if self._get_input_value(input_files, 'directory_path') else ['directory_path']

        if task_type == 'pdf_merge':
            missing = []
            if not self._get_input_value(input_files, 'directory_path'):
                missing.append('directory_path')
            if not (input_files.get('relative_paths') or []):
                missing.append('relative_paths')
            return missing

        if task_type == 'file_rename':
            missing = []
            if not self._get_input_value(input_files, 'directory_path'):
                missing.append('directory_path')
            if not (input_files.get('relative_paths') or []):
                missing.append('relative_paths')
            return missing

        if task_type == 'pdf_tools':
            missing = []
            if not self._get_input_value(input_files, 'directory_path'):
                missing.append('directory_path')
            if not self._get_input_value(input_files, 'relative_path'):
                missing.append('relative_path')
            return missing

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
            mode = params.get('mode', NUMBER_CHECK_MODE_ALIGNMENT)
            if mode in {NUMBER_CHECK_MODE_ALIGNMENT, 'single', 'excel', 'alignment_excel'}:
                return [] if self._get_input_value(input_files, 'alignment_path', 'single_path') else ['alignment_path']
            missing = []
            if not self._get_input_value(input_files, 'source_path', 'original_path'):
                missing.append('source_path')
            if not self._get_input_value(input_files, 'target_path', 'translated_path'):
                missing.append('target_path')
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
            input_files = await self._ensure_task_input_files(task_id, task_type, params, input_files)
            if task_type == 'number_check':
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
            elif task_type == 'word_count':
                result = await self._execute_word_count(task_id, display_no, input_files, params, update)
                output_path = result.get('report_excel') if result else None
            elif task_type == 'pdf_merge':
                result = await self._execute_pdf_merge(task_id, display_no, input_files, params, update)
                output_path = result.get('output_pdf') if result else None
            elif task_type == 'file_rename':
                result = await self._execute_file_rename(task_id, display_no, input_files, params, update)
                output_path = None
            elif task_type == 'pdf_tools':
                result = await self._execute_pdf_tools(task_id, display_no, input_files, params, update)
                output_path = (result.get('output_pdf') or result.get('archive_zip')) if result else None
            elif task_type == 'msg_convert':
                result = await self._execute_msg_convert(task_id, display_no, input_files, params, update)
                output_path = (result.get('output_docx') or result.get('output_pdf')) if result else None
            elif task_type == 'english_variant':
                result = await self._execute_english_variant(task_id, display_no, input_files, params, update)
                output_path = result.get('output_file') if result else None
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
        mode = params.get('mode', NUMBER_CHECK_MODE_ALIGNMENT)

        def make_upload(role: str, fallback_name: str, *legacy_roles: str) -> Optional[UploadFile]:
            roles = (role, *legacy_roles)
            selected_role = next((candidate for candidate in roles if input_files.get(f'{candidate}_path')), None)
            if not selected_role:
                return None
            path_key = f'{selected_role}_path'
            return UploadFile(
                filename=input_files.get(f'{selected_role}_filename') or fallback_name,
                file=io.BytesIO(Path(input_files[path_key]).read_bytes()),
            )

        job = asyncio.create_task(
            run_number_check_task(
                alignment_file=make_upload('alignment', 'alignment.xlsx', 'single'),
                source_file=make_upload('source', 'source.docx', 'original'),
                target_file=make_upload('target', 'target.docx', 'translated'),
                source_hf_file=make_upload('source_hf', 'source_hf.docx'),
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
        return await execute_doc_translate_task(task_id=task_id, display_no=display_no, input_path=input_files['input_path'], original_filename=input_files.get('original_filename') or 'input.pdf', source_lang=params.get('source_lang', 'zh'), target_langs=target_langs, translate_mode=params.get('translate_mode', 'standard'), ocr_model=params.get('ocr_model', 'google/gemini-3-flash-preview'), gemini_route=params.get('gemini_route', 'openrouter'), translation_engine=params.get('translation_engine', DOC_TRANSLATE_DEFAULT_TRANSLATION_ENGINE), translation_rules=params.get('translation_rules', ''), progress_callback=update, executor=self._task_executor)

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

    async def _execute_english_variant(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'english variant conversion started')
        return await execute_english_variant_task(
            task_id=task_id,
            display_no=display_no,
            input_path=input_files['input_path'],
            original_filename=input_files.get('original_filename') or 'input.docx',
            target_style=params.get('target_style', 'british'),
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_pdf2docx(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'pdf2docx started')
        return await execute_pdf2docx_task_from_path(task_id=task_id, display_no=display_no, input_path=input_files['input_path'], original_filename=input_files.get('original_filename') or 'input.pdf', model=params.get('model', PDF2DOCX_DEFAULT_MODEL), gemini_route=params.get('gemini_route', PDF2DOCX_DEFAULT_GEMINI_ROUTE), layout_mode=params.get('layout_mode', PDF2DOCX_DEFAULT_LAYOUT_MODE), progress_callback=update, executor=self._task_executor)

    async def _execute_word_count(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        await update(5, 'word count started')
        return await execute_word_count_task(
            task_id=task_id,
            display_no=display_no,
            directory_path=input_files['directory_path'],
            recursive=bool(params.get('recursive', True)),
            include_hidden=bool(params.get('include_hidden', False)),
            extensions=params.get('extensions') or None,
            ocr_mode=params.get('ocr_mode', 'auto'),
            ocr_model=params.get('ocr_model') or PDF2DOCX_DEFAULT_MODEL,
            ocr_route=params.get('ocr_route') or PDF2DOCX_DEFAULT_GEMINI_ROUTE,
            input_source=params.get('input_source', 'path'),
            original_filename=input_files.get('original_filename'),
            relative_paths=input_files.get('relative_paths'),
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_pdf_merge(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        return await execute_pdf_merge_task(
            task_id=task_id,
            display_no=display_no,
            directory_path=input_files['directory_path'],
            relative_paths=input_files.get('relative_paths') or [],
            output_filename=params.get('output_filename') or '合并结果.pdf',
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_file_rename(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        return await execute_file_rename_copy_task(
            task_id=task_id,
            display_no=display_no,
            directory_path=input_files['directory_path'],
            relative_paths=input_files.get('relative_paths') or [],
            mode=params.get('mode') or 'numbering',
            recursive=bool(params.get('recursive', True)),
            include_hidden=bool(params.get('include_hidden', False)),
            regex_pattern=params.get('regex_pattern') or '',
            replacement=params.get('replacement') or '',
            ignore_case=bool(params.get('ignore_case', False)),
            cleanup_remove_leading_number=bool(params.get('cleanup_remove_leading_number', True)),
            cleanup_leading_number_max_digits=int(params.get('cleanup_leading_number_max_digits', 6)),
            cleanup_leading_number_space=bool(params.get('cleanup_leading_number_space', True)),
            cleanup_leading_number_underscore=bool(params.get('cleanup_leading_number_underscore', True)),
            cleanup_remove_datetime=bool(params.get('cleanup_remove_datetime', True)),
            cleanup_datetime_compact=bool(params.get('cleanup_datetime_compact', True)),
            cleanup_datetime_dotted=bool(params.get('cleanup_datetime_dotted', True)),
            cleanup_remove_translated=bool(params.get('cleanup_remove_translated', True)),
            cleanup_translated_suffix=str(params.get('cleanup_translated_suffix', '_translated')),
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_pdf_tools(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        return await execute_pdf_tools_task(
            task_id=task_id,
            display_no=display_no,
            directory_path=input_files['directory_path'],
            relative_path=input_files['relative_path'],
            operation=params.get('operation') or '',
            options=params.get('options') or {},
            progress_callback=update,
            executor=self._task_executor,
        )

    async def _execute_msg_convert(self, task_id: str, display_no: str, input_files: Dict[str, Any], params: Dict[str, Any], update: Callable[[int, str], Any]) -> Dict[str, Any]:
        return await execute_msg_convert_task(
            task_id=task_id,
            display_no=display_no,
            input_path=input_files['input_path'],
            original_filename=input_files.get('original_filename') or 'input.msg',
            output_format=params.get('output_format', MSG_CONVERT_DEFAULT_OUTPUT_FORMAT),
            progress_callback=update,
            executor=self._task_executor,
        )

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

        if task_type == 'number_check':
            add_result('corrected_docx')
            add_result('revised_file')
            add_result('report_excel', ftype='report')
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
        elif task_type == 'word_count':
            add_result('report_excel', ftype='report')
            add_result('report_json', ftype='report')
            add_result('ocr_text_archive', ftype='report')
        elif task_type == 'pdf_merge':
            add_result('output_pdf')
        elif task_type == 'pdf_tools':
            add_result('output_pdf')
            add_result('archive_zip')
            for item in result.get('output_files', []):
                if isinstance(item, dict) and isinstance(item.get('path'), str):
                    files.append(
                        {
                            'name': item.get('filename') or friendly(item['path'], original_filename, 'part'),
                            'path': item['path'],
                            'type': 'output',
                        }
                    )
        elif task_type == 'msg_convert':
            add_result('output_docx')
            add_result('output_pdf')
        elif task_type == 'english_variant':
            add_result('output_file')
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
