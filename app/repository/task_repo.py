from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from sqlalchemy import func
from sqlalchemy.orm import Session

from app.core.file_naming import build_display_no
from app.model.entity import Task

TASK_TYPE_LABELS = {
    'ocr': 'OCR',
    'number_check': 'Number Check',
    'zhongfanyi': 'Zhongfanyi',
    'alignment': 'Alignment',
    'drivers_license': 'Drivers License',
    'doc_translate': 'Doc Translate',
    'business_licence': '证件翻译（营业执照）',
    'pdf2docx': 'PDF2DOCX',
}


def _now() -> datetime:
    return datetime.utcnow()


def create_task(db: Session, *, task_id: str, task_type: str, filename: str, status: str = 'queued', progress: int = 0, message: Optional[str] = None, params_json: Optional[str] = None, input_files_json: Optional[str] = None) -> Task:
    now = _now()
    task = Task(task_id=task_id, task_type=task_type, task_label=TASK_TYPE_LABELS.get(task_type, task_type), filename=filename, status=status, progress=progress, message=message, params_json=params_json, input_files_json=input_files_json, created_at=now, updated_at=now)
    db.add(task)
    db.commit()
    db.refresh(task)
    if not task.display_no:
        task.display_no = build_display_no(task.id)
        db.commit()
        db.refresh(task)
    return task


def get_task_by_task_id(db: Session, task_id: str) -> Optional[Task]:
    return db.query(Task).filter(Task.task_id == task_id).first()


def update_task_input_files(db: Session, task_id: str, input_files_json: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    task.input_files_json = input_files_json
    task.updated_at = _now()
    db.commit()
    db.refresh(task)
    return task


def count_tasks_ahead(db: Session, task: Task) -> int:
    return db.query(Task).filter(Task.status == 'queued', Task.created_at < task.created_at).count()


def list_queued_tasks(db: Session, *, limit: int = 20, task_type: Optional[str] = None) -> List[Task]:
    query = db.query(Task).filter(Task.status == 'queued', Task.cancel_requested.is_(False))
    if task_type:
        query = query.filter(Task.task_type == task_type)
    return query.order_by(Task.created_at.asc(), Task.id.asc()).limit(limit).all()


def claim_queued_task_by_task_id(db: Session, task_id: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task or task.status != 'queued' or task.cancel_requested:
        return None

    now = _now()
    updated_rows = db.query(Task).filter(
        Task.id == task.id,
        Task.status == 'queued',
        Task.cancel_requested.is_(False),
    ).update(
        {
            Task.status: 'running',
            Task.progress: 1,
            Task.message: 'Processing',
            Task.started_at: now,
            Task.updated_at: now,
            Task.error_message: None,
        },
        synchronize_session=False,
    )
    if updated_rows != 1:
        db.rollback()
        return None
    db.commit()
    return get_task_by_task_id(db, task.task_id)


def claim_next_queued_task(db: Session, task_type: Optional[str] = None) -> Optional[Task]:
    tasks = list_queued_tasks(db, limit=1, task_type=task_type)
    task = tasks[0] if tasks else None
    if not task:
        return None
    return claim_queued_task_by_task_id(db, task.task_id)


def update_task_progress(db: Session, task_id: str, *, progress: Optional[int] = None, message: Optional[str] = None, status: Optional[str] = None) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    if progress is not None:
        task.progress = progress
    if message is not None:
        task.message = message
    if status is not None:
        task.status = status
    task.updated_at = _now()
    db.commit()
    db.refresh(task)
    return task


def complete_task(db: Session, task_id: str, *, result_json: Optional[str] = None, output_path: Optional[str] = None, output_files_json: Optional[str] = None, message: str = 'Completed') -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    now = _now()
    task.status = 'done'
    task.progress = 100
    task.message = message
    task.result_json = result_json
    task.output_path = output_path
    task.output_files_json = output_files_json
    task.error_message = None
    task.finished_at = now
    task.updated_at = now
    db.commit()
    db.refresh(task)
    return task


def fail_task(db: Session, task_id: str, error_message: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    now = _now()
    task.status = 'failed'
    task.message = f'Failed: {error_message}'
    task.error_message = error_message
    task.finished_at = now
    task.updated_at = now
    db.commit()
    db.refresh(task)
    return task


def cancel_task(db: Session, task_id: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    if task.status in ('done', 'failed', 'cancelled'):
        return task
    now = _now()
    task.cancel_requested = True
    if task.status == 'queued':
        task.status = 'cancelled'
        task.message = 'Cancelled'
        task.finished_at = now
    else:
        task.message = 'Cancelling'
    task.updated_at = now
    db.commit()
    db.refresh(task)
    return task


def is_cancel_requested(db: Session, task_id: str) -> bool:
    task = get_task_by_task_id(db, task_id)
    return bool(task and task.cancel_requested)


def mark_cancelled(db: Session, task_id: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None
    now = _now()
    task.status = 'cancelled'
    task.message = 'Cancelled by user'
    task.finished_at = now
    task.updated_at = now
    db.commit()
    db.refresh(task)
    return task


def requeue_running_tasks(db: Session, *, task_type: Optional[str] = None, max_retry_count: int = 1) -> Dict[str, int]:
    query = db.query(Task).filter(Task.status == 'running')
    if task_type:
        query = query.filter(Task.task_type == task_type)
    tasks = query.order_by(Task.created_at.asc(), Task.id.asc()).all()
    summary = {'requeued': 0, 'cancelled': 0, 'failed': 0}
    if not tasks:
        return summary

    now = _now()
    for task in tasks:
        task.started_at = None
        task.updated_at = now

        if task.cancel_requested:
            task.status = 'cancelled'
            task.message = 'Cancelled during restart recovery'
            task.finished_at = now
            summary['cancelled'] += 1
            continue

        if task.retry_count >= max_retry_count:
            task.status = 'failed'
            task.message = 'Failed: auto requeue limit reached after restart'
            task.error_message = 'auto requeue limit reached after restart'
            task.finished_at = now
            summary['failed'] += 1
            continue

        task.status = 'queued'
        task.progress = 0
        task.message = f'Requeued after restart ({task.retry_count + 1}/{max_retry_count})'
        task.finished_at = None
        task.error_message = None
        task.retry_count += 1
        summary['requeued'] += 1

    db.commit()
    return summary


def list_tasks(db: Session, *, status: Optional[str] = None, task_type: Optional[str] = None, keyword: Optional[str] = None, page: int = 1, page_size: int = 20) -> Tuple[List[Task], int]:
    query = db.query(Task)
    if status:
        statuses = [s.strip() for s in status.split(',') if s.strip()]
        if statuses:
            query = query.filter(Task.status.in_(statuses))
    if task_type:
        query = query.filter(Task.task_type == task_type)
    if keyword:
        query = query.filter(Task.filename.ilike(f'%{keyword}%'))
    total = query.count()
    tasks = query.order_by(Task.created_at.desc(), Task.id.desc()).offset((page - 1) * page_size).limit(page_size).all()
    return tasks, total


def count_by_status(db: Session) -> Dict[str, Any]:
    rows = db.query(Task.status, func.count(Task.id)).group_by(Task.status).all()
    counts = {row[0]: row[1] for row in rows}
    type_rows = db.query(Task.task_type, func.count(Task.id)).group_by(Task.task_type).all()
    by_type = {row[0]: row[1] for row in type_rows}
    return {'total': sum(counts.values()), 'queued': counts.get('queued', 0), 'running': counts.get('running', 0), 'done': counts.get('done', 0), 'failed': counts.get('failed', 0), 'cancelled': counts.get('cancelled', 0), 'by_type': by_type}
