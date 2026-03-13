from datetime import datetime
from typing import Optional

from sqlalchemy.orm import Session

from app.model.entity import Task


def create_task(
    db: Session,
    *,
    task_id: str,
    task_type: str,
    filename: str,
    status: str = "queued",
    progress: int = 0,
    message: Optional[str] = None,
    params_json: Optional[str] = None,
    input_files_json: Optional[str] = None,
) -> Task:
    task = Task(
        task_id=task_id,
        task_type=task_type,
        filename=filename,
        status=status,
        progress=progress,
        message=message,
        params_json=params_json,
        input_files_json=input_files_json,
    )
    db.add(task)
    db.commit()
    db.refresh(task)
    return task


def get_task_by_task_id(db: Session, task_id: str) -> Optional[Task]:
    return db.query(Task).filter(Task.task_id == task_id).first()


def count_tasks_ahead(db: Session, task: Task) -> int:
    return (
        db.query(Task)
        .filter(Task.status == "queued", Task.created_at < task.created_at)
        .count()
    )


def claim_next_queued_task(db: Session, task_type: Optional[str] = None) -> Optional[Task]:
    query = db.query(Task).filter(Task.status == "queued", Task.cancel_requested.is_(False))
    if task_type:
        query = query.filter(Task.task_type == task_type)

    task = query.order_by(Task.created_at.asc(), Task.id.asc()).first()
    if not task:
        return None

    updated_rows = (
        db.query(Task)
        .filter(Task.id == task.id, Task.status == "queued")
        .update(
            {
                Task.status: "running",
                Task.progress: 1,
                Task.message: "任务已开始处理",
                Task.started_at: datetime.utcnow(),
                Task.error_message: None,
            },
            synchronize_session=False,
        )
    )
    if updated_rows != 1:
        db.rollback()
        return None

    db.commit()
    return get_task_by_task_id(db, task.task_id)


def update_task_progress(
    db: Session,
    task_id: str,
    *,
    progress: Optional[int] = None,
    message: Optional[str] = None,
    status: Optional[str] = None,
) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None

    if progress is not None:
        task.progress = progress
    if message is not None:
        task.message = message
    if status is not None:
        task.status = status

    db.commit()
    db.refresh(task)
    return task


def complete_task(
    db: Session,
    task_id: str,
    *,
    result_json: Optional[str] = None,
    output_path: Optional[str] = None,
    message: str = "处理完成",
) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None

    task.status = "done"
    task.progress = 100
    task.message = message
    task.result_json = result_json
    task.output_path = output_path
    task.error_message = None
    task.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(task)
    return task


def fail_task(db: Session, task_id: str, error_message: str) -> Optional[Task]:
    task = get_task_by_task_id(db, task_id)
    if not task:
        return None

    task.status = "failed"
    task.message = f"处理失败: {error_message}"
    task.error_message = error_message
    task.finished_at = datetime.utcnow()
    db.commit()
    db.refresh(task)
    return task


def requeue_running_tasks(db: Session, *, task_type: Optional[str] = None) -> int:
    query = db.query(Task).filter(Task.status == "running")
    if task_type:
        query = query.filter(Task.task_type == task_type)

    count = query.count()
    if count == 0:
        return 0

    query.update(
        {
            Task.status: "queued",
            Task.progress: 0,
            Task.message: "服务启动后重新入队",
            Task.started_at: None,
            Task.finished_at: None,
        },
        synchronize_session=False,
    )
    db.commit()
    return count

