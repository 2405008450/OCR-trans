from sqlalchemy.orm import Session
from app.model.entity import Task

def create_task(db: Session, filename: str) -> Task:
    task = Task(filename=filename)
    db.add(task)
    db.commit()
    db.refresh(task)
    return task

def update_task_result(db: Session, task: Task, output_path: str):
    task.status = "DONE"
    task.output_path = output_path
    db.commit()
