from datetime import datetime

from sqlalchemy import Boolean, Column, DateTime, Integer, String, Text

from app.db.database import Base


class Task(Base):
    __tablename__ = "task"

    id = Column(Integer, primary_key=True, index=True)
    task_id = Column(String, unique=True, index=True, nullable=True)
    display_no = Column(String, unique=True, index=True, nullable=True)
    task_type = Column(String, index=True, nullable=False, default="ocr")
    filename = Column(String, nullable=False)
    status = Column(String, index=True, default="queued")
    progress = Column(Integer, nullable=False, default=0)
    message = Column(String, nullable=True)
    output_path = Column(String, nullable=True)
    params_json = Column(Text, nullable=True)
    input_files_json = Column(Text, nullable=True)
    result_json = Column(Text, nullable=True)
    error_message = Column(Text, nullable=True)
    retry_count = Column(Integer, nullable=False, default=0)
    cancel_requested = Column(Boolean, nullable=False, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    started_at = Column(DateTime, nullable=True)
    finished_at = Column(DateTime, nullable=True)

