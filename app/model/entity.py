from sqlalchemy import Column, Integer, String, DateTime
from datetime import datetime
from app.db.database import Base

class Task(Base):
    __tablename__ = "task"

    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, nullable=False)
    status = Column(String, default="PENDING")
    output_path = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
