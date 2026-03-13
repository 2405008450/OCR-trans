from sqlalchemy import inspect, text

from app.db.database import Base, engine
from app.model import entity


def _ensure_task_table_columns():
    inspector = inspect(engine)
    tables = inspector.get_table_names()
    if "task" not in tables:
        return

    existing_columns = {column["name"] for column in inspector.get_columns("task")}
    missing_columns = {
        "task_id": "ALTER TABLE task ADD COLUMN task_id VARCHAR",
        "task_type": "ALTER TABLE task ADD COLUMN task_type VARCHAR DEFAULT 'ocr' NOT NULL",
        "progress": "ALTER TABLE task ADD COLUMN progress INTEGER DEFAULT 0 NOT NULL",
        "message": "ALTER TABLE task ADD COLUMN message VARCHAR",
        "params_json": "ALTER TABLE task ADD COLUMN params_json TEXT",
        "input_files_json": "ALTER TABLE task ADD COLUMN input_files_json TEXT",
        "result_json": "ALTER TABLE task ADD COLUMN result_json TEXT",
        "error_message": "ALTER TABLE task ADD COLUMN error_message TEXT",
        "retry_count": "ALTER TABLE task ADD COLUMN retry_count INTEGER DEFAULT 0 NOT NULL",
        "cancel_requested": "ALTER TABLE task ADD COLUMN cancel_requested BOOLEAN DEFAULT 0 NOT NULL",
        "started_at": "ALTER TABLE task ADD COLUMN started_at DATETIME",
        "finished_at": "ALTER TABLE task ADD COLUMN finished_at DATETIME",
    }

    with engine.begin() as connection:
        for column_name, ddl in missing_columns.items():
            if column_name not in existing_columns:
                connection.execute(text(ddl))
        connection.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS ix_task_task_id ON task (task_id)"))
        connection.execute(text("CREATE INDEX IF NOT EXISTS ix_task_status ON task (status)"))
        connection.execute(text("CREATE INDEX IF NOT EXISTS ix_task_task_type ON task (task_type)"))


def init_db():
    Base.metadata.create_all(bind=engine)
    _ensure_task_table_columns()


if __name__ == "__main__":
    init_db()

