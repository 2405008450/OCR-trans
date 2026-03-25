from datetime import datetime

from sqlalchemy import inspect, text

from app.db.database import Base, engine
from app.model import entity


LABEL_MAP = {
    'ocr': 'OCR',
    'number_check': 'Number Check',
    'zhongfanyi': 'Zhongfanyi',
    'alignment': 'Alignment',
    'drivers_license': 'Drivers License',
    'doc_translate': 'Doc Translate',
    'pdf2docx': 'PDF2DOCX',
}


def _ensure_task_table_columns():
    inspector = inspect(engine)
    if 'task' not in inspector.get_table_names():
        return
    existing_columns = {column['name'] for column in inspector.get_columns('task')}
    missing_columns = {
        'task_id': 'ALTER TABLE task ADD COLUMN task_id VARCHAR',
        'display_no': 'ALTER TABLE task ADD COLUMN display_no VARCHAR',
        'task_type': "ALTER TABLE task ADD COLUMN task_type VARCHAR DEFAULT 'ocr' NOT NULL",
        'progress': 'ALTER TABLE task ADD COLUMN progress INTEGER DEFAULT 0 NOT NULL',
        'message': 'ALTER TABLE task ADD COLUMN message VARCHAR',
        'params_json': 'ALTER TABLE task ADD COLUMN params_json TEXT',
        'input_files_json': 'ALTER TABLE task ADD COLUMN input_files_json TEXT',
        'result_json': 'ALTER TABLE task ADD COLUMN result_json TEXT',
        'output_files_json': 'ALTER TABLE task ADD COLUMN output_files_json TEXT',
        'error_message': 'ALTER TABLE task ADD COLUMN error_message TEXT',
        'retry_count': 'ALTER TABLE task ADD COLUMN retry_count INTEGER DEFAULT 0 NOT NULL',
        'cancel_requested': 'ALTER TABLE task ADD COLUMN cancel_requested BOOLEAN DEFAULT 0 NOT NULL',
        'started_at': 'ALTER TABLE task ADD COLUMN started_at DATETIME',
        'finished_at': 'ALTER TABLE task ADD COLUMN finished_at DATETIME',
        'updated_at': 'ALTER TABLE task ADD COLUMN updated_at DATETIME',
        'task_label': 'ALTER TABLE task ADD COLUMN task_label VARCHAR',
    }
    with engine.begin() as connection:
        for column_name, ddl in missing_columns.items():
            if column_name not in existing_columns:
                connection.execute(text(ddl))
        connection.execute(text('CREATE UNIQUE INDEX IF NOT EXISTS ix_task_task_id ON task (task_id)'))
        connection.execute(text('CREATE UNIQUE INDEX IF NOT EXISTS ix_task_display_no ON task (display_no)'))
        connection.execute(text('CREATE INDEX IF NOT EXISTS ix_task_status ON task (status)'))
        connection.execute(text('CREATE INDEX IF NOT EXISTS ix_task_task_type ON task (task_type)'))
        connection.execute(text('CREATE INDEX IF NOT EXISTS ix_task_created_at ON task (created_at)'))
        connection.execute(text('CREATE INDEX IF NOT EXISTS ix_task_updated_at ON task (updated_at)'))
        connection.execute(text('UPDATE task SET updated_at = created_at WHERE updated_at IS NULL AND created_at IS NOT NULL'))
        for task_type, label in LABEL_MAP.items():
            connection.execute(text('UPDATE task SET task_label = :label WHERE task_type = :tt AND task_label IS NULL'), {'label': label, 'tt': task_type})
        rows = connection.execute(text('SELECT id, created_at FROM task WHERE display_no IS NULL ORDER BY id')).fetchall()
        for row in rows:
            task_id = row[0]
            created_at = row[1]
            if isinstance(created_at, datetime):
                dt = created_at
            else:
                try:
                    dt = datetime.fromisoformat(str(created_at))
                except Exception:
                    dt = datetime.now()
            display_no = f'{dt:%Y%m%d}-{int(task_id):06d}'
            connection.execute(text('UPDATE task SET display_no = :display_no WHERE id = :task_id'), {'display_no': display_no, 'task_id': task_id})


def init_db():
    Base.metadata.create_all(bind=engine)
    _ensure_task_table_columns()


if __name__ == '__main__':
    init_db()
