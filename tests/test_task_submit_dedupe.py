import io
import json
import sys
from pathlib import Path

import anyio
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
from starlette.datastructures import UploadFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.db.database import Base
from app.model.entity import Task
from app.repository import task_repo
from app.service import task_queue_service as queue_module


def _build_test_service(tmp_path, monkeypatch):
    engine = create_engine(
        f"sqlite:///{tmp_path / 'tasks.db'}",
        connect_args={"check_same_thread": False},
    )
    Base.metadata.create_all(bind=engine)
    with engine.begin() as connection:
        connection.execute(
            text(
                "CREATE UNIQUE INDEX ux_task_active_request_fingerprint "
                "ON task (request_fingerprint) "
                "WHERE request_fingerprint IS NOT NULL AND status IN ('queued', 'running')"
            )
        )

    testing_session = sessionmaker(autocommit=False, autoflush=False, bind=engine)
    monkeypatch.setattr(queue_module, "SessionLocal", testing_session)
    monkeypatch.setattr(queue_module.settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    return queue_module.TaskQueueService(), testing_session


def _upload(filename, content):
    return UploadFile(filename=filename, file=io.BytesIO(content))


def test_request_fingerprint_changes_with_params_or_file_content():
    files = [{"role": "input", "filename": "a.pdf", "size": 3, "sha256": "abc"}]
    same = queue_module.TaskQueueService.build_request_fingerprint("pdf2docx", {"model": "m1"}, files)
    same_again = queue_module.TaskQueueService.build_request_fingerprint("pdf2docx", {"model": "m1"}, files)
    different_param = queue_module.TaskQueueService.build_request_fingerprint("pdf2docx", {"model": "m2"}, files)
    different_file = queue_module.TaskQueueService.build_request_fingerprint(
        "pdf2docx",
        {"model": "m1"},
        [{"role": "input", "filename": "a.pdf", "size": 3, "sha256": "def"}],
    )

    assert same == same_again
    assert same != different_param
    assert same != different_file


def test_pdf2docx_reuses_active_duplicate_task(tmp_path, monkeypatch):
    service, testing_session = _build_test_service(tmp_path, monkeypatch)

    async def scenario():
        first = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
        )
        second = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
        )

        assert first.task_id == second.task_id
        assert first.deduped is False
        assert second.deduped is True

    anyio.run(scenario)

    with testing_session() as db:
        tasks = db.query(Task).all()
        assert len(tasks) == 1
        task = tasks[0]
        assert task.status == "queued"
        assert json.loads(task.input_files_json)["original_filename"] == "same.pdf"
        assert json.loads(task.file_fingerprints_json)[0]["sha256"]


def test_pdf2docx_layout_mode_participates_in_dedupe(tmp_path, monkeypatch):
    service, testing_session = _build_test_service(tmp_path, monkeypatch)

    async def scenario():
        first = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
            layout_mode="ocr_html",
        )
        second = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
            layout_mode="chat_preserve",
        )
        third = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
            layout_mode="web_asset_preserve",
        )

        assert first.task_id != second.task_id
        assert second.task_id != third.task_id
        assert first.task_id != third.task_id
        assert first.deduped is False
        assert second.deduped is False
        assert third.deduped is False

    anyio.run(scenario)

    with testing_session() as db:
        tasks = db.query(Task).all()
        assert len(tasks) == 3
        params = [json.loads(task.params_json) for task in tasks]
        assert {item["layout_mode"] for item in params} == {"ocr_html", "chat_preserve", "web_asset_preserve"}


def test_completed_duplicate_can_be_submitted_again(tmp_path, monkeypatch):
    service, testing_session = _build_test_service(tmp_path, monkeypatch)

    async def scenario():
        first = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
        )
        with testing_session() as db:
            task_repo.complete_task(db, first.task_id, message="done for test")

        second = await service.submit_pdf2docx_task(
            file=_upload("same.pdf", b"same-content"),
            model="model-a",
            gemini_route="openrouter",
        )

        assert first.task_id != second.task_id
        assert second.deduped is False

    anyio.run(scenario)

    with testing_session() as db:
        assert db.query(Task).count() == 2


def test_batch_metadata_is_persisted_for_batch_submissions(tmp_path, monkeypatch):
    service, testing_session = _build_test_service(tmp_path, monkeypatch)

    async def scenario():
        first = await service.submit_pdf2docx_task(
            file=_upload("a.pdf", b"a-content"),
            model="model-a",
            gemini_route="openrouter",
            batch_id="batch-1",
            batch_name="批量结果.zip",
            batch_index=1,
            batch_total=2,
        )
        second = await service.submit_pdf2docx_task(
            file=_upload("b.pdf", b"b-content"),
            model="model-a",
            gemini_route="openrouter",
            batch_id="batch-1",
            batch_name="批量结果.zip",
            batch_index=2,
            batch_total=2,
        )
        return first, second

    first, second = anyio.run(scenario)

    with testing_session() as db:
        tasks = task_repo.list_tasks_by_batch_id(db, "batch-1")
        assert [task.task_id for task in tasks] == [first.task_id, second.task_id]
        assert [task.batch_index for task in tasks] == [1, 2]
        assert all(task.batch_total == 2 for task in tasks)
