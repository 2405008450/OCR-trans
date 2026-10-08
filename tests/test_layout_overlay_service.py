import asyncio
import io
import json
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import fitz
import numpy as np
import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from lxml import etree
from PIL import Image, ImageDraw
from starlette.datastructures import UploadFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.service.layout_overlay import classification, imaging, ocr_providers, pipeline, qa, translation
from app.service.layout_overlay.docx_builder import build_docx
from app.service.layout_overlay.models import OcrLine, OverlayPage, Segment, parse_json, validate_bbox


def line(text="姓名", bbox=(20, 20, 140, 60), confidence=1.0):
    x1, y1, x2, y2 = bbox
    return OcrLine(text, [(x1, y1), (x2, y1), (x2, y2), (x1, y2)], bbox, confidence)


@pytest.fixture
def page(tmp_path):
    path = tmp_path / "original.png"
    image = Image.new("RGB", (400, 240), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((25, 25, 110, 48), fill="black")
    draw.rectangle((250, 40, 350, 180), fill="#205080")
    image.save(path)
    original = line()
    segment = Segment("p1s1", original.text, original.bbox, [original], translation="Name", font_size=10)
    page = OverlayPage(path, 400, 240, 200, 120, [original], [segment],
                       [{"kind": "photo", "bbox": (250, 40, 351, 181)}], path, "google")
    return page


@pytest.mark.parametrize("provider,parser", [
    ("google", ocr_providers.parse_google_response),
    ("qwen", ocr_providers.parse_qwen_response),
    ("glm", ocr_providers.parse_glm_response),
])
def test_provider_fixtures(provider, parser):
    fixture = Path(__file__).parent / "fixtures" / "layout_overlay" / f"{provider}.json"
    lines = parser(json.loads(fixture.read_text(encoding="utf-8")), 200, 100)
    assert [item.text for item in lines] == ["姓名", "张三"]
    assert lines[0].bbox == pytest.approx((10, 10, 50, 25))
    assert lines[1].bbox == pytest.approx((10, 30, 60, 45))


def test_google_and_glm_scale_to_input_pixels():
    fixtures = Path(__file__).parent / "fixtures" / "layout_overlay"
    for name, parser in [("google", ocr_providers.parse_google_response), ("glm", ocr_providers.parse_glm_response)]:
        lines = parser(json.loads((fixtures / f"{name}.json").read_text(encoding="utf-8")), 400, 300)
        assert lines[0].bbox == pytest.approx((20, 30, 100, 75))


def test_legacy_glm_location():
    lines = ocr_providers.parse_glm_response({"words_result": [{"words": "姓名", "location": {
        "left": 10, "top": 20, "width": 30, "height": 10}, "probability": {"average": 0.9}}]}, 200, 100)
    assert lines[0].bbox == (10, 20, 40, 30)
    assert lines[0].confidence == 0.9


def test_qwen_rejects_plain_text_without_coordinates():
    with pytest.raises(ValueError, match="缺少行坐标"):
        ocr_providers.parse_qwen_response({"output": {"choices": [{"message": {"content": [{"text": "姓名"}]}}]}}, 200, 100)


def test_router_fallback_and_diagnostics(monkeypatch):
    failed = Mock()
    failed.recognize.side_effect = RuntimeError("secret must not be logged")
    passed = Mock()
    passed.recognize.return_value = [line()]
    monkeypatch.setattr(ocr_providers, "PROVIDERS", {"google": lambda **kw: failed, "qwen": lambda **kw: passed})
    router = ocr_providers.OcrRouter("google", ["qwen"])
    assert router.recognize(b"image")
    assert router.last_provider == "qwen"
    assert "secret" not in str(router.warnings)


def test_router_empty_results_fail(monkeypatch):
    monkeypatch.setitem(ocr_providers.PROVIDERS, "google", lambda **kw: SimpleNamespace(recognize=lambda data: []))
    with pytest.raises(RuntimeError, match="所有 OCR"):
        ocr_providers.OcrRouter("google", []).recognize(b"image")


def test_empty_clean_background_is_valid_for_residual_qa(page, monkeypatch):
    monkeypatch.setattr(qa, "bounded_call", lambda *args, **kwargs: [])
    assert qa.residual_issues(page, ocr_providers.OcrRouter("google", [])) == []
    monkeypatch.setitem(ocr_providers.PROVIDERS, "google", lambda **kw: SimpleNamespace(recognize=lambda data: []))
    assert ocr_providers.OcrRouter("google", []).recognize(b"image", allow_empty=True) == []


def test_pdf_text_and_rotated_coordinates():
    with fitz.open() as document:
        source = document.new_page(width=300, height=200)
        source.insert_text((30, 50), "Name: Ada", fontsize=12)
        lines = ocr_providers.extract_pdf_text_lines(source, 600, 400)
        assert lines[0].text == "Name: Ada"
        assert lines[0].bbox[0] == pytest.approx(60)
        assert lines[0].font_size == pytest.approx(12)
        source.set_rotation(90)
        lines = ocr_providers.extract_pdf_text_lines(source, 400, 600)
        x1, y1, x2, y2 = lines[0].bbox
        assert 0 <= x1 < x2 <= 400 and 0 <= y1 < y2 <= 600
        assert y1 == pytest.approx(60)


@pytest.mark.parametrize("value", [[0, 0, float("nan"), 10], [0, 0, 0, 10], [1, 2, 3], [-5, 0, -1, 2]])
def test_invalid_bbox(value):
    with pytest.raises(ValueError):
        validate_bbox(value, 400, 240)


def test_classification_protection_and_partition(page, monkeypatch):
    monkeypatch.setattr(classification, "call_vision", lambda *a: {
        "segments": [{"line_ids": [0], "kind": "text"}],
        "protected_regions": [{"kind": "seal", "bbox": [20, 20, 50, 50]}]})
    classification.classify_page(page, 0, "mock", "mock")
    assert page.segments[0].protected
    monkeypatch.setattr(classification, "call_vision", lambda *a: {"segments": [], "protected_regions": []})
    with pytest.raises(ValueError, match="遗漏"):
        classification.classify_page(page, 0, "mock", "mock")


@pytest.mark.parametrize("text", ["AB1234567", "2026-10-08", "123456789", "P<CHN<ZHANG<<SAN"])
def test_identifiers_preserved(text):
    assert classification.is_identifier(text)


def test_erase_preserves_photo_and_unrelated_pixels(page):
    original = np.array(Image.open(page.image_path))
    clean = imaging.erase_page(original, page)
    assert np.array_equal(clean[40:181, 250:351], original[40:181, 250:351])
    assert np.array_equal(clean[180:240, :], original[180:240, :])
    assert (clean[25:48, 25:110] == 255).all()


def test_erase_subtracts_protection_mask(page):
    page.protected_regions = [{"kind": "seal", "bbox": (50, 25, 70, 49)}]
    original = np.array(Image.open(page.image_path))
    clean = imaging.erase_page(original, page)
    assert np.array_equal(clean[25:49, 50:70], original[25:49, 50:70])


def test_unchanged_translation_preserves_original(page):
    page.segments[0].translation = page.segments[0].text
    original = np.array(Image.open(page.image_path))
    assert np.array_equal(imaging.erase_page(original, page), original)


def test_fit_overflow(page):
    segment = page.segments[0]
    imaging.fit_segment(segment, page)
    assert not segment.overflow
    segment.translation = "Extremely long translation " * 100
    imaging.fit_segment(segment, page)
    assert segment.overflow
    assert segment.font_size >= pipeline.settings.LAYOUT_OVERLAY_MIN_FONT_PT


@pytest.mark.parametrize("items", [
    [{"segment_id": "a", "text": "Name"}, {"segment_id": "a", "text": "Name"}],
    [{"segment_id": "b", "text": "Name"}], [], [{"segment_id": "a", "text": ""}],
])
def test_translation_id_alignment(items):
    with pytest.raises(ValueError):
        translation.validate_translations({"translations": items}, {"a"})


def test_translation_can_be_reordered():
    result = translation.validate_translations(parse_json('```json\n{"translations":[{"segment_id":"b","text":"Two"},{"segment_id":"a","text":"One"}]}\n```'), {"a", "b"})
    assert result == {"a": "One", "b": "Two"}


def test_docx_editable_textboxes_and_page_sizes(page, tmp_path):
    clean = imaging.erase_page(np.array(Image.open(page.image_path)), page)
    page.clean_path = tmp_path / "clean.png"
    Image.fromarray(clean).save(page.clean_path)
    imaging.fit_segment(page.segments[0], page)
    import copy
    second = copy.deepcopy(page)
    second.width_pt, second.height_pt = 240, 144
    second.segments[0].segment_id = "p2s1"
    output = tmp_path / "output.docx"
    build_docx([page, second], output)
    with zipfile.ZipFile(output) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "v": "urn:schemas-microsoft-com:vml"}
    assert len(root.xpath("//v:imagedata", namespaces=ns)) == 2
    assert len(root.xpath("//w:txbxContent", namespaces=ns)) == 2
    assert root.xpath("//w:txbxContent//w:t/text()", namespaces=ns) == ["Name", "Name"]
    assert len(root.xpath("//w:sectPr", namespaces=ns)) == 2


@pytest.mark.parametrize("issue", [
    {"action": "exec", "id": "p1s1"}, {"action": "mark_protected", "id": "unknown"},
    {"action": "set_font_scale", "id": "p1s1", "s": 0.1},
    {"action": "set_font_scale", "id": "p1s1", "s": float("nan")},
    {"action": "nudge_bbox", "id": "p1s1", "dx": 100},
    {"action": "re_ocr", "id": "p1s1", "provider": "unknown"},
])
def test_restricted_patch_rejects_invalid_actions(page, issue):
    before = page.segments[0].bbox
    with pytest.raises(ValueError):
        qa.apply_patch_action(issue, page, retranslate=Mock(), reocr=Mock())
    assert page.segments[0].bbox == before


def test_patch_mark_protected(page):
    qa.apply_patch_action({"id": "p1s1", "action": "mark_protected"}, page, retranslate=Mock(), reocr=Mock())
    assert page.segments[0].protected
    with pytest.raises(ValueError):
        qa.apply_patch_action({"id": "p1s1", "action": "re_erase"}, page, retranslate=Mock(), reocr=Mock())


def test_source_language_qa_skips_protected(page):
    page.segments[0].translation = "姓名 Name"
    assert qa.deterministic_issues([page], "zh", "en")[0]["kind"] == "source_language"
    page.segments[0].kind = "no_translate"
    assert qa.deterministic_issues([page], "zh", "en") == []


def test_qa_render_failure_keeps_docx(page, tmp_path, monkeypatch):
    docx = tmp_path / "output.docx"
    build_docx([page], docx)
    before = docx.read_bytes()
    monkeypatch.setattr(qa, "residual_issues", lambda *a, **kw: [])
    monkeypatch.setattr(qa, "render_docx", Mock(side_effect=FileNotFoundError()))
    report = qa.run_qa([page], docx, tmp_path, pipeline.normalize_options(), Mock())
    assert report["status"] == "partial"
    assert not report["visual_checked"]
    assert docx.read_bytes() == before


def test_qa_round_limit_and_best_result(page, tmp_path, monkeypatch):
    docx = tmp_path / "output.docx"
    build_docx([page], docx)
    original = docx.read_bytes()
    monkeypatch.setattr(qa.settings, "LAYOUT_OVERLAY_MAX_QA_ROUNDS", 99)
    monkeypatch.setattr(qa, "residual_issues", lambda *a, **kw: [])
    monkeypatch.setattr(qa, "render_docx", lambda *a: tmp_path / "fake.pdf")
    calls = []

    def issues(*args):
        calls.append(1)
        return [{"id": "p1s1", "kind": "residual", "action": "re_erase"}] * len(calls)

    monkeypatch.setattr(qa, "visual_issues", issues)
    report = qa.run_qa([page], docx, tmp_path, pipeline.normalize_options(), lambda: docx.write_bytes(b"worse"))
    assert len(calls) == 3
    assert len(report["rounds"]) == 3
    assert report["selected_round"] == 0
    assert docx.read_bytes() == original


def test_pipeline_pdf_text_layer_skips_ocr(tmp_path, monkeypatch):
    source = tmp_path / "source.pdf"
    with fitz.open() as document:
        page = document.new_page(width=300, height=200)
        page.insert_text((30, 50), "Name", fontsize=12)
        document.save(source)
    monkeypatch.setattr(pipeline.settings, "OUTPUT_DIR", str(tmp_path / "out"))
    monkeypatch.setattr(pipeline.OcrRouter, "recognize", Mock(side_effect=AssertionError("should skip cloud OCR")))
    monkeypatch.setattr(classification, "call_vision", lambda *a: {"segments": [{"line_ids": [0], "kind": "text"}], "protected_regions": []})
    monkeypatch.setattr(translation, "call_translation", lambda *a: '{"translations":[{"segment_id":"p1s1","text":"Nombre"}]}')
    updates = []

    async def progress(value, message):
        updates.append(value)

    result = asyncio.run(pipeline.execute_layout_overlay_task(task_id="test", display_no="test", input_path=str(source),
                        source_lang="en", target_lang="es", enable_qa=False, progress_callback=progress))
    assert Path(result["output_docx"]).is_file()
    assert result["qa_report"]["status"] == "disabled"
    assert result["page_count"] == 1
    assert 95 in updates
    with zipfile.ZipFile(result["debug_archive"]) as archive:
        assert {"original-1.png", "clean-1.png", "layout.json", "qa_report.json"} <= set(archive.namelist())


def test_layout_overlay_api_and_validation(monkeypatch):
    from app.controller import task as controller
    async def fake_submit(**kwargs):
        assert kwargs["enable_qa"] is False
        return SimpleNamespace(task_id="test", deduped=False)
    monkeypatch.setattr(controller.task_queue_service, "submit_layout_overlay_task", fake_submit)
    app = FastAPI()
    app.include_router(controller.router)
    with TestClient(app) as client:
        assert client.get("/task/layout-overlay/config").json()["default_ocr_provider"] == "google"
        response = client.post("/task/layout-overlay", files={"file": ("scan.png", b"image", "image/png")}, data={"enable_qa": "false"})
        assert response.status_code == 200
        assert response.json()["task_id"] == "test"
        assert client.post("/task/layout-overlay", files={"file": ("scan.txt", b"data")}).status_code == 400
        assert client.post("/task/layout-overlay", files={"file": ("scan.pdf", b"data")}, data={"ocr_provider": "invalid"}).status_code == 400


def test_queue_output_registration():
    from app.service.task_queue_service import TaskQueueService
    result = {"output_docx": "out.docx", "qa_report_path": "qa.json", "debug_archive": "debug.zip"}
    files = TaskQueueService._extract_output_files("layout_overlay", result, "out.docx", "scan.pdf")
    assert {f["path"] for f in files} == {"out.docx", "qa.json", "debug.zip"}
    assert TaskQueueService.DEFAULT_TASK_TYPE_LIMITS["layout_overlay"] == 1


def test_queue_submission_execution_status_and_download(tmp_path, monkeypatch):
    from sqlalchemy import create_engine, text
    from sqlalchemy.orm import sessionmaker
    from app.db.database import Base
    from app.model.entity import Task
    from app.service import task_queue_service as queue_module
    from app.controller import task as controller

    engine = create_engine(f"sqlite:///{tmp_path / 'tasks.db'}", connect_args={"check_same_thread": False})
    Base.metadata.create_all(engine)
    with engine.begin() as connection:
        connection.execute(text("CREATE UNIQUE INDEX ux_test_overlay ON task (request_fingerprint) "
                                "WHERE request_fingerprint IS NOT NULL AND status IN ('queued', 'running')"))
    sessions = sessionmaker(bind=engine, expire_on_commit=False)
    monkeypatch.setattr(queue_module, "SessionLocal", sessions)
    monkeypatch.setattr(controller, "SessionLocal", sessions)
    monkeypatch.setattr(pipeline.settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(pipeline.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    service = queue_module.TaskQueueService()
    monkeypatch.setattr(controller, "task_queue_service", service)
    monkeypatch.setattr(classification, "call_vision", lambda *a: {"segments": [{"line_ids": [0], "kind": "text"}], "protected_regions": []})
    monkeypatch.setattr(translation, "call_translation", lambda *a: '{"translations":[{"segment_id":"p1s1","text":"Nombre"}]}')
    document = fitz.open()
    source = document.new_page(width=300, height=200)
    source.insert_text((30, 50), "Name", fontsize=12)
    content = document.tobytes()
    document.close()

    def upload():
        return UploadFile(filename="test.pdf", file=io.BytesIO(content))

    async def scenario():
        first = await service.submit_layout_overlay_task(file=upload(), source_lang="en", target_lang="es", enable_qa=False)
        second = await service.submit_layout_overlay_task(file=upload(), source_lang="en", target_lang="es", enable_qa=False)
        assert second.task_id == first.task_id and second.deduped
        await service._execute_task(first.task_id)
        return first.task_id

    task_id = asyncio.run(scenario())
    app = FastAPI()
    app.include_router(controller.router)
    with TestClient(app) as client:
        snapshot = client.get(f"/task/layout-overlay/status/{task_id}").json()
        assert snapshot["status"] == "done"
        assert len(snapshot["output_files"]) == 3
        download = client.get(f"/task/{task_id}/download", params={"file_path": snapshot["result"]["output_docx"]})
        assert download.status_code == 200
        with zipfile.ZipFile(io.BytesIO(download.content)) as archive:
            assert b"Nombre" in archive.read("word/document.xml")
    with sessions() as db:
        task = db.query(Task).filter_by(task_id=task_id).one()
        task.task_type = "doc_translate"
        db.commit()
    with TestClient(app) as client:
        assert client.get(f"/task/layout-overlay/status/{task_id}").status_code == 404
    engine.dispose()


def test_queue_upload_size_limit_cleans_staging(tmp_path, monkeypatch):
    from app.service.task_queue_service import TaskQueueService, UploadSizeLimitError
    monkeypatch.setattr(pipeline.settings, "UPLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(pipeline.settings, "LAYOUT_OVERLAY_UPLOAD_MAX_MB", 0)
    upload = UploadFile(filename="input.png", file=io.BytesIO(b"too large"))
    with pytest.raises(UploadSizeLimitError):
        asyncio.run(TaskQueueService().submit_layout_overlay_task(file=upload))
    assert not list(tmp_path.rglob("*.png"))


def test_qa_cancellation_is_not_swallowed(page, tmp_path, monkeypatch):
    from app.service.task_queue_service import TaskCancelledError
    output = tmp_path / "output.docx"
    build_docx([page], output)
    def cancel(*args):
        raise TaskCancelledError("cancelled")
    with pytest.raises(TaskCancelledError):
        qa.run_qa([page], output, tmp_path, pipeline.normalize_options(), Mock(), cancel)


def test_textured_erase_preserves_protected_pixels(page):
    grid = np.indices((page.height, page.width)).sum(axis=0) % 2
    image = np.repeat((160 + grid * 90)[..., None], 3, axis=2).astype(np.uint8)
    image[25:40, 30:100] = 0
    clean = imaging.erase_page(image, page)
    assert np.array_equal(clean[40:181, 250:351], image[40:181, 250:351])
    assert np.array_equal(clean[100:120, :200], image[100:120, :200])
    assert clean[30, 50, 0] > image[30, 50, 0]


def test_qa_budget_stops_timed_out_calls():
    import time
    from app.service.layout_overlay.budget import bounded_call
    started = time.monotonic()
    with pytest.raises(TimeoutError):
        bounded_call(time.sleep, 30, budget_seconds=0.2)
    assert time.monotonic() - started < 5
    assert bounded_call(abs, -7, budget_seconds=5) == 7
