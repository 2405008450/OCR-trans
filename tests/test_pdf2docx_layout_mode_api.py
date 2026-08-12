import io
import sys
from pathlib import Path
from types import SimpleNamespace

import anyio
import pytest
from fastapi import HTTPException
from starlette.datastructures import UploadFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.controller import task as task_controller
from app.service import pdf2docx_service
from app.service.chat_preserve_docx_service import ChatPreserveResult
from app.service.web_asset_preserve_docx_service import WebAssetPreserveResult


def _upload(filename: str) -> UploadFile:
    return UploadFile(filename=filename, file=io.BytesIO(b"content"))


def test_pdf2docx_config_exposes_layout_modes():
    result = anyio.run(task_controller.get_pdf2docx_config)

    assert result["default_layout_mode"] == "ocr_html"
    assert "ocr_html" in result["layout_modes"]
    assert "fixed_layout" in result["layout_modes"]
    assert "chat_preserve" in result["layout_modes"]
    assert "web_asset_preserve" in result["layout_modes"]
    assert result["layout_modes"]["web_asset_preserve"]["recommended_model"] == "anthropic/claude-sonnet-5"
    assert "anthropic/claude-sonnet-5" in result["models"]


def test_execute_pdf2docx_fixed_layout_branch(tmp_path, monkeypatch):
    input_path = tmp_path / "scan.png"
    input_path.write_bytes(b"fake-image")
    monkeypatch.setattr(pdf2docx_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(pdf2docx_service, "ensure_gemini_route_configured", lambda route: route or "openrouter")
    monkeypatch.setattr(
        pdf2docx_service,
        "ocr_file",
        lambda **kwargs: {"text": "<p>姓名：张三</p>", "total_pages": 1},
    )
    fixed_calls = []

    def fake_fixed(html_text, output_path, **kwargs):
        from docx import Document

        fixed_calls.append(html_text)
        Document().save(output_path)
        Path(kwargs["html_output_path"]).write_text(html_text, encoding="utf-8")
        Path(kwargs["debug_layout_path"]).write_text("{}", encoding="utf-8")

    monkeypatch.setattr(pdf2docx_service, "convert_html_to_fixed_layout_docx", fake_fixed)

    async def scenario():
        return await pdf2docx_service.execute_pdf2docx_task_from_path(
            task_id="task-fixed",
            display_no="000003",
            input_path=str(input_path),
            original_filename="scan.png",
            layout_mode="fixed_layout",
        )

    result = anyio.run(scenario)

    assert fixed_calls and "姓名：张三" in fixed_calls[0]
    assert result["layout_mode"] == "fixed_layout"
    assert result["fixed_layout_rendered"] is True
    assert result["output_layout_json"].endswith("_layout.json")
    assert result["warnings"] == []


def test_execute_pdf2docx_fixed_layout_failure_falls_back(tmp_path, monkeypatch):
    input_path = tmp_path / "scan.png"
    input_path.write_bytes(b"fake-image")
    monkeypatch.setattr(pdf2docx_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(pdf2docx_service, "ensure_gemini_route_configured", lambda route: route or "openrouter")
    monkeypatch.setattr(pdf2docx_service, "ocr_file", lambda **kwargs: "<p>正文</p>")
    monkeypatch.setattr(
        pdf2docx_service,
        "convert_html_to_fixed_layout_docx",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("browser unavailable")),
    )
    fallback_calls = []

    def fake_libreoffice(text, output_path, **kwargs):
        from docx import Document

        fallback_calls.append(text)
        Document().save(output_path)
        Path(kwargs["html_output_path"]).write_text(text, encoding="utf-8")

    monkeypatch.setattr(pdf2docx_service, "convert_text_to_word_via_libreoffice", fake_libreoffice)

    async def scenario():
        return await pdf2docx_service.execute_pdf2docx_task_from_path(
            task_id="task-fixed-fallback",
            display_no="000004",
            input_path=str(input_path),
            original_filename="scan.png",
            layout_mode="fixed_layout",
        )

    result = anyio.run(scenario)

    assert fallback_calls == ["<p>正文</p>"]
    assert result["fixed_layout_rendered"] is False
    assert result["output_layout_json"] is None
    assert any("已回退到通用文档模式" in warning for warning in result["warnings"])


def test_run_pdf2docx_passes_layout_mode(monkeypatch):
    captured = {}

    async def fake_submit_pdf2docx_task(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(task_id="task-1", deduped=False)

    monkeypatch.setattr(task_controller.task_queue_service, "submit_pdf2docx_task", fake_submit_pdf2docx_task)

    result = anyio.run(
        task_controller.run_pdf2docx,
        _upload("chat.png"),
        "model-a",
        "openrouter",
        "chat_preserve",
    )

    assert result["status"] == "ACCEPTED"
    assert captured["layout_mode"] == "chat_preserve"


def test_run_pdf2docx_batch_passes_layout_mode(monkeypatch):
    captured = []

    async def fake_submit_pdf2docx_task(**kwargs):
        captured.append(kwargs)
        return SimpleNamespace(task_id=f"task-{len(captured)}", deduped=False)

    monkeypatch.setattr(task_controller.task_queue_service, "submit_pdf2docx_task", fake_submit_pdf2docx_task)

    result = anyio.run(
        task_controller.run_pdf2docx_batch,
        [_upload("a.png"), _upload("b.png")],
        "model-a",
        "openrouter",
        "chat_preserve",
    )

    assert result["status"] == "ACCEPTED"
    assert len(captured) == 2
    assert {item["layout_mode"] for item in captured} == {"chat_preserve"}


def test_run_pdf2docx_rejects_invalid_layout_mode():
    with pytest.raises(HTTPException) as exc_info:
        anyio.run(
            task_controller.run_pdf2docx,
            _upload("chat.png"),
            "model-a",
            "openrouter",
            "bad_mode",
        )

    assert exc_info.value.status_code == 400


def test_execute_pdf2docx_chat_preserve_branch_returns_layout_json(tmp_path, monkeypatch):
    input_path = tmp_path / "chat.png"
    input_path.write_bytes(b"fake-image")
    monkeypatch.setattr(pdf2docx_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(pdf2docx_service, "ensure_gemini_route_configured", lambda route: route or "openrouter")

    def fake_convert_chat_screenshot_to_docx(**kwargs):
        from docx import Document

        Document().save(kwargs["output_docx_path"])
        Path(kwargs["layout_json_path"]).write_text("{}", encoding="utf-8")
        return ChatPreserveResult(
            raw_text="Fabian\nHello",
            total_pages=1,
            asset_count=2,
            fallback_count=0,
            layout={
                "mode": "chat_preserve",
                "render": {
                    "debug_overlays": [
                        str(Path(kwargs["assets_dir"]) / "debug_overlay_page_001.png").replace("\\", "/")
                    ]
                },
            },
        )

    monkeypatch.setattr(pdf2docx_service, "convert_chat_screenshot_to_docx", fake_convert_chat_screenshot_to_docx)

    async def scenario():
        return await pdf2docx_service.execute_pdf2docx_task_from_path(
            task_id="task-1",
            display_no="000001",
            input_path=str(input_path),
            original_filename="chat.png",
            model=pdf2docx_service.PDF2DOCX_DEFAULT_MODEL,
            gemini_route="openrouter",
            layout_mode="chat_preserve",
        )

    result = anyio.run(scenario)

    assert result["layout_mode"] == "chat_preserve"
    assert result["output_layout_json"].endswith("_layout.json")
    assert result["output_debug_overlay"].endswith("debug_overlay_page_001.png")
    assert result["asset_count"] == 2
    assert (tmp_path / "outputs" / "pdf2docx" / "000001" / "chat_raw.txt").read_text(encoding="utf-8") == "Fabian\nHello"


def test_execute_pdf2docx_web_asset_preserve_branch_returns_layout_json(tmp_path, monkeypatch):
    input_path = tmp_path / "web.png"
    input_path.write_bytes(b"fake-image")
    monkeypatch.setattr(pdf2docx_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(pdf2docx_service, "ensure_gemini_route_configured", lambda route: route or "openrouter")

    def fake_convert_web_screenshot_to_docx(**kwargs):
        from docx import Document

        Document().save(kwargs["output_docx_path"])
        Path(kwargs["layout_json_path"]).write_text("{}", encoding="utf-8")
        return WebAssetPreserveResult(
            raw_text="Product Overview\n[保留图片: product - Main image]",
            total_pages=1,
            asset_count=1,
            fallback_count=0,
            layout={
                "mode": "web_asset_preserve",
                "render": {
                    "debug_overlays": [
                        str(Path(kwargs["assets_dir"]) / "debug_overlay_page_001.png").replace("\\", "/")
                    ]
                },
            },
        )

    monkeypatch.setattr(pdf2docx_service, "convert_web_screenshot_to_docx", fake_convert_web_screenshot_to_docx)

    async def scenario():
        return await pdf2docx_service.execute_pdf2docx_task_from_path(
            task_id="task-2",
            display_no="000002",
            input_path=str(input_path),
            original_filename="web.png",
            model="anthropic/claude-sonnet-5",
            gemini_route="openrouter",
            layout_mode="web_asset_preserve",
        )

    result = anyio.run(scenario)

    assert result["layout_mode"] == "web_asset_preserve"
    assert result["model"] == "anthropic/claude-sonnet-5"
    assert result["output_layout_json"].endswith("_layout.json")
    assert result["output_debug_overlay"].endswith("debug_overlay_page_001.png")
    assert result["asset_count"] == 1
    raw_text = tmp_path / "outputs" / "pdf2docx" / "000002" / "web_raw.txt"
    assert "Product Overview" in raw_text.read_text(encoding="utf-8")
