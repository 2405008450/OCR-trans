# -*- coding: utf-8 -*-

from __future__ import annotations

import io
from pathlib import Path

import pytest
import anyio
from fastapi import UploadFile
from fastapi.testclient import TestClient

from app.controller import task as task_controller
from app.main import app
from app.service import audio_check_service, gemini_service
from app.service.audio_check_service import (
    AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
    AUDIO_CHECK_DEFAULT_USER_PROMPT,
    normalize_audio_check_options,
    render_markdown,
    validate_audio_filename,
)
from app.service.task_queue_service import TaskQueueService, TaskSubmitResult, UploadSizeLimitError


def test_audio_filename_and_options_validation() -> None:
    for extension in (".wav", ".mp3", ".aiff", ".aac", ".ogg", ".flac"):
        actual_extension, mime_type = validate_audio_filename(f"sample{extension}")
        assert actual_extension == extension
        assert mime_type.startswith("audio/")

    with pytest.raises(ValueError, match="不支持"):
        validate_audio_filename("fake.txt")

    options = normalize_audio_check_options(
        model_name="google/gemini-3.6-flash",
        gemini_route="google",
        system_prompt=AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
        user_prompt=AUDIO_CHECK_DEFAULT_USER_PROMPT,
        temperature=0.1,
        max_output_tokens=32768,
        timeout_seconds=120,
    )
    assert options["temperature"] == 0.1
    with pytest.raises(ValueError, match="温度"):
        normalize_audio_check_options(**{**options, "temperature": 2.1})
    with pytest.raises(ValueError, match="token"):
        normalize_audio_check_options(**{**options, "max_output_tokens": 100})


def test_audio_check_sends_original_bytes_without_preprocessing(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "sample.mp3"
    original_bytes = b"ID3-original-audio-bytes"
    source.write_bytes(original_bytes)
    monkeypatch.setattr(audio_check_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    captured = {}

    def fake_generate_audio_analysis(**kwargs):
        captured.update(kwargs)
        return "## 结论\n\n原始音频检查完成。"

    monkeypatch.setattr(audio_check_service, "generate_audio_analysis", fake_generate_audio_analysis)
    params = normalize_audio_check_options(
        model_name="google/gemini-3.6-flash",
        gemini_route="google",
        system_prompt=AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
        user_prompt="只根据原始音频判断",
        temperature=0.1,
        max_output_tokens=32768,
        timeout_seconds=120,
    )
    result = audio_check_service._run_audio_check(
        display_no="20260727-000001",
        input_path=str(source),
        original_filename=source.name,
        params=params,
        log_callback=None,
    )
    assert captured["audio_bytes"] == original_bytes
    assert captured["user_prompt"] == "只根据原始音频判断"
    assert "metrics" not in result
    report = Path(result["report_markdown"]).read_text(encoding="utf-8")
    assert "未执行转码、剪切、降噪、响度归一化或指标预处理" in report
    assert "客观指标" not in report


def test_markdown_renderer_disables_raw_html() -> None:
    rendered = render_markdown("# 检查\n\n<script>alert(1)</script>")
    assert "<h1>检查</h1>" in rendered
    assert "<script>" not in rendered
    assert "&lt;script&gt;" in rendered


def test_openrouter_audio_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    captured = {}

    class FakeCompletions:
        def create(self, **kwargs):
            captured.update(kwargs)
            message = type("Message", (), {"content": "检查完成"})()
            choice = type("Choice", (), {"message": message})()
            return type("Response", (), {"choices": [choice]})()

    class FakeClient:
        def __init__(self, **kwargs):
            self.chat = type("Chat", (), {"completions": FakeCompletions()})()

    monkeypatch.setattr(gemini_service, "OpenAI", FakeClient)
    monkeypatch.setattr(gemini_service.settings, "OPENROUTER_API_KEY", "test-key")
    result = gemini_service.generate_audio_analysis(
        system_prompt="系统",
        user_prompt="检查",
        audio_bytes=b"RIFFdata",
        mime_type="audio/wav",
        audio_format="wav",
        model="google/gemini-3.6-flash",
        route="openrouter",
    )
    assert result == "检查完成"
    content = captured["messages"][1]["content"]
    assert content[1]["type"] == "input_audio"
    assert content[1]["input_audio"]["format"] == "wav"
    assert content[1]["input_audio"]["data"]


def test_audio_check_submit_endpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    async def fake_submit_audio_check_task(*, file, params):
        assert file.filename == "sample.mp3"
        assert params["model_name"] == "google/gemini-3.6-flash"
        return TaskSubmitResult(task_id="audio-task-id")

    monkeypatch.setattr(
        task_controller.task_queue_service,
        "submit_audio_check_task",
        fake_submit_audio_check_task,
    )
    client = TestClient(app)
    response = client.post(
        "/task/audio-check",
        files={"file": ("sample.mp3", b"ID3data", "audio/mpeg")},
        data={
            "system_prompt": AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
            "user_prompt": AUDIO_CHECK_DEFAULT_USER_PROMPT,
            "model_name": "google/gemini-3.6-flash",
            "gemini_route": "google",
            "temperature": "0.1",
            "max_output_tokens": "32768",
            "timeout_seconds": "120",
        },
    )
    assert response.status_code == 200
    assert response.json()["task_id"] == "audio-task-id"

    invalid = client.post(
        "/task/audio-check",
        files={"file": ("fake.txt", b"not audio", "text/plain")},
    )
    assert invalid.status_code == 400


def test_audio_upload_rejects_empty_and_oversize(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(audio_check_service.settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(audio_check_service.settings, "AUDIO_CHECK_MAX_MB", 1)
    service = TaskQueueService()
    params = normalize_audio_check_options(
        model_name="google/gemini-3.6-flash",
        gemini_route="google",
        system_prompt=AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
        user_prompt=AUDIO_CHECK_DEFAULT_USER_PROMPT,
        temperature=0.1,
        max_output_tokens=32768,
        timeout_seconds=120,
    )

    async def submit(content: bytes):
        upload = UploadFile(filename="sample.mp3", file=io.BytesIO(content))
        return await service.submit_audio_check_task(file=upload, params=params)

    with pytest.raises(ValueError, match="不能为空"):
        anyio.run(submit, b"")
    with pytest.raises(UploadSizeLimitError, match="音频文件超过"):
        anyio.run(submit, b"x" * (1024 * 1024 + 1))
    staged_dir = tmp_path / "uploads" / "_tmp_uploads" / "audio_check"
    assert not list(staged_dir.glob("*"))
