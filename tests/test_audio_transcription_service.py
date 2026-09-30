# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.controller import task as task_controller
from app.main import app
from app.service import audio_transcription_service as service
from app.service.audio_transcription_service import (
    AudioTranscriptionError,
    normalize_audio_transcription_options,
    validate_audio_transcription_filename,
)
from app.service.task_queue_service import TaskSubmitResult


def test_audio_transcription_filename_and_options() -> None:
    for extension in (".wav", ".mp3", ".m4a", ".mp4", ".aac", ".flac", ".ogg", ".opus", ".wma", ".amr"):
        assert validate_audio_transcription_filename(f"sample{extension}") == extension
    with pytest.raises(ValueError, match="不支持"):
        validate_audio_transcription_filename("sample.txt")
    assert normalize_audio_transcription_options(language="AUTO", enable_itn=True) == {
        "model_name": "qwen3-asr-flash-filetrans",
        "language": "auto",
        "enable_itn": True,
    }
    with pytest.raises(ValueError, match="语言"):
        normalize_audio_transcription_options(language="xx", enable_itn=False)


def test_dashscope_session_ignores_environment_proxy() -> None:
    with service._create_dashscope_session() as session:
        assert session.trust_env is False


def test_normalize_and_write_all_timeline_outputs(tmp_path: Path) -> None:
    raw = {
        "audio_info": {"format": "wav", "sample_rate": 16000},
        "transcripts": [
            {
                "channel_id": 0,
                "text": "你好，世界。",
                "sentences": [
                    {
                        "sentence_id": 1,
                        "begin_time": 120,
                        "end_time": 1560,
                        "text": "你好，世界。",
                        "language": "zh",
                        "emotion": "neutral",
                        "words": [
                            {"begin_time": 120, "end_time": 500, "text": "你"},
                            {"begin_time": 500, "end_time": 800, "text": "好", "punctuation": "，"},
                            {"begin_time": 900, "end_time": 1200, "text": "世"},
                            {"begin_time": 1200, "end_time": 1560, "text": "界", "punctuation": "。"},
                        ],
                    }
                ],
            }
        ],
    }
    normalized = service._normalize_result(raw, {"usage": {"seconds": 2}}, "task-1", "访谈.wav")
    assert normalized["segments"][0]["start"] == 0.12
    assert normalized["words"][1]["word"] == "好，"
    outputs = service._write_outputs(tmp_path, "访谈.wav", normalized, raw)
    for path in outputs.values():
        assert Path(path).is_file()
    assert "00:00:00.120 --> 00:00:01.560" in Path(outputs["timeline_txt"]).read_text(encoding="utf-8")
    assert "00:00:00,120 --> 00:00:01,560" in Path(outputs["srt"]).read_text(encoding="utf-8")
    assert Path(outputs["word_tsv"]).read_text(encoding="utf-8-sig").splitlines()[1].endswith("\t你\t1")
    payload = json.loads(Path(outputs["result_json"]).read_text(encoding="utf-8"))
    assert payload["text"] == "你好，世界。"
    with zipfile.ZipFile(outputs["archive_zip"]) as archive:
        assert len(archive.namelist()) == 6


def test_long_model_sentence_is_resegmented_from_word_timestamps() -> None:
    characters = list("这是一个很长的测试句子需要根据逐词时间戳重新切分避免字幕持续时间过长影响阅读体验")
    words = []
    for index, character in enumerate(characters):
        punctuation = "，" if index in {11, 23} else ("。" if index == len(characters) - 1 else "")
        words.append({
            "begin_time": index * 300,
            "end_time": (index + 1) * 300,
            "text": character,
            "punctuation": punctuation,
        })
    raw = {
        "transcripts": [{
            "channel_id": 0,
            "text": "".join(characters),
            "sentences": [{
                "sentence_id": 0,
                "begin_time": 0,
                "end_time": len(characters) * 300,
                "text": "".join(characters),
                "language": "zh",
                "words": words,
            }],
        }],
    }
    normalized = service._normalize_result(raw, {}, "task-long", "long.wav")
    assert len(normalized["model_segments"]) == 1
    assert len(normalized["segments"]) >= 2
    assert normalized["timeline_source"] == "word_timestamps_resegmented"
    assert all(segment["end"] > segment["start"] for segment in normalized["segments"])
    assert max(segment["end"] - segment["start"] for segment in normalized["segments"]) <= 6.3


def test_audio_transcription_submit_endpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    async def fake_submit_audio_transcription_task(*, file, params):
        assert file.filename == "meeting.m4a"
        assert params["language"] == "zh"
        assert params["enable_itn"] is True
        return TaskSubmitResult(task_id="transcription-task-id")

    monkeypatch.setattr(
        task_controller.task_queue_service,
        "submit_audio_transcription_task",
        fake_submit_audio_transcription_task,
    )
    client = TestClient(app)
    response = client.post(
        "/task/audio-transcription",
        files={"file": ("meeting.m4a", b"audio-data", "audio/mp4")},
        data={"language": "zh", "enable_itn": "true"},
    )
    assert response.status_code == 200
    assert response.json()["task_id"] == "transcription-task-id"

    async def fake_submit_video(*, file, params):
        assert file.filename == "interview.mp4"
        assert params["language"] == "auto"
        return TaskSubmitResult(task_id="video-task-id")

    monkeypatch.setattr(
        task_controller.task_queue_service,
        "submit_audio_transcription_task",
        fake_submit_video,
    )
    video = client.post(
        "/task/audio-transcription",
        files={"file": ("interview.mp4", b"video-data", "video/mp4")},
        data={"language": "auto", "enable_itn": "true"},
    )
    assert video.status_code == 200
    assert video.json()["task_id"] == "video-task-id"
    invalid = client.post(
        "/task/audio-transcription",
        files={"file": ("fake.txt", b"not-audio", "text/plain")},
    )
    assert invalid.status_code == 400


def test_extract_audio_stream_copies_supported_codec(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    source = tmp_path / "clip.mp4"
    source.write_bytes(b"video-bytes")
    commands: list[list[str]] = []

    def fake_which(name: str) -> str:
        return name

    def fake_run(command: list[str], **kwargs):
        commands.append(command)
        if command[0] == "ffprobe":
            return subprocess.CompletedProcess(command, 0, stdout="aac\n", stderr="")
        Path(command[-1]).write_bytes(b"audio-bytes")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr(service.shutil, "which", fake_which)
    monkeypatch.setattr(service.subprocess, "run", fake_run)
    extracted = service._extract_audio_from_video(source)
    assert extracted.suffix == ".m4a"
    assert extracted.read_bytes() == b"audio-bytes"
    assert commands[-1][commands[-1].index("-c:a") + 1] == "copy"


def test_extract_audio_transcodes_unsupported_codec(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    source = tmp_path / "clip.mp4"
    source.write_bytes(b"video-bytes")
    commands: list[list[str]] = []

    def fake_run(command: list[str], **kwargs):
        commands.append(command)
        if command[0] == "ffprobe":
            return subprocess.CompletedProcess(command, 0, stdout="ac3\n", stderr="")
        output = Path(command[-1])
        if "copy" in command:
            return subprocess.CompletedProcess(command, 1, stdout="", stderr="copy failed")
        output.write_bytes(b"aac-bytes")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr(service.shutil, "which", lambda name: name)
    monkeypatch.setattr(service.subprocess, "run", fake_run)
    extracted = service._extract_audio_from_video(source)
    assert extracted.name == "clip.model-audio.m4a"
    assert extracted.read_bytes() == b"aac-bytes"
    assert "aac" in commands[-1]
    assert "copy" not in commands[-1]


def test_extract_audio_reports_missing_soundtrack(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    source = tmp_path / "silent.mp4"
    source.write_bytes(b"video-bytes")

    def fake_run(command: list[str], **kwargs):
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr(service.shutil, "which", lambda name: name)
    monkeypatch.setattr(service.subprocess, "run", fake_run)
    with pytest.raises(AudioTranscriptionError, match="没有可识别的音轨"):
        service._extract_audio_from_video(source)


def test_video_transcription_uploads_extracted_audio_then_deletes_it(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "talk.mp4"
    source.write_bytes(b"video")
    extracted = tmp_path / "talk.model-audio.m4a"
    uploaded: list[Path] = []

    def fake_extract(path: Path, log_callback=None) -> Path:
        extracted.write_bytes(b"audio")
        return extracted

    def fake_upload(session, audio_path: Path, timeout: int) -> str:
        uploaded.append(audio_path)
        assert audio_path.read_bytes() == b"audio"
        return "oss://bucket/talk.model-audio.m4a"

    monkeypatch.setattr(service.settings, "DASHSCOPE_API_KEY", "test-key")
    monkeypatch.setattr(service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(service, "_extract_audio_from_video", fake_extract)
    monkeypatch.setattr(service, "_upload_temporary_file", fake_upload)
    monkeypatch.setattr(service, "_submit_task", lambda session, **kwargs: "task-1")
    monkeypatch.setattr(
        service,
        "_wait_and_download",
        lambda *args, **kwargs: (
            {"transcripts": [{"text": "你好", "sentences": [{"text": "你好", "begin_time": 0, "end_time": 400}]}]},
            {},
        ),
    )
    result = service._run_audio_transcription(
        display_no="T1",
        input_path=str(source),
        original_filename="访谈.mp4",
        params={"language": "zh", "enable_itn": True},
    )
    assert uploaded == [extracted]
    assert not extracted.exists()
    assert source.exists()
    assert result["text"] == "你好"


def test_extract_real_mp4_audio_track(tmp_path: Path) -> None:
    if not shutil.which("ffmpeg") or not shutil.which("ffprobe"):
        pytest.skip("ffmpeg is not installed")
    source = tmp_path / "clip.mp4"
    created = subprocess.run(
        [
            "ffmpeg", "-y",
            "-f", "lavfi", "-i", "sine=frequency=440:duration=0.4",
            "-f", "lavfi", "-i", "color=c=black:s=32x32:r=10:d=0.4",
            "-shortest", "-c:v", "mpeg4", "-c:a", "aac",
            str(source),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if created.returncode != 0 or not source.is_file():
        pytest.skip("ffmpeg cannot synthesize a test video")
    logs: list[str] = []
    extracted = service._extract_audio_from_video(source, logs.append)
    try:
        assert extracted.suffix == ".m4a"
        assert extracted.stat().st_size > 0
        assert any("直接封装" in item for item in logs)
    finally:
        extracted.unlink(missing_ok=True)
