# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.controller import task as task_controller
from app.main import app
from app.service import audio_transcription_service as service
from app.service.audio_transcription_service import (
    normalize_audio_transcription_options,
    validate_audio_transcription_filename,
)
from app.service.task_queue_service import TaskSubmitResult


def test_audio_transcription_filename_and_options() -> None:
    for extension in (".wav", ".mp3", ".m4a", ".aac", ".flac", ".ogg", ".opus", ".wma", ".amr"):
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
    invalid = client.post(
        "/task/audio-transcription",
        files={"file": ("fake.txt", b"not-audio", "text/plain")},
    )
    assert invalid.status_code == 400
