"""M4A 转 WAV 的路径、转换、队列和下载验收。"""
from __future__ import annotations

import asyncio
import io
import json
import shutil
import subprocess
import wave
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from starlette.datastructures import UploadFile

from app.controller import task as controller
from app.db.database import Base
from app.main import app
from app.repository import task_repo
from app.service import audio_convert_service as service
from app.service import task_queue_service as queue_module


@pytest.fixture
def sandbox(tmp_path, monkeypatch):
    shared = tmp_path / "共享资料"
    shared.mkdir()
    monkeypatch.setattr(service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps([str(shared)]))
    monkeypatch.setattr(service.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", "{}")
    monkeypatch.setattr(service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")
    monkeypatch.setattr(service.settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_FILES", 50)
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_FILE_MB", 200)
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_TOTAL_MB", 200)
    monkeypatch.setattr(service.settings, "FFMPEG_BIN_DIR", "")
    return shared


@pytest.fixture
def queue(sandbox, tmp_path, monkeypatch):
    engine = create_engine(f"sqlite:///{tmp_path / 'tasks.db'}", connect_args={"check_same_thread": False})
    Base.metadata.create_all(engine)
    sessions = sessionmaker(bind=engine, expire_on_commit=False)
    monkeypatch.setattr(queue_module, "SessionLocal", sessions)
    monkeypatch.setattr(controller, "SessionLocal", sessions)
    monkeypatch.setattr(controller, "BASE_DIR", tmp_path)
    instance = queue_module.TaskQueueService()
    monkeypatch.setattr(controller, "task_queue_service", instance)
    yield instance, sessions
    engine.dispose()


def make_m4a(path: Path, rate=44100, channels=2):
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        pytest.skip("真实转换验收需要 FFmpeg")
    path.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run([ffmpeg, "-nostdin", "-v", "error", "-y", "-f", "lavfi", "-i",
                    f"sine=frequency=440:sample_rate={rate}:duration=0.5", "-ac", str(channels),
                    "-c:a", "aac", str(path)], check=True, capture_output=True)
    return path


def upload(name, content):
    return UploadFile(filename=name, file=io.BytesIO(content))


def resolve_output(web_path):
    return Path(service.settings.OUTPUT_DIR) / web_path.removeprefix("outputs/")


def test_discover_single_recursive_hidden_and_truncated(sandbox, monkeypatch):
    (sandbox / "录音.M4A").write_bytes(b"audio")
    (sandbox / "空.m4a").touch()
    (sandbox / ".隐藏.m4a").write_bytes(b"audio")
    (sandbox / "text.txt").write_text("ignored", encoding="utf-8")
    child = sandbox / "客户"
    child.mkdir()
    (child / "录音.m4a").write_bytes(b"audio")
    result = service.discover_audio_files(directory_path=str(sandbox))
    assert len(result["files"]) == 3
    assert result["selectable_count"] == 2
    assert any(not item["selectable"] and item["name"] == "空.m4a" for item in result["files"])
    shallow = service.discover_audio_files(directory_path=str(sandbox), recursive=False)
    assert len(shallow["files"]) == 2
    single = service.discover_audio_files(directory_path=str(sandbox / "录音.M4A"))
    assert single["files"][0]["relative_path"] == "录音.M4A"
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_FILES", 1)
    assert service.discover_audio_files(directory_path=str(sandbox))["truncated"]


def test_shared_unc_mapping_and_revalidation(sandbox, monkeypatch):
    (sandbox / "录音.m4a").write_bytes(b"audio")
    unc = r"\\win-server\服务器资料7"
    monkeypatch.setattr(service.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", json.dumps({unc: str(sandbox)}))
    result = service.discover_audio_files(directory_path=unc)
    assert result["files"][0]["name"] == "录音.m4a"
    assert service.resolve_selected_file(unc, "录音.m4a").is_file()
    (sandbox / "录音.m4a").unlink()
    with pytest.raises(FileNotFoundError):
        service.resolve_selected_file(unc, "录音.m4a")


@pytest.mark.parametrize("relative", ["../外部.m4a", "/外部.m4a", "C:/外部.m4a", r"..\外部.m4a", "./录音.m4a", "客户//录音.m4a"])
def test_reject_escaping_paths(sandbox, relative):
    with pytest.raises((ValueError, PermissionError)):
        service.resolve_selected_file(str(sandbox), relative)


def test_shared_path_scope_and_scan_options(sandbox, tmp_path):
    outside = tmp_path / "外部.m4a"
    outside.write_bytes(b"audio")
    with pytest.raises(PermissionError):
        service.discover_audio_files(directory_path=str(outside))
    child = sandbox / "客户"
    child.mkdir()
    (child / "录音.m4a").write_bytes(b"audio")
    with pytest.raises(ValueError, match="扫描选项"):
        service.resolve_selected_file(str(sandbox), "客户/录音.m4a", False)
    (sandbox / "单个.m4a").write_bytes(b"audio")
    with pytest.raises(PermissionError):
        service.resolve_selected_file(str(sandbox / "单个.m4a"), "客户/录音.m4a")


def test_symlink_cannot_escape_shared_directory(sandbox, tmp_path):
    outside = tmp_path / "outside.m4a"
    outside.write_bytes(b"audio")
    link = sandbox / "link.m4a"
    try:
        link.symlink_to(outside)
    except OSError:
        pytest.skip("当前环境不允许创建符号链接")
    assert service.discover_audio_files(directory_path=str(sandbox))["files"] == []
    with pytest.raises(PermissionError):
        service.resolve_selected_file(str(sandbox), "link.m4a")


@pytest.mark.parametrize("rate,channels", [(44100, 2), (16000, 1), (48000, 2)])
def test_real_conversion_preserves_audio_spec_and_duration(sandbox, rate, channels):
    source = make_m4a(sandbox / "客户 录音.M4A", rate, channels)
    target = sandbox.parent / "转换.wav"
    service.convert_file(source, target)
    with wave.open(str(target), "rb") as wav:
        assert wav.getsampwidth() == 2
        assert wav.getframerate() == rate
        assert wav.getnchannels() == channels
        assert abs(wav.getnframes() / rate - 0.5) < 0.08
    assert source.is_file()


def test_timeout_removes_incomplete_output(sandbox, monkeypatch):
    output = sandbox.parent / "timeout.wav"
    def timeout(command, **kwargs):
        output.write_bytes(b"partial")
        assert "-ar" not in command and "-ac" not in command
        raise subprocess.TimeoutExpired(command, kwargs["timeout"])
    monkeypatch.setattr(service.subprocess, "run", timeout)
    with pytest.raises(service.AudioConvertError, match="超时"):
        service.convert_file(sandbox / "录音.m4a", output)
    assert not output.exists()


def test_no_audio_stream_and_corrupt_file_cleanup(sandbox):
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        pytest.skip("需要 FFmpeg")
    video = sandbox / "无音轨.m4a"
    subprocess.run([ffmpeg, "-nostdin", "-v", "error", "-y", "-f", "lavfi", "-i",
                    "color=size=16x16:duration=0.2", "-an", "-c:v", "mpeg4", "-f", "mp4", str(video)],
                   check=True, capture_output=True)
    for source in (video, sandbox / "损坏.m4a"):
        if not source.exists():
            source.write_bytes(b"broken")
        output = sandbox.parent / "invalid.wav"
        with pytest.raises(service.AudioConvertError, match="没有音轨"):
            service.convert_file(source, output)
        assert not output.exists()


def test_missing_configured_ffmpeg_reports_unavailable(sandbox, monkeypatch):
    monkeypatch.setattr(service.settings, "FFMPEG_BIN_DIR", str(sandbox / "missing"))
    assert not service.get_audio_convert_config()["ffmpeg_available"]
    with pytest.raises(service.AudioConvertError, match="FFMPEG_BIN_DIR"):
        service.ffmpeg_binary()


def test_api_shared_partial_failure_persistence_and_downloads(queue, sandbox):
    instance, sessions = queue
    source = make_m4a(sandbox / "客户 空格" / "录音.m4a")
    (sandbox / "损坏.m4a").write_bytes(b"broken")
    client = TestClient(app)
    assert client.get("/audio-convert").status_code == 200
    config = client.get("/task/audio-convert/config").json()
    assert config["allowed_extensions"] == [".m4a"]
    response = client.post("/task/audio-convert/discover", json={"directory_path":str(sandbox)})
    assert response.status_code == 200
    paths = [item["relative_path"] for item in response.json()["files"]]
    body = {"directory_path":str(sandbox), "relative_paths":paths}
    first = client.post("/task/audio-convert", json=body)
    assert first.status_code == 200, first.text
    task_id = first.json()["task_id"]
    repeated = client.post("/task/audio-convert", json=body).json()
    assert repeated["task_id"] == task_id and repeated["deduped"]
    asyncio.run(instance._execute_task(task_id))
    status = client.get(f"/task/audio-convert/status/{task_id}").json()
    assert status["status"] == "done"
    result = status["result"]
    assert result["success_count"] == result["failed_count"] == 1
    assert "失败 1" in status["message"]
    assert len(status["output_files"]) == 2
    success = next(item for item in result["files"] if item["status"] == "done")
    response = client.get(f"/task/{task_id}/download", params={"file_path":success["output_path"]})
    assert response.status_code == 200 and response.content[:4] == b"RIFF"
    response = client.get(f"/task/{task_id}/download", params={"file_path":result["archive_zip"]})
    assert response.status_code == 200
    with zipfile.ZipFile(io.BytesIO(response.content)) as bundle:
        assert bundle.namelist() == ["客户 空格/录音.wav"]
    assert source.is_file() and (sandbox / "损坏.m4a").read_bytes() == b"broken"
    assert not (Path(service.settings.UPLOAD_DIR) / "audio_convert_staging" / status["display_no"]).exists()
    with sessions() as db:
        task = task_repo.get_task_by_task_id(db, task_id)
        assert task.task_type == "audio_convert" and task.task_label == "音频转换"


def test_upload_same_names_dedupe_and_outputs(queue, sandbox):
    instance, _ = queue
    content = make_m4a(sandbox / "录音.m4a").read_bytes()
    client = TestClient(app)
    def submit():
        return client.post("/task/audio-convert/upload", files=[
            ("files", ("录音.m4a", content, "audio/mp4")),
            ("files", ("录音.M4A", content, "audio/mp4")),
        ])
    first = submit()
    assert first.status_code == 200, first.text
    repeated = submit().json()
    task_id = first.json()["task_id"]
    assert repeated["task_id"] == task_id and repeated["deduped"]
    asyncio.run(instance._execute_task(task_id))
    result = client.get(f"/task/audio-convert/status/{task_id}").json()["result"]
    assert result["success_count"] == 2
    with zipfile.ZipFile(resolve_output(result["archive_zip"])) as bundle:
        assert bundle.namelist() == ["录音.wav", "录音_2.wav"]
    assert not list((Path(service.settings.UPLOAD_DIR) / "_tmp_uploads" / "audio_convert").glob("*"))


def test_all_failed_keeps_file_errors(queue):
    instance, _ = queue
    client = TestClient(app)
    response = client.post("/task/audio-convert/upload", files=[("files", ("损坏.m4a", b"broken", "audio/mp4"))])
    task_id = response.json()["task_id"]
    asyncio.run(instance._execute_task(task_id))
    task = client.get(f"/task/audio-convert/status/{task_id}").json()
    assert task["status"] == "failed"
    assert task["result"]["failed_count"] == 1
    assert task["result"]["files"][0]["error"]
    assert task["result"]["archive_zip"] is None


def test_path_changed_after_submit_fails_only_changed_file(queue, sandbox):
    instance, sessions = queue
    keep = make_m4a(sandbox / "保留.m4a")
    gone = sandbox / "删除.m4a"
    gone.write_bytes(keep.read_bytes())
    async def run():
        submitted = await instance.submit_audio_convert_task(directory_path=str(sandbox), relative_paths=[keep.name, gone.name])
        gone.unlink()
        await instance._execute_task(submitted.task_id)
        return instance.get_task_status(submitted.task_id)
    task = asyncio.run(run())
    assert task["status"] == "done" and task["result"]["failed_count"] == 1


def test_single_file_path_and_changed_content_uses_new_task(queue, sandbox):
    instance, _ = queue
    path = make_m4a(sandbox / "单个 录音.M4A")
    async def run():
        first = await instance.submit_audio_convert_task(directory_path=str(path), relative_paths=[path.name])
        # 修改文件快照后，相同路径应创建新的转换任务。
        path.write_bytes(path.read_bytes() + b"\0")
        changed = await instance.submit_audio_convert_task(directory_path=str(path), relative_paths=[path.name])
        assert first.task_id != changed.task_id and not changed.deduped
        await instance._execute_task(changed.task_id)
        return instance.get_task_status(changed.task_id)
    task = asyncio.run(run())
    assert task["status"] == "done"
    assert task["result"]["output_files"][0]["filename"] == "单个 录音.wav"


def test_api_limits_invalid_format_and_missing_path(queue, sandbox, monkeypatch):
    client = TestClient(app)
    assert client.post("/task/audio-convert/upload", files=[("files", ("a.mp3", b"x"))]).status_code == 400
    assert client.post("/task/audio-convert/upload", files=[("files", ("a.m4a", b""))]).status_code == 400
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_FILE_MB", 1)
    response = client.post("/task/audio-convert/upload", files=[("files", ("a.m4a", b"x" * (1048576 + 1)))])
    assert response.status_code == 413
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_TOTAL_MB", 1)
    response = client.post("/task/audio-convert/upload", files=[("files", ("a.m4a", b"x" * 600000)), ("files", ("b.m4a", b"x" * 600000))])
    assert response.status_code == 413
    monkeypatch.setattr(service.settings, "AUDIO_CONVERT_MAX_FILES", 1)
    assert client.post("/task/audio-convert/upload", files=[("files", ("a.m4a", b"x")), ("files", ("b.m4a", b"x"))]).status_code == 400
    assert client.post("/task/audio-convert/discover", json={"directory_path":str(sandbox / "missing")}).status_code == 400
    assert client.get("/task/audio-convert/status/missing").status_code == 404
    temporary = Path(service.settings.UPLOAD_DIR) / "_tmp_uploads" / "audio_convert"
    assert not temporary.exists() or not list(temporary.glob("*"))


def test_copy_rechecks_actual_size_and_shared_permission(sandbox, monkeypatch):
    source = sandbox / "大文件.m4a"
    source.write_bytes(b"x" * 20)
    target = sandbox.parent / "stage.m4a"
    with pytest.raises(ValueError, match="大小超过"):
        service._stage_shared_file(source, target, 10)
    def denied(path, *args, **kwargs):
        raise PermissionError("共享目录无读取权限")
    monkeypatch.setattr(service, "resolve_allowed_shared_input_path", denied)
    with pytest.raises(PermissionError, match="无读取权限"):
        service.discover_audio_files(directory_path=str(sandbox))
