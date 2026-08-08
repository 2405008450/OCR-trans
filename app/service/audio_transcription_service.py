# -*- coding: utf-8 -*-
from __future__ import annotations

import asyncio
import json
import re
import time
import zipfile
from concurrent.futures import Executor
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Optional

import requests

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename


AUDIO_TRANSCRIPTION_MODEL = "qwen3-asr-flash-filetrans"
AUDIO_TRANSCRIPTION_ALLOWED_EXTENSIONS = {
    ".aac", ".amr", ".flac", ".m4a", ".mp3", ".ogg", ".opus", ".wav", ".wma"
}
AUDIO_TRANSCRIPTION_LANGUAGES = {
    "auto": "自动识别",
    "zh": "中文",
    "yue": "粤语",
    "en": "英语",
    "ja": "日语",
    "ko": "韩语",
    "de": "德语",
    "fr": "法语",
    "es": "西班牙语",
    "it": "意大利语",
    "pt": "葡萄牙语",
    "ru": "俄语",
}
TIMELINE_MAX_DURATION_SECONDS = 6.0
TIMELINE_PAUSE_SPLIT_SECONDS = 0.35
TIMELINE_MAX_TEXT_UNITS = 42
TIMELINE_SOFT_SPLIT_UNITS = 28
TIMELINE_TERMINAL_PUNCTUATION = "。！？!?．."
TIMELINE_SOFT_PUNCTUATION = "，,；;：:、"


class AudioTranscriptionError(RuntimeError):
    pass


def _create_dashscope_session() -> requests.Session:
    """创建仅供音频转写使用的直连会话，不读取系统代理环境变量。"""
    session = requests.Session()
    session.trust_env = False
    return session


def get_audio_transcription_config() -> dict[str, Any]:
    return {
        "model": AUDIO_TRANSCRIPTION_MODEL,
        "model_label": "Qwen3 ASR FileTrans",
        "configured": bool(settings.DASHSCOPE_API_KEY),
        "allowed_extensions": sorted(AUDIO_TRANSCRIPTION_ALLOWED_EXTENSIONS),
        "max_file_mb": settings.AUDIO_TRANSCRIPTION_MAX_MB,
        "languages": AUDIO_TRANSCRIPTION_LANGUAGES,
        "defaults": {"language": "auto", "enable_itn": True},
        "outputs": ["带时间轴 TXT", "纯文本 TXT", "SRT", "VTT", "逐词 TSV", "JSON", "ZIP"],
    }


def validate_audio_transcription_filename(filename: str) -> str:
    extension = Path(filename or "").suffix.lower()
    if extension not in AUDIO_TRANSCRIPTION_ALLOWED_EXTENSIONS:
        supported = "、".join(item.lstrip(".").upper() for item in sorted(AUDIO_TRANSCRIPTION_ALLOWED_EXTENSIONS))
        raise ValueError(f"不支持的音频格式，仅支持：{supported}")
    return extension


def normalize_audio_transcription_options(*, language: str, enable_itn: bool) -> dict[str, Any]:
    normalized_language = (language or "auto").strip().lower()
    if normalized_language not in AUDIO_TRANSCRIPTION_LANGUAGES:
        raise ValueError("不支持的音频语言")
    return {
        "model_name": AUDIO_TRANSCRIPTION_MODEL,
        "language": normalized_language,
        "enable_itn": bool(enable_itn),
    }


def _json_response(response: requests.Response, operation: str) -> dict[str, Any]:
    try:
        body: Any = response.json()
    except ValueError:
        body = response.text[:2000]
    if not response.ok:
        raise AudioTranscriptionError(f"{operation}失败：HTTP {response.status_code}：{body}")
    if not isinstance(body, dict):
        raise AudioTranscriptionError(f"{operation}返回格式异常")
    return body


def _upload_temporary_file(session: requests.Session, audio_path: Path, timeout: int) -> str:
    base_url = settings.DASHSCOPE_BASE_URL.rstrip("/")
    api_key = settings.DASHSCOPE_API_KEY
    response = session.get(
        f"{base_url}/uploads",
        headers={"Authorization": f"Bearer {api_key}"},
        params={"action": "getPolicy", "model": AUDIO_TRANSCRIPTION_MODEL},
        timeout=timeout,
    )
    policy_body = _json_response(response, "获取上传凭证")
    policy = policy_body.get("data")
    if not isinstance(policy, dict):
        raise AudioTranscriptionError("上传凭证缺少 data")
    required = {
        "upload_host", "upload_dir", "oss_access_key_id", "policy", "signature",
        "x_oss_object_acl", "x_oss_forbid_overwrite",
    }
    missing = sorted(required - policy.keys())
    if missing:
        raise AudioTranscriptionError(f"上传凭证缺少字段：{', '.join(missing)}")
    object_key = f"{str(policy['upload_dir']).rstrip('/')}/{audio_path.name}"
    form_data = [
        ("OSSAccessKeyId", str(policy["oss_access_key_id"])),
        ("policy", str(policy["policy"])),
        ("Signature", str(policy["signature"])),
        ("key", object_key),
        ("x-oss-object-acl", str(policy["x_oss_object_acl"])),
        ("x-oss-forbid-overwrite", str(policy["x_oss_forbid_overwrite"])),
        ("success_action_status", "200"),
    ]
    with audio_path.open("rb") as audio_file:
        upload_response = session.post(
            str(policy["upload_host"]),
            data=form_data,
            files={"file": (audio_path.name, audio_file, "application/octet-stream")},
            timeout=max(timeout, 300),
        )
    if not upload_response.ok:
        raise AudioTranscriptionError(
            f"上传音频失败：HTTP {upload_response.status_code}：{upload_response.text[:1000]}"
        )
    return f"oss://{object_key}"


def _submit_task(
    session: requests.Session,
    *,
    file_url: str,
    language: str,
    enable_itn: bool,
    timeout: int,
) -> str:
    parameters: dict[str, Any] = {
        "channel_id": [0],
        "enable_itn": enable_itn,
        "enable_words": True,
    }
    if language != "auto":
        parameters["language"] = language
    response = session.post(
        f"{settings.DASHSCOPE_BASE_URL.rstrip('/')}/services/audio/asr/transcription",
        headers={
            "Authorization": f"Bearer {settings.DASHSCOPE_API_KEY}",
            "Content-Type": "application/json",
            "X-DashScope-Async": "enable",
            "X-DashScope-OssResourceResolve": "enable",
        },
        json={
            "model": AUDIO_TRANSCRIPTION_MODEL,
            "input": {"file_url": file_url},
            "parameters": parameters,
        },
        timeout=timeout,
    )
    body = _json_response(response, "提交转写任务")
    output = body.get("output")
    task_id = output.get("task_id") if isinstance(output, dict) else None
    if not task_id:
        raise AudioTranscriptionError("提交结果缺少 task_id")
    return str(task_id)


def _wait_and_download(session: requests.Session, task_id: str, timeout: int, log_callback=None) -> tuple[dict[str, Any], dict[str, Any]]:
    deadline = time.monotonic() + settings.AUDIO_TRANSCRIPTION_MAX_WAIT_SECONDS
    last_status = ""
    task_body: dict[str, Any] = {}
    while time.monotonic() < deadline:
        response = session.get(
            f"{settings.DASHSCOPE_BASE_URL.rstrip('/')}/tasks/{task_id}",
            headers={"Authorization": f"Bearer {settings.DASHSCOPE_API_KEY}"},
            timeout=timeout,
        )
        task_body = _json_response(response, "查询转写任务")
        output = task_body.get("output")
        status = str(output.get("task_status") or "") if isinstance(output, dict) else ""
        if status != last_status and log_callback:
            log_callback(f"[audio-transcription] DashScope 状态：{status or 'UNKNOWN'}")
            last_status = status
        if status == "SUCCEEDED":
            result = output.get("result") if isinstance(output, dict) else None
            url = result.get("transcription_url") if isinstance(result, dict) else None
            if not url:
                raise AudioTranscriptionError("任务结果缺少 transcription_url")
            transcription = session.get(str(url), timeout=max(timeout, 300))
            return _json_response(transcription, "下载转写结果"), task_body
        if status in {"FAILED", "UNKNOWN"}:
            raise AudioTranscriptionError(f"转写任务失败：{task_body}")
        time.sleep(max(1.0, settings.AUDIO_TRANSCRIPTION_POLL_INTERVAL_SECONDS))
    raise AudioTranscriptionError(f"等待转写任务超时：{task_id}")


def _is_cjk_character(value: str) -> bool:
    if not value:
        return False
    codepoint = ord(value[0])
    return (
        0x3400 <= codepoint <= 0x4DBF
        or 0x4E00 <= codepoint <= 0x9FFF
        or 0x3040 <= codepoint <= 0x30FF
        or 0xAC00 <= codepoint <= 0xD7AF
    )


def _join_timeline_words(items: list[dict[str, Any]]) -> str:
    text = ""
    closing_punctuation = set("，。！？、；：,.!?;:%)]}）】》〉」』'")
    opening_punctuation = set("([{（【《〈「『'")
    for item in items:
        token = str(item.get("word") or "").strip()
        if not token:
            continue
        if not text:
            text = token
        elif (
            token[0] in closing_punctuation
            or text[-1] in opening_punctuation
            or _is_cjk_character(text[-1])
            or _is_cjk_character(token[0])
        ):
            text += token
        else:
            text += f" {token}"
    return re.sub(r"\s+([，。！？、；：,.!?;:%）】》〉」』])", r"\1", text).strip()


def _timeline_text_units(text: str) -> int:
    return sum(2 if _is_cjk_character(character) else 1 for character in text if not character.isspace())


def _build_timeline_segments(
    words: list[dict[str, Any]],
    model_segments: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """根据逐词时间戳重新切分字幕，避免直接使用模型的过长句段。"""
    if not words:
        return [dict(item) for item in model_segments]

    sentence_metadata = {
        item.get("id"): {
            "language": item.get("language"),
            "emotion": item.get("emotion"),
        }
        for item in model_segments
    }
    ordered_words = sorted((dict(item) for item in words), key=lambda item: (item["start"], item["end"]))
    result: list[dict[str, Any]] = []
    buffer: list[dict[str, Any]] = []

    def flush(count: Optional[int] = None) -> None:
        if not buffer:
            return
        take = len(buffer) if count is None else max(1, min(count, len(buffer)))
        items = buffer[:take]
        text = _join_timeline_words(items)
        if not text:
            del buffer[:take]
            return
        start = max(0.0, float(items[0]["start"]))
        end = max(float(item["end"]) for item in items)
        if end <= start:
            end = start + 0.2
        sentence_ids = list(dict.fromkeys(item.get("sentence_id") for item in items))
        metadata = next(
            (sentence_metadata.get(sentence_id) for sentence_id in sentence_ids if sentence_metadata.get(sentence_id)),
            {},
        )
        result.append({
            "id": len(result),
            "channel_id": items[0].get("channel_id", 0),
            "start": round(start, 3),
            "end": round(end, 3),
            "text": text,
            "language": metadata.get("language"),
            "emotion": metadata.get("emotion"),
            "source_sentence_ids": sentence_ids,
        })
        del buffer[:take]

    for word in ordered_words:
        if buffer:
            pause = max(0.0, float(word["start"]) - max(float(item["end"]) for item in buffer))
            current_duration = max(float(item["end"]) for item in buffer) - float(buffer[0]["start"])
            current_units = _timeline_text_units(_join_timeline_words(buffer))
            pause_break = (
                pause >= TIMELINE_PAUSE_SPLIT_SECONDS
                and current_duration >= 1.4
                and current_units >= 10
            ) or (
                pause >= 0.7
                and current_duration >= 0.8
                and current_units >= 8
            )
            if (
                word.get("channel_id", 0) != buffer[0].get("channel_id", 0)
                or pause_break
            ):
                flush()

        buffer.append(word)
        text = _join_timeline_words(buffer)
        duration = max(float(item["end"]) for item in buffer) - float(buffer[0]["start"])
        units = _timeline_text_units(text)
        token = str(word.get("word") or "").rstrip()
        terminal_break = bool(token and token[-1] in TIMELINE_TERMINAL_PUNCTUATION and duration >= 0.8)
        soft_break = bool(
            token
            and token[-1] in TIMELINE_SOFT_PUNCTUATION
            and duration >= 2.5
            and units >= TIMELINE_SOFT_SPLIT_UNITS
        )
        forced_break = duration >= TIMELINE_MAX_DURATION_SECONDS or units >= TIMELINE_MAX_TEXT_UNITS
        if terminal_break or soft_break:
            flush()
        elif forced_break:
            fallback_index = None
            for index, candidate in enumerate(buffer[:-1]):
                candidate_token = str(candidate.get("word") or "").rstrip()
                if candidate_token and candidate_token[-1] in TIMELINE_SOFT_PUNCTUATION:
                    prefix_duration = max(float(item["end"]) for item in buffer[:index + 1]) - float(buffer[0]["start"])
                    if prefix_duration >= 1.2:
                        fallback_index = index
            flush(fallback_index + 1 if fallback_index is not None else None)
    flush()

    merged: list[dict[str, Any]] = []
    for item in result:
        duration = item["end"] - item["start"]
        if merged:
            previous = merged[-1]
            combined_text = _join_timeline_words([{"word": previous["text"]}, {"word": item["text"]}])
            combined_span = item["end"] - previous["start"]
            gap = item["start"] - previous["end"]
            previous_is_terminal = bool(previous["text"] and previous["text"][-1] in TIMELINE_TERMINAL_PUNCTUATION)
            if (
                duration < 1.2
                and not previous_is_terminal
                and gap <= 1.2
                and combined_span <= TIMELINE_MAX_DURATION_SECONDS
                and _timeline_text_units(combined_text) <= TIMELINE_MAX_TEXT_UNITS
            ):
                previous["end"] = item["end"]
                previous["text"] = combined_text
                previous["source_sentence_ids"] = list(dict.fromkeys(
                    previous["source_sentence_ids"] + item["source_sentence_ids"]
                ))
                continue
        merged.append(item)
    result = merged

    for index, item in enumerate(result):
        item["id"] = index
        if index:
            item["start"] = max(item["start"], result[index - 1]["end"])
        if item["end"] <= item["start"]:
            next_start = result[index + 1]["start"] if index + 1 < len(result) else item["start"] + 0.2
            item["end"] = round(max(item["start"] + 0.05, next_start), 3)
    return result


def _normalize_result(raw: dict[str, Any], task_body: dict[str, Any], task_id: str, source_name: str) -> dict[str, Any]:
    transcripts = raw.get("transcripts")
    if not isinstance(transcripts, list) or not transcripts:
        raise AudioTranscriptionError("转写结果缺少 transcripts")
    full_texts: list[str] = []
    segments: list[dict[str, Any]] = []
    words: list[dict[str, Any]] = []
    for transcript in transcripts:
        if not isinstance(transcript, dict):
            continue
        text = str(transcript.get("text") or transcript.get("transcript") or "").strip()
        if text:
            full_texts.append(text)
        channel_id = transcript.get("channel_id", 0)
        for sentence in transcript.get("sentences") or []:
            if not isinstance(sentence, dict):
                continue
            sentence_id = sentence.get("sentence_id")
            segment = {
                "id": sentence_id,
                "channel_id": channel_id,
                "start": round(float(sentence.get("begin_time", 0)) / 1000, 3),
                "end": round(float(sentence.get("end_time", 0)) / 1000, 3),
                "text": str(sentence.get("text") or "").strip(),
                "language": sentence.get("language"),
                "emotion": sentence.get("emotion"),
            }
            if segment["end"] < segment["start"]:
                raise AudioTranscriptionError("检测到结束时间早于开始时间")
            segments.append(segment)
            for word in sentence.get("words") or []:
                if not isinstance(word, dict):
                    continue
                words.append({
                    "word": f"{word.get('text', '')}{word.get('punctuation') or ''}",
                    "start": round(float(word.get("begin_time", 0)) / 1000, 3),
                    "end": round(float(word.get("end_time", 0)) / 1000, 3),
                    "sentence_id": sentence_id,
                    "channel_id": channel_id,
                })
    segments.sort(key=lambda item: (item["start"], item["end"]))
    words.sort(key=lambda item: (item["start"], item["end"]))
    timeline_segments = _build_timeline_segments(words, segments)
    return {
        "provider": "aliyun-dashscope",
        "model_name": AUDIO_TRANSCRIPTION_MODEL,
        "task_id": task_id,
        "source_name": source_name,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "text": "\n".join(full_texts),
        "segments": timeline_segments,
        "model_segments": segments,
        "timeline_source": "word_timestamps_resegmented" if words else "model_sentences",
        "words": words,
        "usage": task_body.get("usage") if isinstance(task_body.get("usage"), dict) else {},
        "audio_info": raw.get("audio_info") or raw.get("properties") or {},
    }


def _srt_time(seconds: float, separator: str = ",") -> str:
    milliseconds = max(0, round(seconds * 1000))
    hours, remainder = divmod(milliseconds, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    secs, millis = divmod(remainder, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}{separator}{millis:03d}"


def _write_outputs(output_dir: Path, original_filename: str, normalized: dict[str, Any], raw: dict[str, Any]) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "timeline_txt": output_dir / build_user_visible_filename(original_filename, suffix="带时间轴转写", ext=".txt"),
        "plain_txt": output_dir / build_user_visible_filename(original_filename, suffix="纯文本转写", ext=".txt"),
        "srt": output_dir / build_user_visible_filename(original_filename, suffix="字幕", ext=".srt"),
        "vtt": output_dir / build_user_visible_filename(original_filename, suffix="网页字幕", ext=".vtt"),
        "word_tsv": output_dir / build_user_visible_filename(original_filename, suffix="逐词时间戳", ext=".tsv"),
        "result_json": output_dir / build_user_visible_filename(original_filename, suffix="完整转写结果", ext=".json"),
    }
    timeline = [
        f"[{_srt_time(item['start'], '.')} --> {_srt_time(item['end'], '.')}] {item['text']}"
        for item in normalized["segments"]
    ]
    paths["timeline_txt"].write_text("\n".join(timeline) + "\n", encoding="utf-8")
    paths["plain_txt"].write_text(normalized["text"].rstrip() + "\n", encoding="utf-8")
    srt_blocks = [
        f"{index}\n{_srt_time(item['start'])} --> {_srt_time(item['end'])}\n{item['text']}"
        for index, item in enumerate(normalized["segments"], 1)
    ]
    paths["srt"].write_text("\n\n".join(srt_blocks) + "\n", encoding="utf-8")
    vtt_blocks = ["WEBVTT", ""] + [
        f"{_srt_time(item['start'], '.')} --> {_srt_time(item['end'], '.')}\n{item['text']}"
        for item in normalized["segments"]
    ]
    paths["vtt"].write_text("\n\n".join(vtt_blocks) + "\n", encoding="utf-8")
    word_rows = ["start_seconds\tend_seconds\tword\tsentence_id"]
    for item in normalized["words"]:
        clean_word = str(item["word"]).replace("\t", " ").replace("\n", " ")
        word_rows.append(f"{item['start']:.3f}\t{item['end']:.3f}\t{clean_word}\t{item.get('sentence_id', '')}")
    paths["word_tsv"].write_text("\n".join(word_rows) + "\n", encoding="utf-8-sig")
    result_payload = {**normalized, "raw_transcription": raw}
    paths["result_json"].write_text(json.dumps(result_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    archive_path = output_dir / build_user_visible_filename(original_filename, suffix="音频转写结果", ext=".zip")
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in paths.values():
            archive.write(path, arcname=path.name)
    paths["archive_zip"] = archive_path
    return {key: str(path).replace("\\", "/") for key, path in paths.items()}


def _run_audio_transcription(*, display_no: str, input_path: str, original_filename: str, params: dict[str, Any], log_callback=None) -> dict[str, Any]:
    if not settings.DASHSCOPE_API_KEY:
        raise AudioTranscriptionError("未配置 DASHSCOPE_API_KEY")
    validate_audio_transcription_filename(original_filename)
    source_path = Path(input_path)
    timeout = max(30, settings.AUDIO_TRANSCRIPTION_TIMEOUT_SECONDS)
    if log_callback:
        log_callback("[audio-transcription] 上传完整原始音频，不分块、不强制降噪")
    # DashScope 会返回额外的 OSS 上传及结果下载地址；整条链路均使用直连，
    # 避免 HTTP_PROXY/HTTPS_PROXY/ALL_PROXY 导致 TLS 握手被中途关闭。
    with _create_dashscope_session() as session:
        file_url = _upload_temporary_file(session, source_path, timeout)
        task_id = _submit_task(
            session,
            file_url=file_url,
            language=params["language"],
            enable_itn=params["enable_itn"],
            timeout=timeout,
        )
        raw, task_body = _wait_and_download(session, task_id, timeout, log_callback)
    normalized = _normalize_result(raw, task_body, task_id, original_filename)
    if not normalized["text"] and not normalized["segments"]:
        raise AudioTranscriptionError("模型返回了空转写结果")
    output_dir = Path(settings.OUTPUT_DIR) / "audio_transcription" / display_no
    output_paths = _write_outputs(output_dir, original_filename, normalized, raw)
    return {
        **output_paths,
        "model_name": AUDIO_TRANSCRIPTION_MODEL,
        "language": params["language"],
        "enable_itn": params["enable_itn"],
        "text": normalized["text"],
        "segment_count": len(normalized["segments"]),
        "model_segment_count": len(normalized["model_segments"]),
        "word_count": len(normalized["words"]),
        "segments": normalized["segments"][:500],
        "timeline_source": normalized["timeline_source"],
        "timestamps_truncated_in_preview": len(normalized["segments"]) > 500,
        "usage": normalized["usage"],
    }


async def execute_audio_transcription_task(
    *,
    display_no: str,
    input_path: str,
    original_filename: str,
    params: dict[str, Any],
    progress_callback: Callable[[int, str], Any],
    executor: Optional[Executor] = None,
    log_callback=None,
) -> dict[str, Any]:
    await progress_callback(10, "正在上传完整音频")
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(
        executor,
        lambda: _run_audio_transcription(
            display_no=display_no,
            input_path=input_path,
            original_filename=original_filename,
            params=params,
            log_callback=log_callback,
        ),
    )
    await progress_callback(95, "正在生成时间轴文本与字幕文件")
    return result
