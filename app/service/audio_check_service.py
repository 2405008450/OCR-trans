from __future__ import annotations

import asyncio
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Callable, Optional

from markdown_it import MarkdownIt

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename
from app.service.gemini_service import generate_audio_analysis

AUDIO_CHECK_ALLOWED_EXTENSIONS = {".wav", ".mp3", ".aiff", ".aac", ".ogg", ".flac"}
AUDIO_CHECK_MIME_TYPES = {
    ".wav": "audio/wav",
    ".mp3": "audio/mp3",
    ".aiff": "audio/aiff",
    ".aac": "audio/aac",
    ".ogg": "audio/ogg",
    ".flac": "audio/flac",
}
AUDIO_CHECK_MODELS = {
    "google/gemini-3.6-flash": {
        "label": "Gemini 3.6 Flash",
        "description": "默认模型，速度与音频理解能力均衡。",
    },
    "google/gemini-3.5-flash": {
        "label": "Gemini 3.5 Flash",
        "description": "适合常规音质检查。",
    },
    "google/gemini-3.1-pro-preview": {
        "label": "Gemini 3.1 Pro",
        "description": "适合复杂音频和更细致的综合判断。",
    },
}
AUDIO_CHECK_DEFAULT_MODEL = "google/gemini-3.6-flash"
AUDIO_CHECK_DEFAULT_TEMPERATURE = 0.1
AUDIO_CHECK_DEFAULT_MAX_OUTPUT_TOKENS = 32768

AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT = """你是一名严谨的音频质量检测工程师。你将收到用户上传的未经转码、剪切、降噪或响度处理的原始音频。

请重点评估响度、底噪、静音、爆音或削波风险、失真、声道异常、语音清晰度、异常声音和影响使用的技术问题。
必须明确区分：
1. 从原始音频中能够明确听到的现象；
2. 只能通过听感推测、无法精确测量的风险；
3. 当前音频和模型能力无法确认的事项。

不要伪造响度、峰值、频率或其他精确测量数值，不要把主观听感描述成仪器测量结果。使用中文 Markdown 输出，结论清晰、可执行。"""

AUDIO_CHECK_DEFAULT_USER_PROMPT = """请对该音频执行技术质量检查，并按以下结构输出：

## 总体结论
给出合格、建议复查或不合格，并简述主要原因。

## 主要问题
按严重程度列出问题；如能定位，请给出大致时间段。

## 音质分析
分析响度、底噪、静音、失真、削波风险、清晰度、声道和异常声音。

## 判断依据
说明结论来自哪些可听见现象，并标明无法精确测量或需要人工复核的部分。

## 改进建议
给出可操作的录制、剪辑或后期处理建议。

若未发现明显问题，也要明确说明检查范围和仍存在的不确定性。"""

LogCallback = Optional[Callable[[str], None]]
_MARKDOWN = MarkdownIt("commonmark", {"html": False, "linkify": False, "typographer": False})


def get_audio_check_config() -> dict[str, Any]:
    return {
        "models": AUDIO_CHECK_MODELS,
        "default_model": AUDIO_CHECK_DEFAULT_MODEL,
        "allowed_extensions": sorted(AUDIO_CHECK_ALLOWED_EXTENSIONS),
        "max_file_mb": settings.AUDIO_CHECK_MAX_MB,
        "default_system_prompt": AUDIO_CHECK_DEFAULT_SYSTEM_PROMPT,
        "default_user_prompt": AUDIO_CHECK_DEFAULT_USER_PROMPT,
        "defaults": {
            "temperature": AUDIO_CHECK_DEFAULT_TEMPERATURE,
            "max_output_tokens": AUDIO_CHECK_DEFAULT_MAX_OUTPUT_TOKENS,
            "timeout_seconds": max(30, min(600, int(settings.GEMINI_TIMEOUT_SECONDS))),
        },
        "ranges": {
            "temperature": {"min": 0, "max": 2, "step": 0.1},
            "max_output_tokens": {"min": 1024, "max": 65536},
            "timeout_seconds": {"min": 30, "max": 600},
        },
    }


def normalize_audio_check_options(
    *,
    model_name: str,
    gemini_route: str,
    system_prompt: str,
    user_prompt: str,
    temperature: float,
    max_output_tokens: int,
    timeout_seconds: int,
) -> dict[str, Any]:
    model = (model_name or "").strip()
    if model not in AUDIO_CHECK_MODELS:
        raise ValueError("不支持的音频检查模型")
    route = (gemini_route or "").strip().lower()
    if route not in {"google", "aistudio", "openrouter"}:
        raise ValueError("不支持的模型调用线路")
    normalized_system_prompt = (system_prompt or "").strip()
    normalized_user_prompt = (user_prompt or "").strip()
    if not normalized_system_prompt:
        raise ValueError("系统提示词不能为空")
    if not normalized_user_prompt:
        raise ValueError("本次检查要求不能为空")
    if len(normalized_system_prompt) > 20_000:
        raise ValueError("系统提示词不能超过 20000 个字符")
    if len(normalized_user_prompt) > 10_000:
        raise ValueError("本次检查要求不能超过 10000 个字符")
    if not 0 <= float(temperature) <= 2:
        raise ValueError("温度必须在 0 到 2 之间")
    if not 1024 <= int(max_output_tokens) <= 65536:
        raise ValueError("最大输出 token 必须在 1024 到 65536 之间")
    if not 30 <= int(timeout_seconds) <= 600:
        raise ValueError("超时必须在 30 到 600 秒之间")
    return {
        "model_name": model,
        "gemini_route": route,
        "system_prompt": normalized_system_prompt,
        "user_prompt": normalized_user_prompt,
        "temperature": float(temperature),
        "max_output_tokens": int(max_output_tokens),
        "timeout_seconds": int(timeout_seconds),
    }


def validate_audio_filename(filename: str) -> tuple[str, str]:
    extension = Path(filename or "").suffix.lower()
    if extension not in AUDIO_CHECK_ALLOWED_EXTENSIONS:
        supported = "、".join(sorted(AUDIO_CHECK_ALLOWED_EXTENSIONS))
        raise ValueError(f"不支持的音频格式，仅支持：{supported}")
    return extension, AUDIO_CHECK_MIME_TYPES[extension]


def render_markdown(markdown_text: str) -> str:
    return _MARKDOWN.render(markdown_text or "")


def _run_audio_check(
    *,
    display_no: str,
    input_path: str,
    original_filename: str,
    params: dict[str, Any],
    log_callback: LogCallback,
) -> dict[str, Any]:
    extension, mime_type = validate_audio_filename(original_filename)
    if log_callback:
        log_callback("[audio-check] 正在读取原始音频，不执行转码或音频预处理")
    audio_bytes = Path(input_path).read_bytes()
    analysis_markdown = generate_audio_analysis(
        system_prompt=params["system_prompt"],
        user_prompt=params["user_prompt"],
        audio_bytes=audio_bytes,
        mime_type=mime_type,
        audio_format=extension.lstrip("."),
        model=params["model_name"],
        route=params["gemini_route"],
        temperature=params["temperature"],
        max_output_tokens=params["max_output_tokens"],
        timeout=params["timeout_seconds"],
        log_callback=log_callback,
    ).strip()
    if not analysis_markdown:
        raise RuntimeError("分析模型未返回音频检查结果")

    report = (
        f"# 音频质量检查报告\n\n"
        f"- 文件：{original_filename}\n"
        f"- 模型：{params['model_name']}\n"
        f"- 线路：{params['gemini_route']}\n\n"
        f"> 检查对象为用户上传的原始音频；系统未执行转码、剪切、降噪、响度归一化或指标预处理。\n\n"
        f"## 原始音频检查结果\n\n{analysis_markdown}\n"
    )
    output_dir = Path(settings.OUTPUT_DIR) / "audio_check" / display_no
    output_dir.mkdir(parents=True, exist_ok=True)
    output_name = build_user_visible_filename(original_filename, suffix="音频质量检查报告", ext=".md")
    output_path = output_dir / output_name
    output_path.write_text(report, encoding="utf-8")
    return {
        "analysis_markdown": analysis_markdown,
        "analysis_html": render_markdown(analysis_markdown),
        "report_markdown": str(output_path).replace("\\", "/"),
        "model_name": params["model_name"],
        "gemini_route": params["gemini_route"],
    }


async def execute_audio_check_task(
    *,
    display_no: str,
    input_path: str,
    original_filename: str,
    params: dict[str, Any],
    progress_callback: Callable[[int, str], Any],
    executor: Optional[Executor] = None,
    log_callback: LogCallback = None,
) -> dict[str, Any]:
    await progress_callback(10, "正在上传原始音频，不进行转码或预处理")
    loop = asyncio.get_running_loop()

    def run() -> dict[str, Any]:
        return _run_audio_check(
            display_no=display_no,
            input_path=input_path,
            original_filename=original_filename,
            params=params,
            log_callback=log_callback,
        )

    result = await loop.run_in_executor(executor, run)
    await progress_callback(95, "正在生成 Markdown 报告")
    return result
