# -*- coding: utf-8 -*-
"""
营业执照图片翻译服务

将 businesslicence/start.py 的完整流程桥接到 Web 端。
使用 SSE 推送进度，并在印章验证时暂停等待用户交互。
"""

import asyncio
import builtins
import os
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.core.config import settings

# businesslicence 项目根目录
BL_DIR = Path(__file__).resolve().parent.parent.parent / "businesslicence"
BL_UPLOADS_DIR = BL_DIR / "uploads"
BL_OUTPUTS_DIR = BL_DIR / "outputs"

# 确保目录存在
BL_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
BL_OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

# 将 businesslicence 加入 sys.path（只加一次）
if str(BL_DIR) not in sys.path:
    sys.path.insert(0, str(BL_DIR))

# 任务状态存储  task_id -> dict
_tasks: dict = {}

# ---------------------------------------------------------------------------
# 线程感知的 builtins.input 补丁
# 模仿 start.py 的 InteractiveInputCapture，但适配 Web 多线程环境。
# 在流水线线程内的 input() 调用（来自 SealTextHandler 等）不会阻塞进程。
# ---------------------------------------------------------------------------
_original_input = builtins.input
_pipeline_threads: dict = {}   # thread_id -> task_id
_pipeline_threads_lock = threading.Lock()


def _global_input_patch(prompt=""):
    """
    全局 input() 替换函数。
    - 流水线线程内：拦截验证提示，通过 SSE 通知前端；其他提示自动返回空。
    - 其他线程：回退到原始 input()。
    """
    thread_id = threading.current_thread().ident
    with _pipeline_threads_lock:
        task_id = _pipeline_threads.get(thread_id)

    if task_id is None:
        # 非流水线线程，正常处理
        return _original_input(prompt)

    prompt_str = str(prompt)

    # SealTextHandler / 终端模式 的验证提示
    if "识别是否正确" in prompt_str:
        task = _tasks.get(task_id)
        if task:
            loop = task.get("_loop")
            # 只在 SSE 已连接且任务运行中时才弹验证弹窗
            if loop and task.get("status") == "running":
                # 推送一个轻量级 SSE 事件（此处 region_info 由 _gui_verification_callback 处理，
                # 这里作为兜底：如果到达此处说明 gui_callback 未拦截，直接自动确认）
                task["log_lines"].append("[自动确认] SealTextHandler 验证已自动确认")
                asyncio.run_coroutine_threadsafe(
                    task["sse_queue"].put({"type": "log", "text": "[自动确认] 印章文字已自动确认"}),
                    loop,
                )
        return "y"

    if "请输入正确的文字内容" in prompt_str:
        return ""

    # 其他 input() 调用（如 "按回车键退出"）：直接返回空，不阻塞
    return ""


# 在模块加载时全局替换一次
builtins.input = _global_input_patch


def _make_task(task_id: str):
    """初始化任务状态"""
    _tasks[task_id] = {
        "status": "pending",          # pending / running / waiting_verify / done / error
        "progress": 0,
        "log_lines": [],              # 实时日志行列表
        "sse_queue": None,            # asyncio.Queue，用于向 SSE 生成器发送事件
        "verify_event": None,         # threading.Event，用于阻塞流水线线程
        "verify_result": None,        # (action, text) 用户验证结果
        "output_path": None,
        "input_filename": None,
        "error": None,
        "_loop": None,                # asyncio 事件循环（供 _global_input_patch 使用）
    }


def get_task(task_id: str) -> Optional[dict]:
    return _tasks.get(task_id)


# ---------------------------------------------------------------------------
# 同步回调：在流水线线程中被调用，需要阻塞等待 Web 用户响应
# 对应 start.py 的 show_verification_dialog，覆盖主流水线 _gui_verification_callback 路径
# ---------------------------------------------------------------------------

def _make_sync_verification_callback(task_id: str, loop: asyncio.AbstractEventLoop):
    """
    返回一个同步的印章验证回调函数，与 start.py 的 show_verification_dialog 签名一致。
    在流水线线程内调用，会阻塞线程直到用户通过 Web 端提交决定。
    """
    def callback(region_info: dict):
        task = _tasks.get(task_id)
        if not task:
            return ("confirm", None)

        # 把验证请求推送到 SSE 队列（线程安全）
        event_data = {
            "type": "verification_request",
            "region_info": {
                "type_name": region_info.get("type_name", ""),
                "text_type": region_info.get("text_type", ""),
                "bbox": list(region_info.get("bbox", [])),
                "text": region_info.get("text", ""),
                "confidence": float(region_info.get("confidence", 0)),
                "index": region_info.get("index", 1),
                "total": region_info.get("total", 1),
            }
        }
        asyncio.run_coroutine_threadsafe(
            task["sse_queue"].put(event_data),
            loop
        )

        # 更新任务状态
        task["status"] = "waiting_verify"
        task["verify_event"] = threading.Event()
        task["verify_result"] = None

        # 阻塞当前线程，等待用户通过 /verify 接口提交结果
        task["verify_event"].wait(timeout=300)  # 最多等 5 分钟

        result = task.get("verify_result") or ("confirm", None)
        task["verify_event"] = None
        task["status"] = "running"
        return result

    return callback


# ---------------------------------------------------------------------------
# 流水线线程函数
# ---------------------------------------------------------------------------

def _run_pipeline_in_thread(task_id: str, input_path: str, output_path: str,
                             source_lang: str, target_lang: str,
                             config_file: str, loop: asyncio.AbstractEventLoop):
    """在独立线程中运行流水线，完成后向 SSE 队列发送 done/error 事件。"""
    thread_id = threading.current_thread().ident

    # 注册本线程为流水线线程，使 _global_input_patch 生效
    with _pipeline_threads_lock:
        _pipeline_threads[thread_id] = task_id

    task = _tasks.get(task_id)
    if not task:
        with _pipeline_threads_lock:
            _pipeline_threads.pop(thread_id, None)
        return

    # 将事件循环存入 task，供 _global_input_patch 使用
    task["_loop"] = loop

    def push(event: dict):
        """线程安全地向 SSE 队列推送事件"""
        asyncio.run_coroutine_threadsafe(task["sse_queue"].put(event), loop)

    def log(text: str):
        """追加日志并推送"""
        task["log_lines"].append(text)
        push({"type": "log", "text": text})

    def progress(pct: int, message: str = ""):
        task["progress"] = pct
        push({"type": "progress", "progress": pct, "message": message})

    try:
        task["status"] = "running"
        progress(5, "初始化配置...")
        log("=" * 60)
        log("营业执照翻译系统启动")
        log("=" * 60)

        # 设置 API Key
        os.environ["DEEPSEEK_API_KEY"] = settings.DEEPSEEK_API_KEY

        # 加载配置
        from src.config.config_manager import ConfigManager
        config_path = str(BL_DIR / config_file)
        config = ConfigManager(config_path=config_path)
        errors = config.validate()
        if errors:
            raise RuntimeError("配置验证失败: " + "; ".join(errors))

        log(f"配置文件: {config_file}")
        log(f"文档方向: {config.get_orientation()}")
        log(f"源语言: {source_lang}  目标语言: {target_lang}")
        progress(10, "配置加载完成，初始化流水线...")

        # 初始化流水线
        from src.pipeline.translation_pipeline import TranslationPipeline
        pipeline = TranslationPipeline(config)

        # 设置 Web 版验证回调（替代 Tkinter GUI，覆盖主流水线验证路径）
        pipeline._gui_verification_callback = _make_sync_verification_callback(task_id, loop)

        log(f"输入图片: {Path(input_path).name}")
        log(f"输出路径: {Path(output_path).name}")
        progress(15, "开始处理图片...")

        # 执行翻译
        pipeline.translate_image(
            input_path,
            output_path,
            source_lang=source_lang,
            target_lang=target_lang
        )

        progress(95, "图片处理完成，保存结果...")

        if not Path(output_path).exists():
            raise RuntimeError("翻译完成但未找到输出文件")

        # 添加右下角水印（尽力而为，失败不影响主流程）
        log("添加水印...")
        try:
            from app.service.image_processor import add_watermark
            add_watermark(output_path, position="top_right")
            log("水印已添加（右上角）")
        except Exception as wm_err:
            log(f"⚠️ 水印添加失败（已跳过）: {wm_err}")

        log("=" * 60)
        log(f"翻译完成！输出文件: {Path(output_path).name}")
        log("=" * 60)

        task["status"] = "done"
        task["output_path"] = output_path
        progress(100, "完成")

        push({
            "type": "done",
            "output_filename": Path(output_path).name,
            "input_filename": task["input_filename"],
        })

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        log(f"错误: {e}")
        log(tb)
        task["status"] = "error"
        task["error"] = str(e)
        push({"type": "error", "message": str(e)})

    finally:
        # 注销本线程，恢复 input() 正常行为
        with _pipeline_threads_lock:
            _pipeline_threads.pop(thread_id, None)


# ---------------------------------------------------------------------------
# 公共接口
# ---------------------------------------------------------------------------

async def start_business_licence_task(
    file_bytes: bytes,
    original_filename: str,
    source_lang: str,
    target_lang: str,
    config_file: str,
) -> str:
    """
    保存上传文件，启动后台线程运行流水线。
    返回 task_id。
    """
    task_id = str(uuid.uuid4())
    _make_task(task_id)

    # 创建运行时专属的 asyncio 队列（在当前事件循环中）
    _tasks[task_id]["sse_queue"] = asyncio.Queue()
    _tasks[task_id]["input_filename"] = original_filename

    # 保存上传文件
    suffix = Path(original_filename).suffix or ".jpg"
    input_path = str(BL_UPLOADS_DIR / f"{task_id}{suffix}")
    with open(input_path, "wb") as f:
        f.write(file_bytes)

    # 生成输出文件名：使用 task_id（纯 ASCII UUID）作为文件名主体，
    # 避免 OpenCV 在 Windows 下无法正确处理非 ASCII 路径的问题。
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"{task_id}_translated_{timestamp}{suffix}"
    output_path = str(BL_OUTPUTS_DIR / output_filename)

    # 获取当前事件循环，传给线程
    loop = asyncio.get_event_loop()

    t = threading.Thread(
        target=_run_pipeline_in_thread,
        args=(task_id, input_path, output_path, source_lang, target_lang, config_file, loop),
        daemon=True,
    )
    t.start()

    return task_id


async def stream_task_events(task_id: str):
    """
    SSE 生成器：从 sse_queue 读取事件并以 text/event-stream 格式 yield。
    """
    import json

    task = _tasks.get(task_id)
    if not task:
        yield f"data: {json.dumps({'type': 'error', 'message': '任务不存在'})}\n\n"
        return

    queue: asyncio.Queue = task["sse_queue"]

    while True:
        try:
            event = await asyncio.wait_for(queue.get(), timeout=30)
        except asyncio.TimeoutError:
            # 心跳，防止连接超时
            yield ": heartbeat\n\n"
            if task["status"] in ("done", "error"):
                break
            continue

        yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"

        if event.get("type") in ("done", "error"):
            break


def submit_verification(task_id: str, action: str, text: Optional[str]) -> bool:
    """
    用户提交印章验证结果，唤醒流水线线程。
    返回 True 表示成功，False 表示任务不存在或未在等待验证。
    """
    task = _tasks.get(task_id)
    if not task or task.get("status") != "waiting_verify":
        return False

    task["verify_result"] = (action, text)
    if task.get("verify_event"):
        task["verify_event"].set()
    return True
