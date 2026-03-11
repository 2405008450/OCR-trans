import os
import uuid
import asyncio
import threading
from fastapi import UploadFile
from typing import Dict, Any, Optional
from app.core.config import settings
from app.service.image_processor import process_image, convert_input_to_images
from app.service.marriage_cert_processor import process_marriage_cert_image

# 在文件顶部创建一个全局锁，限制同时只能执行1个耗时的 AI 处理任务
ai_process_lock = asyncio.Lock()

os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)
os.makedirs(settings.TEMP_IMAGES_DIR, exist_ok=True)

# OCR 任务状态字典：task_id -> {"status": "processing"|"done"|"error", "result": ..., "error": ...}
_ocr_tasks: Dict[str, Dict[str, Any]] = {}
_ocr_tasks_lock = threading.Lock()


def get_ocr_task_status(task_id: str) -> Optional[Dict[str, Any]]:
    with _ocr_tasks_lock:
        return _ocr_tasks.get(task_id)

async def run_llm_task(
    file: UploadFile,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = 'front',
    doc_type: str = 'id_card',
    marriage_page_template: str = 'page2',
    registrar_signature_text: Optional[str] = None,
    registered_by_text: Optional[str] = None,
    registered_by_offset_x: int = 0,
    registered_by_offset_y: int = 0,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12,
    enable_merge: bool = True,
    enable_overlap_fix: bool = True,
    enable_colon_fix: bool = False,
    font_size: Optional[int] = None,
) -> Dict[str, Any]:
    """
    处理上传的图片文件
    
    Args:
        file: 上传的文件
        from_lang: 源语言，默认'zh'
        to_lang: 目标语言，默认'en'
        enable_correction: 是否启用透视矫正（已停用，仅兼容旧参数）
        enable_visualization: 是否生成可视化图片
        card_side: 证件面，'front'=正面，'back'=背面（仅身份证有效）
        doc_type: 证件类型，'id_card'=身份证，'marriage_cert'=结婚证
        marriage_page_template: [结婚证] 模板页，'page1' / 'page2' / 'page3'
        registrar_signature_text: [结婚证] 婚姻登记员手写签名（手动输入）
        registered_by_text: [结婚证] Registered by中的xxx（手动输入）
        registered_by_offset_x: [结婚证] Registered by 右移偏移(px)
        registered_by_offset_y: [结婚证] Registered by 纵向偏移(px)
        registrar_signature_offset_x: [结婚证] 婚姻登记员手写签名右移偏移(px)
        registrar_signature_offset_y: [结婚证] 婚姻登记员手写签名纵向偏移(px)
        enable_merge: [结婚证] 框体合并开关
        enable_overlap_fix: [结婚证] 重叠修正开关
        enable_colon_fix: [结婚证] 冒号修正开关
        font_size: [结婚证] 字体大小（像素）
    
    Returns:
        包含处理结果的字典
    """
    task_id = str(uuid.uuid4())
    
    # 1. 保存上传文件（使用 UUID 风格文件名，避免中文路径导致 cv2 等库无法处理）
    file_ext = os.path.splitext(file.filename)[1].lower() or ".jpg"
    safe_filename = f"{task_id}{file_ext}"
    input_path = os.path.join(settings.UPLOAD_DIR, safe_filename)
    
    with open(input_path, "wb") as f:
        content = await file.read()
        f.write(content)
    
    # 2. 转换输入（如果是PDF则转为图片）
    image_paths = convert_input_to_images(input_path, settings.TEMP_IMAGES_DIR)
    
    if not image_paths:
        raise ValueError(f"不支持的文件格式: {file_ext}")
    
    # 3. 处理每张图片（如果是PDF可能有多页）
    results = []
    for idx, img_path in enumerate(image_paths):
        print(f"\n处理第 {idx + 1}/{len(image_paths)} 张图片...")
        
        # 根据证件类型路由到不同处理器
        if doc_type == 'marriage_cert':
            template = (marriage_page_template or 'page2').lower().strip()
            confidence_threshold = 0.5  # 默认置信度阈值

            if template == 'page1':
                # 第一页（封面）模板
                enable_merge = True
                enable_overlap_fix = True
                enable_colon_fix = False
                confidence_threshold = 0.8  # 第一页只保留高置信度
            elif template == 'page2':
                # 第二页模板
                enable_merge = False
                enable_overlap_fix = True
                enable_colon_fix = True
            elif template == 'page3':
                # 第三页模板
                enable_merge = True
                enable_overlap_fix = True
                enable_colon_fix = False

            # 手动输入元素仅第一页生效，第二/三页强制忽略
            if template != 'page1':
                registrar_signature_text = None
                registered_by_text = None
                registered_by_offset_x = 0
                registered_by_offset_y = 0
                registrar_signature_offset_x = 36
                registrar_signature_offset_y = -12

            print(
                f"结婚证模板: {template} | "
                f"merge={enable_merge}, overlap_fix={enable_overlap_fix}, "
                f"colon_fix={enable_colon_fix}, confidence={confidence_threshold}"
            )

            # 使用全局锁进行排队（防止双并发耗尽物理内存），且放到线程池执行（避免阻塞主事件循环导致服务假死）
            async with ai_process_lock:
                from fastapi.concurrency import run_in_threadpool
                result = await run_in_threadpool(
                    process_marriage_cert_image,
                    input_path=img_path,
                    output_dir=settings.OUTPUT_DIR,
                    from_lang=from_lang,
                    to_lang=to_lang,
                    enable_correction=enable_correction,
                    enable_visualization=enable_visualization,
                    enable_merge=enable_merge,
                    enable_overlap_fix=enable_overlap_fix,
                    enable_colon_fix=enable_colon_fix,
                    font_size=font_size if font_size else 18,
                    confidence_threshold=confidence_threshold,
                    page_template=template,
                    registrar_signature_text=registrar_signature_text,
                    registered_by_text=registered_by_text,
                    registered_by_offset_x=registered_by_offset_x,
                    registered_by_offset_y=registered_by_offset_y,
                    registrar_signature_offset_x=registrar_signature_offset_x,
                    registrar_signature_offset_y=registrar_signature_offset_y,
                )
        else:
            # 默认：身份证处理
            async with ai_process_lock:
                from fastapi.concurrency import run_in_threadpool
                result = await run_in_threadpool(
                    process_image,
                    input_path=img_path,
                    output_dir=settings.OUTPUT_DIR,
                    from_lang=from_lang,
                    to_lang=to_lang,
                    enable_correction=enable_correction,
                    enable_visualization=enable_visualization,
                    card_side=card_side
                )
        
        # 转换字段名称以匹配前端期望，并标准化路径格式
        def normalize_path(path):
            """将路径转换为URL格式（使用正斜杠）"""
            if path:
                # 将反斜杠转换为正斜杠
                return path.replace("\\", "/")
            return None
        
        formatted_result = {
            # 透视矫正步骤已停用，前端不再展示“矫正后的图片”
            "corrected_image": None,
            "visualization_image": normalize_path(result.get("visualization")),
            "translated_image": normalize_path(result.get("final_output")),
            "ocr_json": normalize_path(result.get("raw_ocr_json")),
            "translated_json": normalize_path(result.get("translated_json"))
        }
        
        # 调试日志
        print(f"\n格式化后的结果:")
        print(f"   - 翻译图片: {formatted_result['translated_image']}")
        print(f"   - 可视化图片: {formatted_result['visualization_image']}")
        print(f"   - OCR JSON: {formatted_result['ocr_json']}")
        
        results.append(formatted_result)
    
    # 4. 返回结果
    return {
        "task_id": task_id,
        "filename": file.filename,
        "results": results,
        "total_images": len(image_paths)
    }


async def submit_ocr_task(
    file: UploadFile,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = 'front',
    doc_type: str = 'id_card',
    marriage_page_template: str = 'page2',
    registrar_signature_text: Optional[str] = None,
    registered_by_text: Optional[str] = None,
    registered_by_offset_x: int = 0,
    registered_by_offset_y: int = 0,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12,
    enable_merge: bool = True,
    enable_overlap_fix: bool = True,
    enable_colon_fix: bool = False,
    font_size: Optional[int] = None,
) -> str:
    """
    异步提交 OCR 任务，立即返回 task_id，任务在后台执行。
    通过 get_ocr_task_status(task_id) 查询结果。
    """
    task_id = str(uuid.uuid4())

    # 先在当前协程中读取文件内容（UploadFile 只能在事件循环中读取）
    file_ext = os.path.splitext(file.filename)[1].lower() or ".jpg"
    safe_filename = f"{task_id}{file_ext}"
    input_path = os.path.join(settings.UPLOAD_DIR, safe_filename)
    content = await file.read()
    original_filename = file.filename

    with open(input_path, "wb") as f:
        f.write(content)

    # 记录任务为排队中（等待 ai_process_lock）
    with _ocr_tasks_lock:
        _ocr_tasks[task_id] = {"status": "queued", "result": None, "error": None}

    # 获取当前事件循环，供后台线程回写结果
    loop = asyncio.get_event_loop()

    def _run_in_thread():
        import asyncio as _asyncio

        async def _async_task():
            from fastapi.concurrency import run_in_threadpool
            # convert_input_to_images 是同步 IO，必须放到 threadpool 避免阻塞事件循环
            image_paths = await run_in_threadpool(convert_input_to_images, input_path, settings.TEMP_IMAGES_DIR)
            if not image_paths:
                raise ValueError(f"不支持的文件格式: {file_ext}")

            results = []
            for idx, img_path in enumerate(image_paths):
                print(f"\n[后台任务 {task_id}] 处理第 {idx + 1}/{len(image_paths)} 张图片...")

                if doc_type == 'marriage_cert':
                    template = (marriage_page_template or 'page2').lower().strip()
                    confidence_threshold = 0.5
                    _em, _eo, _ec = enable_merge, enable_overlap_fix, enable_colon_fix
                    _rs, _rb = registrar_signature_text, registered_by_text
                    _rbx, _rby = registered_by_offset_x, registered_by_offset_y
                    _rsx, _rsy = registrar_signature_offset_x, registrar_signature_offset_y

                    if template == 'page1':
                        _em, _eo, _ec = True, True, False
                        confidence_threshold = 0.8
                    elif template == 'page2':
                        _em, _eo, _ec = False, True, True
                    elif template == 'page3':
                        _em, _eo, _ec = True, True, False

                    if template != 'page1':
                        _rs = _rb = None
                        _rbx = _rby = 0
                        _rsx, _rsy = 36, -12

                    async with ai_process_lock:
                        # 获得锁后（真正开始处理时）才将状态更新为 processing
                        if idx == 0:
                            with _ocr_tasks_lock:
                                _ocr_tasks[task_id]["status"] = "processing"
                        result = await run_in_threadpool(
                            process_marriage_cert_image,
                            input_path=img_path,
                            output_dir=settings.OUTPUT_DIR,
                            from_lang=from_lang,
                            to_lang=to_lang,
                            enable_correction=enable_correction,
                            enable_visualization=enable_visualization,
                            enable_merge=_em,
                            enable_overlap_fix=_eo,
                            enable_colon_fix=_ec,
                            font_size=font_size if font_size else 18,
                            confidence_threshold=confidence_threshold,
                            page_template=template,
                            registrar_signature_text=_rs,
                            registered_by_text=_rb,
                            registered_by_offset_x=_rbx,
                            registered_by_offset_y=_rby,
                            registrar_signature_offset_x=_rsx,
                            registrar_signature_offset_y=_rsy,
                        )
                else:
                    async with ai_process_lock:
                        # 获得锁后（真正开始处理时）才将状态更新为 processing
                        if idx == 0:
                            with _ocr_tasks_lock:
                                _ocr_tasks[task_id]["status"] = "processing"
                        result = await run_in_threadpool(
                            process_image,
                            input_path=img_path,
                            output_dir=settings.OUTPUT_DIR,
                            from_lang=from_lang,
                            to_lang=to_lang,
                            enable_correction=enable_correction,
                            enable_visualization=enable_visualization,
                            card_side=card_side,
                        )

                def normalize_path(path):
                    return path.replace("\\", "/") if path else None

                formatted_result = {
                    "corrected_image": None,
                    "visualization_image": normalize_path(result.get("visualization")),
                    "translated_image": normalize_path(result.get("final_output")),
                    "ocr_json": normalize_path(result.get("raw_ocr_json")),
                    "translated_json": normalize_path(result.get("translated_json")),
                }
                print(f"[后台任务 {task_id}] 翻译图片: {formatted_result['translated_image']}")
                results.append(formatted_result)

            return {
                "task_id": task_id,
                "filename": original_filename,
                "results": results,
                "total_images": len(image_paths),
            }

        try:
            # 在主事件循环上运行异步任务（与 ai_process_lock 共享同一循环）
            future = _asyncio.run_coroutine_threadsafe(_async_task(), loop)
            result_data = future.result()
            with _ocr_tasks_lock:
                _ocr_tasks[task_id] = {"status": "done", "result": result_data, "error": None}
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            print(f"[后台任务 {task_id}] 失败:\n{tb}")
            with _ocr_tasks_lock:
                _ocr_tasks[task_id] = {"status": "error", "result": None, "error": str(e)}

    thread = threading.Thread(target=_run_in_thread, daemon=True)
    thread.start()

    return task_id
