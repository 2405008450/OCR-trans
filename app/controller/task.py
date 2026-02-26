import os
import traceback
import asyncio
from fastapi import APIRouter, UploadFile, File, Query, HTTPException, BackgroundTasks
from app.service.llm_service import run_llm_task
from app.service.number_check_service import run_number_check_task, _get_task_progress
from app.service.alignment_service import (
    run_alignment_task, get_alignment_progress, _complete_task as _alignment_complete_task,
    AVAILABLE_MODELS as ALIGNMENT_MODELS, SUPPORTED_LANGUAGES, THRESHOLD_MAP, BUFFER_CHARS,
)
from typing import Optional

router = APIRouter(prefix="/task", tags=["Task"])

@router.post("/run")
async def run_task(
    file: UploadFile = File(..., description="上传的图片文件"),
    from_lang: str = Query("zh", description="源语言，默认'zh'"),
    to_lang: str = Query("en", description="目标语言，默认'en'"),
    enable_correction: bool = Query(False, description="是否启用透视矫正（已停用，仅兼容旧参数）"),
    enable_visualization: bool = Query(True, description="是否生成可视化图片"),
    card_side: str = Query("front", description="[身份证] 证件面: 'front'=正面, 'back'=背面"),
    doc_type: str = Query("id_card", description="证件类型: 'id_card'=身份证, 'marriage_cert'=结婚证"),
    marriage_page_template: str = Query("page2", description="[结婚证] 模板页: 'page1'=第一页, 'page2'=第二页, 'page3'=第三页"),
    registrar_signature_text: Optional[str] = Query(None, description="[结婚证] 婚姻登记员手写签名(手动输入)，用于右侧补填"),
    registered_by_text: Optional[str] = Query(None, description="[结婚证] Registered by中的xxx(手动输入)，用于签名上方补填"),
    registered_by_offset_x: int = Query(20, description="[结婚证] Registered by右移偏移(px)，负数左移，正数右移"),
    registered_by_offset_y: int = Query(-80, description="[结婚证] Registered by纵向偏移(px)，负数上移，正数下移"),
    registrar_signature_offset_x: int = Query(48, description="[结婚证] 婚姻登记员手写签名右移偏移(px)，值越大越靠右"),
    registrar_signature_offset_y: int = Query(-12, description="[结婚证] 婚姻登记员手写签名纵向偏移(px)，负数上移，正数下移"),
    enable_merge: bool = Query(True, description="[结婚证] 框体合并: True=连续文本框合并翻译, False=每个框单独翻译"),
    enable_overlap_fix: bool = Query(True, description="[结婚证] 重叠修正: True=自动检测并右移重叠框, False=保持原位"),
    enable_colon_fix: bool = Query(True, description="[结婚证] 冒号修正: True=为字段名自动添加冒号, False=保持原样"),
    font_size: Optional[int] = Query(None, description="[结婚证] 字体大小(px)，默认18，建议范围8-30"),
):
    """
    处理图片文件，进行OCR识别和翻译
    
    支持的证件类型：
    - 身份证（id_card）：正面/背面处理，智能分割、地址合并、字段映射
    - 结婚证（marriage_cert）：文本合并、重叠修正、冒号修正、智能翻译
    
    通用功能：
    - OCR文字识别（PaddleOCR）
    - 多语言翻译（DeepSeek API）
    - 图像修复（LaMa / OpenCV）
    - 透视矫正（可选）
    - 可视化结果（可选）
    """
    try:
        result = await run_llm_task(
            file=file,
            from_lang=from_lang,
            to_lang=to_lang,
            enable_correction=enable_correction,
            enable_visualization=enable_visualization,
            card_side=card_side,
            doc_type=doc_type,
            marriage_page_template=marriage_page_template,
            registrar_signature_text=registrar_signature_text,
            registered_by_text=registered_by_text,
            registered_by_offset_x=registered_by_offset_x,
            registered_by_offset_y=registered_by_offset_y,
            registrar_signature_offset_x=registrar_signature_offset_x,
            registrar_signature_offset_y=registrar_signature_offset_y,
            enable_merge=enable_merge,
            enable_overlap_fix=enable_overlap_fix,
            enable_colon_fix=enable_colon_fix,
            font_size=font_size,
        )
        return {
            "status": "DONE",
            "task_id": result["task_id"],
            "filename": result["filename"],
            "total_images": result["total_images"],
            "results": result["results"]
        }
    except Exception as e:
        tb = traceback.format_exc()
        print("=" * 60)
        print("任务执行失败:")
        print(tb)
        print("=" * 60)
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(e),
                "type": type(e).__name__,
                "traceback": tb.split("\n")[-10:] if tb else []  # 最后10行便于排查
            }
        )


@router.post("/number-check")
async def run_number_check(
    background_tasks: BackgroundTasks,
    original_file: UploadFile = File(..., description="原文 docx"),
    translated_file: UploadFile = File(..., description="译文 docx"),
):
    """
    数值专检：上传原文/译文 docx，生成修复后的译文与对比报告。
    任务在后台异步执行，返回 task_id 用于查询进度和结果。
    """
    import uuid
    from pathlib import Path
    from app.core.config import settings

    task_id = str(uuid.uuid4())

    # 先保存文件
    upload_dir = Path(settings.UPLOAD_DIR) / "number_check"
    upload_dir.mkdir(parents=True, exist_ok=True)

    original_path = upload_dir / f"{task_id}_original.docx"
    translated_path = upload_dir / f"{task_id}_translated.docx"

    # 读取并保存文件内容（需要await）
    original_content = await original_file.read()
    translated_content = await translated_file.read()

    with open(original_path, "wb") as f:
        f.write(original_content)
    with open(translated_path, "wb") as f:
        f.write(translated_content)

    # 创建后台任务
    async def run_task_in_background():
        try:
            # 重新读取文件
            from fastapi import UploadFile
            with open(original_path, "rb") as f:
                original_bytes = f.read()
            with open(translated_path, "rb") as f:
                translated_bytes = f.read()

            # 创建临时的UploadFile对象
            import io
            original_upload = UploadFile(
                filename=original_file.filename,
                file=io.BytesIO(original_bytes)
            )
            translated_upload = UploadFile(
                filename=translated_file.filename,
                file=io.BytesIO(translated_bytes)
            )

            await run_number_check_task(original_upload, translated_upload, task_id=task_id)
        except Exception as e:
            # 标记任务失败
            from app.service.number_check_service import _complete_task
            tb = traceback.format_exc()
            print("=" * 60)
            print("数字专检后台任务失败:")
            print(tb)
            print("=" * 60)
            _complete_task(task_id, error=str(e))

    background_tasks.add_task(run_task_in_background)

    # 立即返回task_id，让前端开始轮询
    return {
        "status": "ACCEPTED",
        "task_id": task_id,
        "message": "任务已提交，正在后台处理"
    }


@router.get("/number-check/status/{task_id}")
async def get_number_check_status(task_id: str):
    """
    查询数值专检任务状态和进度
    """
    progress = _get_task_progress(task_id)
    if not progress:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")

    return progress


# ── 多语对照记忆 ──────────────────────────────────────────

@router.get("/alignment/config")
async def get_alignment_config():
    """返回可用模型、语言列表、阈值默认值"""
    return {
        "models": {
            name: {
                "description": info["description"],
                "id": info["id"],
                "max_output": info["max_output"],
            }
            for name, info in ALIGNMENT_MODELS.items()
        },
        "languages": {k: v["description"] for k, v in SUPPORTED_LANGUAGES.items()},
        "thresholds": THRESHOLD_MAP,
        "buffer_chars": BUFFER_CHARS,
    }


@router.post("/alignment")
async def run_alignment(
    background_tasks: BackgroundTasks,
    original_file: UploadFile = File(..., description="原文文件 (docx/doc/pptx/xlsx/xls)"),
    translated_file: UploadFile = File(..., description="译文文件 (docx/doc/pptx/xlsx/xls)"),
    source_lang: str = Query("中文", description="原文语言"),
    target_lang: str = Query("英语", description="译文语言"),
    model_name: str = Query("Google Gemini 2.5 Flash", description="模型名称"),
    enable_post_split: bool = Query(True, description="启用后处理细粒度分句"),
    threshold_2: int = Query(25000, description="分割阈值 2 份"),
    threshold_3: int = Query(50000, description="分割阈值 3 份"),
    threshold_4: int = Query(75000, description="分割阈值 4 份"),
    threshold_5: int = Query(100000, description="分割阈值 5 份"),
    threshold_6: int = Query(125000, description="分割阈值 6 份"),
    threshold_7: int = Query(150000, description="分割阈值 7 份"),
    threshold_8: int = Query(175000, description="分割阈值 8 份"),
    buffer_chars: int = Query(2000, description="缓冲区字数"),
):
    """
    多语对照记忆：上传原文/译文文档，通过 LLM 进行句级对齐，输出 Excel。
    支持 DOCX / DOC / PPTX / XLSX / XLS 格式。
    """
    import uuid
    from pathlib import Path
    from app.core.config import settings

    allowed_ext = {'.docx', '.doc', '.pptx', '.xlsx', '.xls'}
    orig_ext = os.path.splitext(original_file.filename or "")[1].lower()
    trans_ext = os.path.splitext(translated_file.filename or "")[1].lower()

    if orig_ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的原文文件格式: {orig_ext}")
    if trans_ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"不支持的译文文件格式: {trans_ext}")

    task_id = str(uuid.uuid4())

    upload_dir = Path(settings.UPLOAD_DIR) / "alignment"
    upload_dir.mkdir(parents=True, exist_ok=True)

    original_path = upload_dir / f"{task_id}_original{orig_ext}"
    translated_path = upload_dir / f"{task_id}_translated{trans_ext}"

    original_content = await original_file.read()
    translated_content = await translated_file.read()

    with open(original_path, "wb") as f:
        f.write(original_content)
    with open(translated_path, "wb") as f:
        f.write(translated_content)

    async def run_task_in_background():
        try:
            await run_alignment_task(
                original_path=str(original_path),
                translated_path=str(translated_path),
                task_id=task_id,
                source_lang=source_lang,
                target_lang=target_lang,
                model_name=model_name,
                enable_post_split=enable_post_split,
                threshold_2=threshold_2,
                threshold_3=threshold_3,
                threshold_4=threshold_4,
                threshold_5=threshold_5,
                threshold_6=threshold_6,
                threshold_7=threshold_7,
                threshold_8=threshold_8,
                buffer_chars=buffer_chars,
            )
        except Exception as e:
            tb = traceback.format_exc()
            print(f"对齐后台任务失败:\n{tb}")
            _alignment_complete_task(task_id, error=str(e))

    background_tasks.add_task(run_task_in_background)

    return {
        "status": "ACCEPTED",
        "task_id": task_id,
        "message": "对齐任务已提交，正在后台处理",
    }


@router.get("/alignment/status/{task_id}")
async def get_alignment_status(task_id: str):
    """查询对齐任务状态和进度"""
    progress = get_alignment_progress(task_id)
    if not progress:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    return progress
