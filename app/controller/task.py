import traceback
from fastapi import APIRouter, UploadFile, File, Query, HTTPException
from app.service.llm_service import run_llm_task
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
