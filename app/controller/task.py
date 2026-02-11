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
    enable_correction: bool = Query(True, description="是否启用透视矫正"),
    enable_visualization: bool = Query(True, description="是否生成可视化图片"),
    card_side: str = Query("front", description="证件面: 'front'=正面, 'back'=背面")
):
    """
    处理图片文件，进行OCR识别和翻译
    
    支持的功能：
    - OCR文字识别
    - 智能分割和合并文本块
    - 多语言翻译（通过DeepSeek API）
    - 图像修复和文字填充
    - 透视矫正（可选）
    - 可视化结果（可选）
    - 正面/背面分别处理（背面取消自动换行）
    """
    try:
        result = await run_llm_task(
            file=file,
            from_lang=from_lang,
            to_lang=to_lang,
            enable_correction=enable_correction,
            enable_visualization=enable_visualization,
            card_side=card_side
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
        print("❌ 任务执行失败:")
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
