import os
import uuid
from fastapi import UploadFile
from typing import Dict, Any
from app.core.config import settings
from app.service.image_processor import process_image, convert_input_to_images

os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)
os.makedirs(settings.TEMP_IMAGES_DIR, exist_ok=True)

async def run_llm_task(
    file: UploadFile,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = True,
    enable_visualization: bool = True
) -> Dict[str, Any]:
    """
    处理上传的文件（图片或PDF）
    
    Args:
        file: 上传的文件
        from_lang: 源语言，默认'zh'
        to_lang: 目标语言，默认'en'
        enable_correction: 是否启用透视矫正
        enable_visualization: 是否生成可视化图片
    
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
        result = process_image(
            input_path=img_path,
            output_dir=settings.OUTPUT_DIR,
            from_lang=from_lang,
            to_lang=to_lang,
            enable_correction=enable_correction,
            enable_visualization=enable_visualization
        )
        
        # 转换字段名称以匹配前端期望，并标准化路径格式
        def normalize_path(path):
            """将路径转换为URL格式（使用正斜杠）"""
            if path:
                # 将反斜杠转换为正斜杠
                return path.replace("\\", "/")
            return None
        
        formatted_result = {
            "corrected_image": normalize_path(result.get("processed_image")),
            "visualization_image": normalize_path(result.get("visualization")),
            "translated_image": normalize_path(result.get("final_output")),
            "ocr_json": normalize_path(result.get("raw_ocr_json")),
            "translated_json": normalize_path(result.get("translated_json"))
        }
        
        # 调试日志
        print(f"\n📦 格式化后的结果:")
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
