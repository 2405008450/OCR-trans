import os
import uuid
from fastapi import UploadFile
from typing import Dict, Any, Optional
from app.core.config import settings
from app.service.image_processor import process_image, convert_input_to_images
from app.service.marriage_cert_processor import process_marriage_cert_image

os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)
os.makedirs(settings.TEMP_IMAGES_DIR, exist_ok=True)

async def run_llm_task(
    file: UploadFile,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = 'front',
    doc_type: str = 'id_card',
    marriage_page_template: str = 'page2',
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

            print(
                f"结婚证模板: {template} | "
                f"merge={enable_merge}, overlap_fix={enable_overlap_fix}, "
                f"colon_fix={enable_colon_fix}, confidence={confidence_threshold}"
            )

            result = process_marriage_cert_image(
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
            )
        else:
            # 默认：身份证处理
            result = process_image(
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
