#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驾驶证翻译文档生成系统 - 简易启动文件

使用方法：
1. 在下方配置区域填写 API 密钥和文件路径
2. 双击运行此文件，或右键选择"用 Python 打开"
3. 等待处理完成，输出文件会自动生成
"""

import os
import sys
import logging
from pathlib import Path

# ==================== 配置区域 ====================
# 请在这里填写您的配置信息

# API 密钥配置（必填）
GLM_API_KEY = "c7970451446347e29f3f586514ef6894.8O7WrHEoZpZ7Y4HM"          # 智谱 GLM API 密钥
DEEPSEEK_API_KEY = "sk-a681d5bc9d3f4278b084ad4606251298"  # DeepSeek API 密钥

# 处理模式选择：
# - "single": 单文件处理（一张图片 → 一个文档）
# - "batch": 批量处理整个目录（每张图片 → 一个文档）
# - "merge": 合并处理（多张图片 → 一个文档，属于同一个驾驶证）
PROCESSING_MODE = "merge"

# 单文件处理配置（PROCESSING_MODE = "single"）
INPUT_IMAGE_PATH = "jsz/1.png"    # 输入图片路径

# 批量处理配置（PROCESSING_MODE = "batch"）
INPUT_DIR = "jsz"                    # 输入目录路径

# 合并处理配置（PROCESSING_MODE = "merge"）
# 多张图片合并为一个驾驶证文档
INPUT_IMAGE_PATHS = [
    "jsz/1.png",


  
   
    
]

OUTPUT_DIR = "jsz_translated"      # 输出目录

# ==================== 配置区域结束 ====================

#控制台打印的日志级别默认INFO,现在这样（WARNING）只会显示警告和错误信息。
def setup_logging():
    """配置日志系统"""
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def validate_config():
    """
    验证配置
    
    Returns:
        tuple: (is_valid, error_message)
    """
    # 验证 API 密钥
    if GLM_API_KEY == "your-glm-api-key-here" or not GLM_API_KEY:
        return False, "请先配置 GLM_API_KEY"
    
    if DEEPSEEK_API_KEY == "your-deepseek-api-key-here" or not DEEPSEEK_API_KEY:
        return False, "请先配置 DEEPSEEK_API_KEY"
    
    # 验证处理模式
    if PROCESSING_MODE not in ["single", "batch", "merge"]:
        return False, f"无效的处理模式: {PROCESSING_MODE}，必须是 'single', 'batch' 或 'merge'"
    
    # 根据模式验证输入
    if PROCESSING_MODE == "single":
        if not INPUT_IMAGE_PATH:
            return False, "单文件模式需要配置 INPUT_IMAGE_PATH"
        if not os.path.exists(INPUT_IMAGE_PATH):
            return False, f"输入文件不存在: {INPUT_IMAGE_PATH}"
    
    elif PROCESSING_MODE == "batch":
        if not INPUT_DIR:
            return False, "批量处理模式需要配置 INPUT_DIR"
        if not os.path.exists(INPUT_DIR):
            return False, f"输入目录不存在: {INPUT_DIR}"
    
    elif PROCESSING_MODE == "merge":
        if not INPUT_IMAGE_PATHS or len(INPUT_IMAGE_PATHS) == 0:
            return False, "合并处理模式需要配置 INPUT_IMAGE_PATHS（至少一张图片）"
        for path in INPUT_IMAGE_PATHS:
            if not os.path.exists(path):
                return False, f"输入文件不存在: {path}"
    
    return True, None


def get_processing_mode():
    """
    确定处理模式和输入源
    
    Returns:
        tuple: (mode, input_source)
        mode: "single", "batch" 或 "merge"
        input_source: 输入文件路径、目录路径或文件列表
    """
    if PROCESSING_MODE == "single":
        return "single", INPUT_IMAGE_PATH
    elif PROCESSING_MODE == "batch":
        return "batch", INPUT_DIR
    elif PROCESSING_MODE == "merge":
        return "merge", INPUT_IMAGE_PATHS
    else:
        raise ValueError(f"未知的处理模式: {PROCESSING_MODE}")


def collect_image_files(directory):
    """
    收集目录中的所有图片文件
    
    Args:
        directory: 目录路径
        
    Returns:
        list: 图片文件路径列表
    """
    image_files = []
    for ext in ['*.png', '*.jpg', '*.jpeg', '*.PNG', '*.JPG', '*.JPEG']:
        image_files.extend(Path(directory).glob(ext))
    return [str(f) for f in image_files]


def open_output_folder(folder_path):
    """
    打开输出文件夹
    
    Args:
        folder_path: 文件夹路径
    """
    import subprocess
    try:
        if sys.platform == 'win32':
            os.startfile(folder_path)
        elif sys.platform == 'darwin':
            subprocess.run(['open', folder_path])
        else:
            subprocess.run(['xdg-open', folder_path])
    except Exception as e:
        print(f"无法打开文件夹: {str(e)}")


def main():
    """主函数"""
    print("=" * 50)
    print("驾驶证翻译文档生成系统")
    print("=" * 50)
    print()
    
    # 配置日志
    setup_logging()
    
    # 验证配置
    is_valid, error_message = validate_config()
    if not is_valid:
        print(f"❌ 错误: {error_message}")
        return
    
    # 确定处理模式
    mode, input_source = get_processing_mode()
    
    # 创建输出目录
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    except Exception as e:
        print(f"❌ 错误: 无法创建输出目录: {str(e)}")
        return
    
    # 导入翻译系统
    try:
        from src.translator_pipeline import TranslatorPipeline
    except ImportError as e:
        print(f"❌ 错误: 无法导入翻译系统: {str(e)}")
        print("请确保已安装所有依赖: pip install -r requirements.txt")
        return
    
    # 初始化翻译流程
    print("正在初始化翻译系统...")
    try:
        pipeline = TranslatorPipeline(GLM_API_KEY, DEEPSEEK_API_KEY)
    except Exception as e:
        print(f"❌ 错误: 初始化失败: {str(e)}")
        return
    
    try:
        if mode == "single":
            # 单文件处理
            print(f"\n正在处理: {input_source}")
            print("-" * 50)
            output_path = pipeline.translate_image(input_source, OUTPUT_DIR)
            print("-" * 50)
            print(f"✅ 翻译完成: {output_path}")
        
        elif mode == "merge":
            # 合并处理：多张图片合并为一个文档
            print(f"\n合并处理模式：将 {len(input_source)} 张图片合并为一个驾驶证文档")
            print("-" * 50)
            for i, img_path in enumerate(input_source, 1):
                print(f"  {i}. {img_path}")
            print("-" * 50)
            
            output_path = pipeline.translate_merge(input_source, OUTPUT_DIR)
            print("-" * 50)
            print(f"[OK] 翻译完成: {output_path}")
        
        else:
            # 批量处理
            image_files = collect_image_files(input_source)
            
            if not image_files:
                print(f"[ERROR] 错误: 在目录 {input_source} 中未找到图片文件")
                return
            
            print(f"\n找到 {len(image_files)} 个图片文件")
            print("-" * 50)
            
            results = pipeline.translate_batch(image_files, OUTPUT_DIR)
            
            print("-" * 50)
            # 统计结果
            success_count = sum(1 for v in results.values() if not v.startswith("ERROR"))
            fail_count = len(results) - success_count
            
            print(f"\n处理完成:")
            print(f"[OK] 成功: {success_count}/{len(results)}")
            if fail_count > 0:
                print(f"[ERROR] 失败: {fail_count}/{len(results)}")
                print("\n失败的文件:")
                for path, result in results.items():
                    if result.startswith("ERROR"):
                        print(f"  - {os.path.basename(path)}: {result}")
        
        # 询问是否打开输出文件夹
        print()
        print("[OK] 处理完成！输出文件已保存到:", OUTPUT_DIR)
    
    except KeyboardInterrupt:
        print("\n\n[WARNING] 用户中断操作")
        return
    except Exception as e:
        print(f"\n[ERROR] 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return
    
    print("\n处理完成！")


if __name__ == "__main__":
    main()
