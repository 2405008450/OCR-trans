"""
统一的备份管理模块
支持 Word 和 PDF 文件的备份
"""

import os
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional


BACKUP_DIR_NAME = "backup"

# 固定输出目录；None 时退化为源文件旁的 backup/ 子目录
OUTPUT_DIR: Optional[str] = r"D:\project\数检_程序-AI\output"


def ensure_backup_copy(src_file_path: str, suffix: str = "corrected") -> str:
    """
    把文件复制到输出目录，文件名格式：原文件名_时间戳.扩展名
    支持 Word (.docx) 和 PDF (.pdf) 文件

    Args:
        src_file_path: 源文件路径
        suffix: 已废弃参数，保留兼容性，不再写入文件名

    Returns:
        输出文件的完整路径

    Raises:
        FileNotFoundError: 如果源文件不存在
    """
    src_file_path = os.path.abspath(src_file_path)

    if not os.path.exists(src_file_path):
        raise FileNotFoundError(f"文件不存在: {src_file_path}")

    stem, ext = os.path.splitext(os.path.basename(src_file_path))

    # 确定输出目录：优先使用 OUTPUT_DIR，否则退化到源文件旁的 backup/ 子目录
    if OUTPUT_DIR:
        out_dir = OUTPUT_DIR
    else:
        out_dir = os.path.join(os.path.dirname(src_file_path), BACKUP_DIR_NAME)
    os.makedirs(out_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst_name = f"{stem}_{timestamp}{ext}"
    dst_path = os.path.join(out_dir, dst_name)

    shutil.copy2(src_file_path, dst_path)
    print(f"✓ 已创建输出文件: {dst_path}")

    return dst_path


def create_backup_with_custom_name(
    src_file_path: str, 
    custom_name: Optional[str] = None
) -> str:
    """
    创建备份文件，可以指定自定义文件名
    
    Args:
        src_file_path: 源文件路径
        custom_name: 自定义文件名（不含扩展名），如果为 None 则自动生成
        
    Returns:
        备份文件的完整路径
    """
    src_file_path = os.path.abspath(src_file_path)
    
    if not os.path.exists(src_file_path):
        raise FileNotFoundError(f"文件不存在: {src_file_path}")
    
    base_dir = os.path.dirname(src_file_path)
    ext = os.path.splitext(src_file_path)[1]
    
    # 创建 backup 目录
    backup_dir = os.path.join(base_dir, BACKUP_DIR_NAME)
    os.makedirs(backup_dir, exist_ok=True)
    
    # 确定目标文件名
    if custom_name:
        dst_name = f"{custom_name}{ext}"
    else:
        src_name = os.path.basename(src_file_path)
        stem = os.path.splitext(src_name)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst_name = f"{stem}_backup_{timestamp}{ext}"
    
    dst_path = os.path.join(backup_dir, dst_name)
    
    # 如果文件已存在，添加序号
    if os.path.exists(dst_path):
        base_name = os.path.splitext(dst_name)[0]
        counter = 1
        while os.path.exists(dst_path):
            dst_name = f"{base_name}_{counter}{ext}"
            dst_path = os.path.join(backup_dir, dst_name)
            counter += 1
    
    # 复制文件
    shutil.copy2(src_file_path, dst_path)
    
    print(f"✓ 已创建备份: {dst_name}")
    
    return dst_path


def get_backup_dir(file_path: str) -> str:
    """
    获取文件的备份目录路径
    
    Args:
        file_path: 文件路径
        
    Returns:
        备份目录的完整路径
    """
    base_dir = os.path.dirname(os.path.abspath(file_path))
    return os.path.join(base_dir, BACKUP_DIR_NAME)


def list_backups(file_path: str) -> list:
    """
    列出文件的所有备份
    
    Args:
        file_path: 原始文件路径
        
    Returns:
        备份文件列表（按时间倒序）
    """
    backup_dir = get_backup_dir(file_path)
    
    if not os.path.exists(backup_dir):
        return []
    
    src_name = os.path.basename(file_path)
    stem = os.path.splitext(src_name)[0]
    ext = os.path.splitext(src_name)[1]
    
    # 查找所有相关备份
    backups = []
    for filename in os.listdir(backup_dir):
        if filename.startswith(stem) and filename.endswith(ext):
            full_path = os.path.join(backup_dir, filename)
            backups.append({
                'path': full_path,
                'name': filename,
                'mtime': os.path.getmtime(full_path),
                'size': os.path.getsize(full_path)
            })
    
    # 按修改时间倒序排序
    backups.sort(key=lambda x: x['mtime'], reverse=True)
    
    return backups


# 使用示例
if __name__ == "__main__":
    # 示例1：创建 Word 文档备份
    word_file = "document.docx"
    if os.path.exists(word_file):
        backup_path = ensure_backup_copy(word_file)
        print(f"Word 备份路径: {backup_path}")
    
    # 示例2：创建 PDF 文件备份
    pdf_file = "document.pdf"
    if os.path.exists(pdf_file):
        backup_path = ensure_backup_copy(pdf_file, suffix="annotated")
        print(f"PDF 备份路径: {backup_path}")
    
    # 示例3：列出所有备份
    if os.path.exists(word_file):
        backups = list_backups(word_file)
        print(f"\n找到 {len(backups)} 个备份:")
        for backup in backups:
            print(f"  - {backup['name']} ({backup['size']} bytes)")
