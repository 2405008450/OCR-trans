import os
import shutil
from datetime import datetime



# =========================
# 5) 文件备份功能
# =========================
# =========================
# 0) 基础配置
# =========================

BACKUP_DIR_NAME = "backup"

def ensure_backup_copy(src_docx_path: str) -> str:
    """
    把译文文件复制到 backup/ 下，生成不重复的新副本文件名
    """
    src_docx_path = os.path.abspath(src_docx_path)
    if not os.path.exists(src_docx_path):
        raise FileNotFoundError(f"译文文件不存在: {src_docx_path}")

    base_dir = os.path.dirname(src_docx_path)
    backup_dir = os.path.join(base_dir, BACKUP_DIR_NAME)
    os.makedirs(backup_dir, exist_ok=True)

    src_name = os.path.basename(src_docx_path)
    stem, ext = os.path.splitext(src_name)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst_name = f"{stem}_corrected_{timestamp}{ext}"
    dst_path = os.path.join(backup_dir, dst_name)

    shutil.copy2(src_docx_path, dst_path)
    return dst_path