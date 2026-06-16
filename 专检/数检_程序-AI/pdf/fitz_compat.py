"""PyMuPDF 导入兼容层，避免误导入第三方旧 fitz 包。"""

try:
    import pymupdf as fitz
except ModuleNotFoundError as exc:
    if exc.name == "pymupdf":
        raise ModuleNotFoundError(
            "PDF 处理需要 PyMuPDF。请使用项目 .venv 运行服务，或在当前 Python 环境安装 PyMuPDF。"
        ) from exc
    raise

