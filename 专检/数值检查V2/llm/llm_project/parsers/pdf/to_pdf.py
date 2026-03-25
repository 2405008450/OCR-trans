from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List


@dataclass
class CompletePDFExtractor:
    pdf_path: str

    def extract_all(self, verbose: bool = False) -> List[Dict[str, Any]]:
        pdf_file = Path(self.pdf_path)
        if not pdf_file.exists():
            raise FileNotFoundError(f"PDF 文件不存在: {pdf_file}")

        raise RuntimeError(
            "当前环境缺少可用的 PDF 解析依赖。"
            "程序启动已恢复，但若要处理 PDF，需要安装可用的 PyMuPDF/pypdf/pdfplumber 之一。"
        )


class ContentFormatter:
    @staticmethod
    def format_to_text(pages: List[Dict[str, Any]], mode: str = "clean") -> str:
        if not pages:
            return ""

        if mode == "structured":
            blocks = []
            for page in pages:
                blocks.append(f"=== 第 {page.get('page_number', '?')} 页 ===")
                blocks.append(page.get("text", ""))
            return "\n\n".join(blocks).strip()

        return "\n".join(page.get("text", "") for page in pages if page.get("text")).strip()
