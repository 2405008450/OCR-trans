"""
Word 转 PDF + 完整内容提取（模块化版本）
支持外部调用和命令行使用

使用示例：
    # 方式1：作为模块导入
    from word_to_pdf_extractor import WordToPDFConverter

    converter = WordToPDFConverter()
    result = converter.convert_and_extract(
        word_path="input.docx",
        output_dir="./output",
        extract_mode="clean"
    )
    print(result['text_content'])

    # 方式2：命令行运行
    python word_to_pdf_extractor.py input.docx --output ./output --mode clean
"""

import os
import sys
import argparse
from pathlib import Path
import pdfplumber
from typing import List, Dict, Tuple, Optional, Any
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime


# ================= 数据结构 =================

@dataclass
class PageContent:
    """单页内容结构"""
    page_num: int
    header: str = ""
    footer: str = ""
    body_text: str = ""
    tables: List[List[List[str]]] = field(default_factory=list)
    full_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return asdict(self)


@dataclass
class ExtractionResult:
    """提取结果结构"""
    success: bool
    word_path: str
    pdf_path: str
    txt_path: str
    text_content: str
    pages_content: List[PageContent]
    statistics: Dict[str, Any]
    error_message: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典（用于 JSON 序列化）"""
        return {
            "success": self.success,
            "word_path": self.word_path,
            "pdf_path": self.pdf_path,
            "txt_path": self.txt_path,
            "text_content": self.text_content,
            "pages_content": [page.to_dict() for page in self.pages_content],
            "statistics": self.statistics,
            "error_message": self.error_message,
        }


# ================= 配置类 =================

class ExtractorConfig:
    """提取器配置"""

    # 页眉/页脚检测参数
    HEADER_HEIGHT_RATIO: float = 0.05
    FOOTER_HEIGHT_RATIO: float = 0.05

    # 表格识别参数
    TABLE_SETTINGS: Dict[str, Any] = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 5,
        "snap_tolerance": 5,
        "join_tolerance": 5,
    }

    # 默认输出文件名
    DEFAULT_PDF_NAME: str = "output.pdf"
    DEFAULT_TXT_NAME: str = "extracted_content.txt"


# ================= Word 转 PDF =================

class WordToPDFService:
    """Word 转 PDF 服务"""

    @staticmethod
    def convert(word_path: str, pdf_path: str, verbose: bool = True) -> bool:
        """
        使用 Word COM 接口转换 PDF

        Args:
            word_path: Word 文档路径
            pdf_path: 输出 PDF 路径
            verbose: 是否打印详细信息

        Returns:
            转换是否成功
        """
        word_path = str(Path(word_path).resolve())
        pdf_path = str(Path(pdf_path).resolve())

        if sys.platform != "win32":
            if verbose:
                print("[Word->PDF] 当前平台不支持 win32com，Word 转 PDF 仅支持 Windows。")
            return False

        try:
            import win32com.client
        except ImportError:
            if verbose:
                print("[Word->PDF] 未安装 win32com/pywin32，无法执行 Word 转 PDF。")
            return False

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        try:
            if verbose:
                print(f"[Word->PDF] 正在打开文档: {Path(word_path).name}")

            doc = word.Documents.Open(word_path)

            if verbose:
                print(f"[Word->PDF] 正在转换为 PDF...")

            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()

            if verbose:
                print(f"[Word->PDF] ✓ 转换成功\n")

            return True

        except Exception as e:
            if verbose:
                print(f"[Word->PDF] ✗ 转换失败: {e}")
            return False

        finally:
            word.Quit()


# ================= PDF 提取器 =================

class CompletePDFExtractor:
    """完整 PDF 提取器"""

    def __init__(self, pdf_path: str, config: Optional[ExtractorConfig] = None):
        self.pdf_path = pdf_path
        self.config = config or ExtractorConfig()
        self.pages_content: List[PageContent] = []

    def extract_all(self, verbose: bool = True) -> List[PageContent]:
        """提取所有页面的完整内容"""
        if verbose:
            print(f"[PDF提取] 正在提取: {Path(self.pdf_path).name}")

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            if verbose:
                print(f"[PDF提取] 总页数: {total_pages}\n")

            for page_num, page in enumerate(pdf.pages, start=1):
                if verbose:
                    print(f"[PDF提取] 处理第 {page_num}/{total_pages} 页...", end="\r")

                page_content = PageContent(page_num=page_num)

                # 1. 提取完整文本
                page_content.full_text = self._extract_full_text(page)

                # 2. 提取表格
                page_content.tables = self._extract_tables(page)

                # 3. 提取页眉和页脚
                page_content.header, page_content.footer = self._extract_header_footer(page)

                # 4. 提取正文
                page_content.body_text = self._extract_body_from_full_text(
                    page_content.full_text,
                    page_content.header,
                    page_content.footer
                )

                self.pages_content.append(page_content)

        if verbose:
            print(f"\n[PDF提取] ✓ 提取完成，共处理 {len(self.pages_content)} 页\n")

        return self.pages_content

    def _extract_full_text(self, page) -> str:
        """提取完整文本"""
        try:
            text = page.extract_text(layout=True)
            if text:
                return self._clean_text(text)

            text = page.extract_text()
            return self._clean_text(text) if text else ""

        except Exception as e:
            print(f"\n[警告] 页面 {page.page_number} 文本提取失败 - {e}")
            return ""

    def _extract_tables(self, page) -> List[List[List[str]]]:
        """提取表格"""
        tables = []

        try:
            extracted_tables = page.extract_tables(
                table_settings=self.config.TABLE_SETTINGS
            )

            if extracted_tables:
                for table in extracted_tables:
                    cleaned_table = []
                    for row in table:
                        cleaned_row = [
                            self._clean_text(cell) if cell else ""
                            for cell in row
                        ]
                        if any(cell.strip() for cell in cleaned_row):
                            cleaned_table.append(cleaned_row)

                    if cleaned_table:
                        tables.append(cleaned_table)

        except Exception as e:
            print(f"\n[警告] 表格提取失败 - {e}")

        return tables

    def _extract_header_footer(self, page) -> Tuple[str, str]:
        """提取页眉和页脚"""
        height = page.height

        # 页眉区域
        header_region = (
            0, 0,
            page.width,
            height * self.config.HEADER_HEIGHT_RATIO
        )
        header_crop = page.crop(header_region)
        header_text = self._clean_text(header_crop.extract_text() or "")

        # 页脚区域
        footer_region = (
            0,
            height * (1 - self.config.FOOTER_HEIGHT_RATIO),
            page.width,
            height
        )
        footer_crop = page.crop(footer_region)
        footer_text = self._clean_text(footer_crop.extract_text() or "")

        return header_text, footer_text

    def _extract_body_from_full_text(
        self,
        full_text: str,
        header: str,
        footer: str
    ) -> str:
        """从完整文本中移除页眉页脚"""
        body = full_text

        if header and header in body:
            body = body.replace(header, "", 1)

        if footer and footer in body:
            idx = body.rfind(footer)
            if idx != -1:
                body = body[:idx] + body[idx + len(footer):]

        return self._clean_text(body)

    @staticmethod
    def _clean_text(text: str) -> str:
        """清理文本"""
        if not text:
            return ""

        lines = text.split('\n')
        lines = [line.rstrip() for line in lines]
        text = '\n'.join(lines)
        text = re.sub(r'\n{4,}', '\n\n', text)

        return text.strip()


# ================= 格式化器 =================

class ContentFormatter:
    """内容格式化输出类"""

    @staticmethod
    def format_to_text(
        pages_content: List[PageContent],
        mode: str = "clean"
    ) -> str:
        """
        格式化输出

        Args:
            pages_content: 页面内容列表
            mode: "structured" 或 "clean"
        """
        if mode == "clean":
            return ContentFormatter._format_clean(pages_content)
        else:
            return ContentFormatter._format_structured(pages_content)

    @staticmethod
    def _format_clean(pages_content: List[PageContent]) -> str:
        """纯净输出"""
        output_parts = []

        unique_headers = set()
        unique_footers = set()

        for page in pages_content:
            if page.header and len(page.header) < 200:
                unique_headers.add(page.header)
            if page.footer and len(page.footer) < 200:
                unique_footers.add(page.footer)

        # 输出页眉
        if unique_headers:
            for header in sorted(unique_headers):
                output_parts.append(header)
            output_parts.append("")

        # 输出正文
        for page in pages_content:
            if page.full_text:
                output_parts.append(page.full_text)
                output_parts.append("")

        # 输出页脚
        if unique_footers:
            for footer in sorted(unique_footers):
                output_parts.append(footer)

        return "\n".join(output_parts)

    @staticmethod
    def _format_structured(pages_content: List[PageContent]) -> str:
        """结构化输出"""
        output_lines = []

        unique_headers = set()
        unique_footers = set()

        for page in pages_content:
            if page.header and len(page.header) < 200:
                unique_headers.add(page.header)
            if page.footer and len(page.footer) < 200:
                unique_footers.add(page.footer)

        # 输出页眉
        if unique_headers:
            output_lines.append("=" * 80)
            output_lines.append("页眉内容")
            output_lines.append("=" * 80)
            for header in sorted(unique_headers):
                output_lines.append(header)
            output_lines.append("")

        # 输出每页内容
        for page in pages_content:
            output_lines.append("=" * 80)
            output_lines.append(f"第 {page.page_num} 页")
            output_lines.append("=" * 80)
            output_lines.append("")

            if page.full_text:
                output_lines.append("【完整内容】")
                output_lines.append(page.full_text)
                output_lines.append("")

            if page.tables:
                output_lines.append("【表格内容】")
                for idx, table in enumerate(page.tables, start=1):
                    output_lines.append(f"\n表格 {idx}:")
                    output_lines.append(ContentFormatter._format_table(table))
                output_lines.append("")

        # 输出页脚
        if unique_footers:
            output_lines.append("=" * 80)
            output_lines.append("页脚内容")
            output_lines.append("=" * 80)
            for footer in sorted(unique_footers):
                output_lines.append(footer)

        return "\n".join(output_lines)

    @staticmethod
    def _format_table(table: List[List[str]]) -> str:
        """格式化表格"""
        if not table:
            return ""

        lines = []
        for row in table:
            line = "\t".join(cell.strip() for cell in row)
            lines.append(line)

        return "\n".join(lines)


# ================= 主转换器（外部调用接口） =================

class WordToPDFConverter:
    """
    Word 转 PDF + 内容提取的统一接口

    使用示例：
        converter = WordToPDFConverter()
        result = converter.convert_and_extract(
            word_path="input.docx",
            output_dir="./output"
        )

        if result.success:
            print(result.text_content)
            print(result.statistics)
    """

    def __init__(self, config: Optional[ExtractorConfig] = None):
        self.config = config or ExtractorConfig()

    def convert_and_extract(
        self,
        word_path: str,
        output_dir: Optional[str] = None,
        pdf_name: Optional[str] = None,
        txt_name: Optional[str] = None,
        extract_mode: str = "clean",
        keep_pdf: bool = True,
        verbose: bool = True
    ) -> ExtractionResult:
        """
        转换 Word 为 PDF 并提取内容

        Args:
            word_path: Word 文档路径
            output_dir: 输出目录（默认为 Word 文档所在目录）
            pdf_name: PDF 文件名（默认自动生成）
            txt_name: TXT 文件名（默认自动生成）
            extract_mode: 提取模式 "clean" 或 "structured"
            keep_pdf: 是否保留 PDF 文件
            verbose: 是否打印详细信息

        Returns:
            ExtractionResult 对象
        """

        # 1. 路径处理
        word_path = Path(word_path).resolve()

        if not word_path.exists():
            return ExtractionResult(
                success=False,
                word_path=str(word_path),
                pdf_path="",
                txt_path="",
                text_content="",
                pages_content=[],
                statistics={},
                error_message=f"文件不存在: {word_path}"
            )

        # 确定输出目录
        if output_dir is None:
            output_dir = word_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)

        # 确定输出文件名
        if pdf_name is None:
            pdf_name = f"{word_path.stem}.pdf"

        if txt_name is None:
            txt_name = f"{word_path.stem}_extracted.txt"

        pdf_path = output_dir / pdf_name
        txt_path = output_dir / txt_name

        if verbose:
            print("=" * 80)
            print("Word 转 PDF + 完整内容提取")
            print("=" * 80)
            print(f"源文件: {word_path.name}")
            print(f"输出目录: {output_dir}")
            print(f"提取模式: {extract_mode}")
            print("")

        try:
            # 2. Word 转 PDF
            word_service = WordToPDFService()
            if not word_service.convert(str(word_path), str(pdf_path), verbose):
                return ExtractionResult(
                    success=False,
                    word_path=str(word_path),
                    pdf_path=str(pdf_path),
                    txt_path=str(txt_path),
                    text_content="",
                    pages_content=[],
                    statistics={},
                    error_message="PDF 转换失败"
                )

            # 3. 提取内容
            extractor = CompletePDFExtractor(str(pdf_path), self.config)
            pages_content = extractor.extract_all(verbose)

            # 4. 格式化输出
            formatter = ContentFormatter()
            text_output = formatter.format_to_text(pages_content, mode=extract_mode)

            # 5. 保存文本文件
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(text_output)

            # 6. 统计信息
            statistics = {
                "total_pages": len(pages_content),
                "total_characters": len(text_output),
                "total_lines": len(text_output.splitlines()),
                "total_tables": sum(len(page.tables) for page in pages_content),
                "extraction_time": datetime.now().isoformat(),
            }

            if verbose:
                print("=" * 80)
                print("提取完成")
                print("=" * 80)
                print(f"✓ 输出文件: {txt_path.name}")
                print(f"✓ 总字符数: {statistics['total_characters']:,}")
                print(f"✓ 总行数: {statistics['total_lines']:,}")
                print(f"✓ 总页数: {statistics['total_pages']}")
                print(f"✓ 总表格数: {statistics['total_tables']}")
                print("=" * 80)

            # 7. 删除 PDF（如果不保留）
            if not keep_pdf:
                pdf_path.unlink()
                if verbose:
                    print(f"✓ 已删除临时 PDF: {pdf_path.name}")

            return ExtractionResult(
                success=True,
                word_path=str(word_path),
                pdf_path=str(pdf_path) if keep_pdf else "",
                txt_path=str(txt_path),
                text_content=text_output,
                pages_content=pages_content,
                statistics=statistics
            )

        except Exception as e:
            return ExtractionResult(
                success=False,
                word_path=str(word_path),
                pdf_path=str(pdf_path),
                txt_path=str(txt_path),
                text_content="",
                pages_content=[],
                statistics={},
                error_message=str(e)
            )


# ================= 命令行接口 =================

def main():
    """命令行主函数"""
    parser = argparse.ArgumentParser(
        description="Word 转 PDF + 完整内容提取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python word_to_pdf_extractor.py input.docx
  python word_to_pdf_extractor.py input.docx --output ./output
  python word_to_pdf_extractor.py input.docx --mode structured --no-keep-pdf
        """
    )

    parser.add_argument(
        "word_file",
        help="Word 文档路径"
    )

    parser.add_argument(
        "--output", "-o",
        help="输出目录（默认为 Word 文档所在目录）"
    )

    parser.add_argument(
        "--pdf-name",
        help="PDF 文件名（默认自动生成）"
    )

    parser.add_argument(
        "--txt-name",
        help="TXT 文件名（默认自动生成）"
    )

    parser.add_argument(
        "--mode", "-m",
        choices=["clean", "structured"],
        default="clean",
        help="提取模式：clean（纯净）或 structured（结构化）"
    )

    parser.add_argument(
        "--no-keep-pdf",
        action="store_true",
        help="不保留 PDF 文件"
    )

    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="静默模式（不打印详细信息）"
    )

    args = parser.parse_args()

    # 执行转换
    converter = WordToPDFConverter()
    result = converter.convert_and_extract(
        word_path=args.word_file,
        output_dir=args.output,
        pdf_name=args.pdf_name,
        txt_name=args.txt_name,
        extract_mode=args.mode,
        keep_pdf=not args.no_keep_pdf,
        verbose=not args.quiet
    )

    # 返回状态码
    if result.success:
        if not args.quiet:
            print("\n【内容预览】")
            print("-" * 80)
            preview_length = min(1000, len(result.text_content))
            print(result.text_content[:preview_length])
            if len(result.text_content) > preview_length:
                print("\n... (后续内容已省略)")
            print("-" * 80)
        sys.exit(0)
    else:
        print(f"\n✗ 错误: {result.error_message}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

