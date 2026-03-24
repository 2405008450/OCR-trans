"""
OCR 图片 → 混合 HTML/Markdown → 格式化 Word 文档 (一步到位)
"""

import base64
import re
import sys
import os
import time
import subprocess
import shutil
from pathlib import Path

from app.service.gemini_service import generate_vision_html

# ============================================================
# 依赖检查与导入
# ============================================================
try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor, Cm, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
except ImportError:
    print("请安装 python-docx: pip install python-docx")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError:
    print("请安装 BeautifulSoup: pip install beautifulsoup4")
    sys.exit(1)


# ============================================================
# 1. LLM OCR 调用
# ============================================================
LIBREOFFICE_PATH = os.getenv("LIBREOFFICE_PATH", "").strip()

SYS_PROMPT = """Convert the document into simple, Microsoft Word-safe HTML.

Return only a complete HTML document. Do not include explanations. Do not use markdown. Do not wrap the output in code fences.

The output must be a full HTML document with:
- <!DOCTYPE html>
- <html>
- <head>
- <meta charset="utf-8">
- <title>Document</title>
- <body>

Important:
- The HTML will be opened by Microsoft Word, not a browser.
- Prioritize Word compatibility over visual fidelity.
- Use simple HTML only.

Rules:
- Include all visible text content.
- Preserve reading order.
- Do not summarize or rewrite.
- Do not invent missing text.
- Replace any sequence of 3 or more repeated characters used as visual separators or fill lines (e.g. -----, _____, ....., =====) with a single short placeholder: ___
- For blank form fields shown as long lines (e.g. "Name: _____________"), keep only one short underscore placeholder: "Name: ___"
- Do NOT reproduce decorative divider lines or page-wide rules; omit them entirely.

Allowed tags:
- <html>, <head>, <meta>, <title>, <body>
- <h1> to <h6>
- <p>
- <br>
- <strong>, <em>, <u>, <sub>, <sup>
- <ul>, <ol>, <li>
- <table>, <thead>, <tbody>, <tr>, <th>, <td>
- <hr>

Formatting rules:
- Prefer <p> over nested <div>.
- Use tables for structured alignment when needed.
- Keep styles minimal and inline only when necessary.
- Only use simple inline styles such as:
  text-align, font-size, font-family, font-weight, margin, text-indent
- Do not use:
  flex, grid, float, position, transform, rgba(), opacity, negative margins, external CSS, JavaScript, SVG, canvas
- Avoid unnecessary nesting.

Images:
- If an image is essential, represent it as a short text note like:
  <p>[Logo]</p>
- Do not rely on external image paths unless explicitly provided and required.

Output must be valid HTML that Microsoft Word can open directly.
Visual color preservation:
- Preserve important visual colors from the document.
- If text is clearly colored (e.g., red titles or red stamps), reflect that using simple HTML inline styles.

Seals / stamps:
- If a stamp or official seal is present, represent it as a paragraph.
- Use red color to reflect the stamp.
- Prefer simple styles compatible with Microsoft Word.
- Do not use rgba() or opacity.
- Example:
  <p style="color:#C00000; font-weight:bold;">[Official Seal]</p>"""


def _ocr_single_image(
    img_b64: str,
    mime_type: str,
    model: str,
    gemini_route: str = "google",
    retries: int = 3,
) -> str:
    """对单张图片（base64）调用 LLM OCR，失败自动重试"""
    for attempt in range(retries):
        try:
            response_text = generate_vision_html(
                system_prompt=SYS_PROMPT,
                image_bytes=base64.standard_b64decode(img_b64),
                mime_type=mime_type,
                model=model,
                route=gemini_route,
                temperature=0,
            )
            print(response_text, end="", flush=True)
            return response_text
        except Exception as e:
            if attempt < retries - 1:
                wait = 3 * (attempt + 1)
                print(f"\n⚠️ 请求失败({e.__class__.__name__})，{wait} 秒后重试 [{attempt + 1}/{retries}]...")
                time.sleep(wait)
            else:
                raise


def ocr_file(
    file_path: str,
    api_key: str = "",
    model: str = "google/gemini-3.1-pro-preview",
    gemini_route: str = "google",
    page_progress_callback=None,
) -> str:
    """调用 LLM 对图片或 PDF 进行 OCR，返回混合 HTML/Markdown 文本。
    PDF 会逐页渲染为图片后分页发送，避免大文件直传超限。
    page_progress_callback(current_page, total_pages) 可选，用于报告逐页进度。
    """
    ext = Path(file_path).suffix.lower()

    if ext == ".pdf":
        try:
            import fitz  # PyMuPDF
        except ImportError:
            print("请先安装 PyMuPDF: pip install pymupdf")
            sys.exit(1)

        doc = fitz.open(file_path)
        total = len(doc)
        all_results = []

        for i, page in enumerate(doc):
            print(f"\n🔄 正在处理第 {i + 1}/{total} 页...")
            if page_progress_callback:
                try:
                    page_progress_callback(i + 1, total)
                except Exception:
                    pass
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img_b64 = base64.standard_b64encode(pix.tobytes("jpeg", jpg_quality=85)).decode("utf-8")
            text = _ocr_single_image(img_b64, "image/jpeg", model, gemini_route=gemini_route)
            all_results.append(text)
            if i < total - 1:
                time.sleep(1)

        print("\n✅ PDF OCR 完成")
        return "\n\n<page_break/>\n\n".join(all_results)

    else:
        mime_map = {
            ".png":  "image/png",
            ".jpg":  "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif":  "image/gif",
            ".webp": "image/webp",
            ".bmp":  "image/bmp",
        }
        mime_type = mime_map.get(ext, "image/png")

        with open(file_path, "rb") as f:
            img_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

        print("🔄 正在调用 LLM 进行 OCR...")
        result = _ocr_single_image(img_b64, mime_type, model, gemini_route=gemini_route)
        print("\n✅ OCR 完成")
        return result


# ============================================================
# 2. 混合格式解析与 DOCX 渲染引擎
# ============================================================
def normalize_to_word_html(raw_text: str, title: str = "Document") -> str:
    """将 OCR/翻译结果整理为可交给 LibreOffice 转 Word 的完整 HTML。"""
    text = raw_text.strip()
    text = re.sub(r"^```(?:html|markdown)?\s*\n?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\n?```\s*$", "", text)

    page_break_pattern = re.compile(r"\s*<page_break\s*/>\s*", flags=re.IGNORECASE)
    parts = [part.strip() for part in page_break_pattern.split(text) if part.strip()]
    if not parts:
        parts = [""]

    body_segments = []
    for part in parts:
        soup = BeautifulSoup(part, "html.parser")
        body = soup.find("body")
        if body:
            segment = "".join(str(child) for child in body.contents).strip()
        else:
            segment = part
        body_segments.append(segment)

    page_break_html = '<p style="page-break-before: always;"></p>'
    body_html = f"\n{page_break_html}\n".join(body_segments)

    return (
        "<!DOCTYPE html>\n"
        "<html>\n"
        "<head>\n"
        '  <meta charset="utf-8">\n'
        f"  <title>{title}</title>\n"
        "</head>\n"
        "<body>\n"
        f"{body_html}\n"
        "</body>\n"
        "</html>\n"
    )


def convert_html_to_docx_via_libreoffice(
    html_text: str,
    output_path: str,
    html_output_path: str | None = None,
    libreoffice_path: str = LIBREOFFICE_PATH,
) -> str:
    """将 HTML 写盘后通过 LibreOffice 转为 DOCX。"""
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    html_file = Path(html_output_path) if html_output_path else output_file.with_suffix(".html")
    html_file.parent.mkdir(parents=True, exist_ok=True)
    html_file.write_text(html_text, encoding="utf-8")
    user_profile_dir = html_file.parent / ".libreoffice-profile"
    user_profile_dir.mkdir(parents=True, exist_ok=True)

    libreoffice = _resolve_libreoffice_path(libreoffice_path)

    expected_docx = html_file.with_suffix(".docx")
    if expected_docx.exists():
        expected_docx.unlink()
    if output_file.exists() and output_file != expected_docx:
        output_file.unlink()

    result = subprocess.run(
        [
            str(libreoffice),
            f"-env:UserInstallation={user_profile_dir.resolve().as_uri()}",
            "--headless",
            "--convert-to",
            "docx:Office Open XML Text",
            "--outdir",
            str(html_file.parent),
            str(html_file),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice 转换失败: "
            f"returncode={result.returncode}, stdout={result.stdout}, stderr={result.stderr}"
        )

    if not expected_docx.exists():
        raise RuntimeError(
            "LibreOffice 未生成 DOCX 文件: "
            f"stdout={result.stdout}, stderr={result.stderr}"
        )

    if expected_docx != output_file:
        expected_docx.replace(output_file)

    return str(output_file)


def _resolve_libreoffice_path(configured_path: str | None = None) -> str:
    """优先读取显式配置，其次从环境变量和 PATH 中查找 LibreOffice 可执行文件。"""
    candidates: list[str] = []

    for candidate in [
        configured_path,
        os.getenv("LIBREOFFICE_PATH", "").strip(),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "soffice",
        "libreoffice",
    ]:
        if candidate and candidate not in candidates:
            candidates.append(candidate)

    for candidate in candidates:
        if any(sep in candidate for sep in ("/", "\\")) or Path(candidate).is_absolute():
            if Path(candidate).exists():
                return str(Path(candidate))
            continue

        resolved = shutil.which(candidate)
        if resolved:
            return resolved

    raise FileNotFoundError(
        "未找到 LibreOffice 可执行文件。请安装 LibreOffice，并在环境变量 LIBREOFFICE_PATH "
        "中指定 soffice 路径，或确保 `soffice` 已加入 PATH。"
    )


def convert_text_to_word_via_libreoffice(
    raw_text: str,
    output_path: str,
    html_output_path: str | None = None,
    title: str = "Document",
) -> tuple[str, str]:
    """将文本规范化为 HTML，并用 LibreOffice 转 DOCX。"""
    html_text = normalize_to_word_html(raw_text, title=title)
    html_file = html_output_path or str(Path(output_path).with_suffix(".html"))
    docx_file = convert_html_to_docx_via_libreoffice(
        html_text=html_text,
        output_path=output_path,
        html_output_path=html_file,
    )
    return html_file, docx_file


class HybridToDocxConverter:
    """
    将混合 HTML + Markdown 的 OCR 输出转换为格式化的 Word 文档。
    自己实现解析和渲染，不依赖 pandoc 或 htmldocx。
    """

    def __init__(self):
        self.doc = Document()
        self._setup_default_styles()

    def _setup_default_styles(self):
        """设置文档默认样式"""
        style = self.doc.styles["Normal"]
        font = style.font
        font.name = "Arial"
        font.size = Pt(10.5)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.line_spacing = 1.15

        # 设置中文字体回退
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="微软雅黑"/>')
            rPr.append(rFonts)
        else:
            rFonts.set(qn("w:eastAsia"), "微软雅黑")

        # 设置页面边距
        for section in self.doc.sections:
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

    # ----------------------------------------------------------
    # 主转换入口
    # ----------------------------------------------------------
    def convert(self, raw_text: str, output_path: str):
        """主转换方法"""
        # Step 1: 预处理 — 统一格式
        processed = self._preprocess(raw_text)

        # Step 2: 按块解析并渲染
        self._parse_and_render(processed)

        # Step 3: 保存
        self.doc.save(output_path)
        print(f"✅ Word 文档已保存: {output_path}")

    # ----------------------------------------------------------
    # 预处理
    # ----------------------------------------------------------
    def _preprocess(self, text: str) -> str:
        """将 Markdown 语法统一转为 HTML 标签，方便后续用 BS4 解析"""
        # 去除可能的 code fence 包裹
        text = re.sub(r"^```(?:markdown|html)?\s*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)

        # Markdown bold **text** → <strong>text</strong>
        # 注意不要破坏已经在 HTML 标签属性里的 ** (极少见)
        text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)

        # Markdown italic *text* → <em>text</em>  (单个 * 且不在 ** 内)
        text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<em>\1</em>", text)

        # Markdown headers: # Header → <h1>Header</h1>
        text = re.sub(r"^######\s+(.+)$", r"<h6>\1</h6>", text, flags=re.MULTILINE)
        text = re.sub(r"^#####\s+(.+)$", r"<h5>\1</h5>", text, flags=re.MULTILINE)
        text = re.sub(r"^####\s+(.+)$", r"<h4>\1</h4>", text, flags=re.MULTILINE)
        text = re.sub(r"^###\s+(.+)$", r"<h3>\1</h3>", text, flags=re.MULTILINE)
        text = re.sub(r"^##\s+(.+)$", r"<h2>\1</h2>", text, flags=re.MULTILINE)
        text = re.sub(r"^#\s+(.+)$", r"<h1>\1</h1>", text, flags=re.MULTILINE)

        # Markdown ordered list: 1. item → <ol_item>item</ol_item> (自定义标记)
        text = re.sub(
            r"^(\d+)\.\s+(.+)$",
            r'<p class="ol_item" data-num="\1">\2</p>',
            text,
            flags=re.MULTILINE,
        )

        # Markdown unordered list: - item → <p class="ul_item">item</p>
        text = re.sub(
            r"^[-*+]\s+(.+)$",
            r'<p class="ul_item">\1</p>',
            text,
            flags=re.MULTILINE,
        )

        # Markdown horizontal rule --- → <hr>
        text = re.sub(r"^---+\s*$", "<hr/>", text, flags=re.MULTILINE)

        # 包装为 HTML body 以便 BS4 解析
        html = f"<body>{text}</body>"
        return html

    # ----------------------------------------------------------
    # 解析与渲染
    # ----------------------------------------------------------
    def _parse_and_render(self, html: str):
        """用 BeautifulSoup 解析 HTML 并逐块渲染到 docx"""
        soup = BeautifulSoup(html, "html.parser")
        body = soup.find("body") or soup

        for element in body.children:
            self._render_element(element)

    def _render_element(self, element):
        """递归渲染单个元素"""
        if isinstance(element, NavigableString):
            text = str(element).strip()
            if text and text != "\n":
                # 纯文本行 → 段落
                lines = text.split("\n")
                for line in lines:
                    line = line.strip()
                    if line:
                        para = self.doc.add_paragraph()
                        self._add_inline_text(para, line)
            return

        if not isinstance(element, Tag):
            return

        tag_name = element.name.lower()

        if tag_name in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._render_heading(element, int(tag_name[1]))
        elif tag_name == "table":
            self._render_table(element)
        elif tag_name == "hr":
            self._render_hr()
        elif tag_name == "page_break":
            para = self.doc.add_paragraph()
            run = para.add_run()
            run.add_break(WD_BREAK.PAGE)
        elif tag_name == "br":
            self.doc.add_paragraph()  # 空行
        elif tag_name == "p":
            self._render_paragraph(element)
        elif tag_name == "div":
            self._render_div(element)
        elif tag_name in ("strong", "b", "em", "i", "span"):
            # 顶层的 inline 元素 → 创建段落
            para = self.doc.add_paragraph()
            self._render_inline(para, element)
        elif tag_name in ("ul", "ol"):
            self._render_list(element)
        else:
            # 其他标签: 递归处理子元素
            for child in element.children:
                self._render_element(child)

    # ----------------------------------------------------------
    # 标题渲染
    # ----------------------------------------------------------
    def _render_heading(self, element, level: int):
        heading = self.doc.add_heading(level=min(level, 9))
        heading.clear()
        self._render_inline_children(heading, element)
        # 调整标题样式
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

    # ----------------------------------------------------------
    # 段落渲染
    # ----------------------------------------------------------
    def _render_paragraph(self, element):
        para = self.doc.add_paragraph()
        css_class = element.get("class", [])

        if "ol_item" in css_class:
            num = element.get("data-num", "1")
            run = para.add_run(f"{num}. ")
            self._render_inline_children(para, element)
            para.paragraph_format.left_indent = Cm(1.0)
        elif "ul_item" in css_class:
            run = para.add_run("• ")
            self._render_inline_children(para, element)
            para.paragraph_format.left_indent = Cm(1.0)
        else:
            self._render_inline_children(para, element)

        # 解析 style 中的 text-align
        self._apply_paragraph_style(para, element)

    def _render_div(self, element):
        """处理 div 元素"""
        align = (element.get("align") or "").lower()
        style = element.get("style", "")

        # 检查 div 内部是否包含块级元素
        has_block = any(
            isinstance(c, Tag) and c.name in ("table", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "hr", "ul", "ol")
            for c in element.children
        )

        if has_block:
            # 包含块级子元素 → 递归处理每个子元素
            for child in element.children:
                if isinstance(child, Tag):
                    if child.name == "table":
                        # 传递父 div 的对齐信息
                        self._render_table(child, parent_align=align)
                    else:
                        self._render_element(child)
                elif isinstance(child, NavigableString):
                    text = str(child).strip()
                    if text:
                        para = self.doc.add_paragraph()
                        self._add_inline_text(para, text)
                        if align == "right":
                            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        elif align == "center":
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            # 纯 inline 内容 → 作为一个段落
            para = self.doc.add_paragraph()
            self._render_inline_children(para, element)
            if align == "right":
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif align == "center":
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            self._apply_paragraph_style(para, element)

    # ----------------------------------------------------------
    # 表格渲染 (核心)
    # ----------------------------------------------------------
    def _render_table(self, table_element, parent_align: str = ""):
        """将 HTML <table> 渲染为 Word 表格"""
        rows_data = []
        for tr in table_element.find_all("tr", recursive=False):
            # 也检查 thead/tbody 内的 tr
            cells = tr.find_all(["td", "th"], recursive=False)
            rows_data.append(cells)

        if not rows_data:
            # 可能 tr 在 thead/tbody 内
            for section in table_element.find_all(["thead", "tbody", "tfoot"], recursive=False):
                for tr in section.find_all("tr", recursive=False):
                    cells = tr.find_all(["td", "th"], recursive=False)
                    rows_data.append(cells)

        if not rows_data:
            return

        num_rows = len(rows_data)
        num_cols = max(len(row) for row in rows_data)

        if num_cols == 0:
            return

        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.autofit = True

        # 解析表格级样式
        table_style = table_element.get("style", "")
        has_border_none = "border: none" in table_style or "border:none" in table_style

        # 设置表格对齐
        if parent_align == "right":
            table.alignment = WD_TABLE_ALIGNMENT.RIGHT
        elif parent_align == "center":
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for row_idx, cells in enumerate(rows_data):
            for col_idx, cell_elem in enumerate(cells):
                if col_idx >= num_cols:
                    break
                cell = table.cell(row_idx, col_idx)

                # 清空默认段落
                cell.paragraphs[0].clear()

                # 渲染单元格内容
                self._render_cell_content(cell, cell_elem)

                # 应用单元格样式
                self._apply_cell_style(cell, cell_elem, has_border_none)

        # 如果表格整体无边框，移除所有边框
        if has_border_none:
            self._remove_table_borders(table)

    def _render_cell_content(self, cell, cell_elem):
        """渲染表格单元格的内容"""
        # 获取单元格的第一个段落
        para = cell.paragraphs[0]

        # 检查单元格内是否有块级元素
        block_elements = cell_elem.find_all(["div", "p", "br"], recursive=False)

        if not block_elements:
            # 简单 inline 内容
            self._render_inline_children(para, cell_elem)
        else:
            # 有块级内容 → 逐个处理
            first_para = True
            for child in cell_elem.children:
                if isinstance(child, NavigableString):
                    text = str(child).strip()
                    if text:
                        if first_para:
                            self._add_inline_text(para, text)
                            first_para = False
                        else:
                            new_para = cell.add_paragraph()
                            self._add_inline_text(new_para, text)
                elif isinstance(child, Tag):
                    if child.name == "br":
                        if first_para:
                            first_para = False
                        else:
                            para = cell.add_paragraph()
                    elif child.name == "div":
                        target_para = para if first_para else cell.add_paragraph()
                        first_para = False
                        self._render_div_in_cell(target_para, child, cell)
                    elif child.name in ("strong", "b"):
                        target_para = para if first_para else cell.add_paragraph()
                        first_para = False
                        run = target_para.add_run(child.get_text())
                        run.bold = True
                    else:
                        target_para = para if first_para else cell.add_paragraph()
                        first_para = False
                        self._render_inline(target_para, child)

    def _render_div_in_cell(self, para, div_elem, cell):
        """在单元格内渲染 div"""
        style = div_elem.get("style", "")

        # 检查是否有上边框（签名线）
        has_border_top = "border-top" in style

        # 如果有 border-top，先添加一条线
        if has_border_top:
            self._add_bottom_border_to_paragraph(para)

        # 检查字体样式（如手写体）
        is_cursive = "cursive" in style or "Brush Script" in style
        font_size_match = re.search(r"font-size:\s*(\d+)px", style)
        font_size = int(font_size_match.group(1)) if font_size_match else None

        # 渲染 div 内容
        for child in div_elem.children:
            if isinstance(child, NavigableString):
                text = str(child).strip()
                if text:
                    run = para.add_run(text)
                    if is_cursive:
                        run.font.name = "Segoe Script"  # 模拟手写体
                        run.italic = True
                    if font_size:
                        run.font.size = Pt(font_size * 0.75)  # px to pt 近似
            elif isinstance(child, Tag):
                if child.name == "br":
                    # 在同一个 cell 中新建段落
                    para = cell.add_paragraph()
                    if has_border_top:
                        # 后续段落不需要边框
                        pass
                elif child.name in ("strong", "b"):
                    run = para.add_run(child.get_text())
                    run.bold = True
                else:
                    self._render_inline(para, child)

    def _apply_cell_style(self, cell, cell_elem, table_border_none: bool):
        """应用单元格样式"""
        style = cell_elem.get("style", "")

        # 文字对齐
        if "text-align: left" in style or "text-align:left" in style:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif "text-align: center" in style or "text-align:center" in style:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif "text-align: right" in style or "text-align:right" in style:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # 垂直对齐
        if "vertical-align: middle" in style:
            cell.vertical_alignment = 1  # WD_ALIGN_VERTICAL.CENTER
        elif "vertical-align: bottom" in style:
            cell.vertical_alignment = 2  # WD_ALIGN_VERTICAL.BOTTOM

        # 宽度
        width_match = re.search(r"width:\s*(\d+)%", style)
        if width_match:
            # 按百分比设置 (基于 A4 可用宽度约 16cm)
            pct = int(width_match.group(1))
            cell.width = Cm(16.0 * pct / 100)

        # 单元格级别的边框移除
        if "border: none" in style or "border:none" in style or table_border_none:
            self._remove_cell_borders(cell)

    def _remove_table_borders(self, table):
        """移除整个表格的所有边框"""
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
            tbl.insert(0, tblPr)

        borders_xml = f"""
        <w:tblBorders {nsdecls('w')}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
        """
        existing = tblPr.find(qn("w:tblBorders"))
        if existing is not None:
            tblPr.remove(existing)
        tblPr.append(parse_xml(borders_xml))

    def _remove_cell_borders(self, cell):
        """移除单个单元格的边框"""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        borders_xml = f"""
        <w:tcBorders {nsdecls('w')}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tcBorders>
        """
        existing = tcPr.find(qn("w:tcBorders"))
        if existing is not None:
            tcPr.remove(existing)
        tcPr.append(parse_xml(borders_xml))

    def _add_bottom_border_to_paragraph(self, para):
        """为段落添加上边框线（模拟签名线 border-top）"""
        pPr = para._p.get_or_add_pPr()
        borders_xml = f"""
        <w:pBdr {nsdecls('w')}>
            <w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        </w:pBdr>
        """
        existing = pPr.find(qn("w:pBdr"))
        if existing is not None:
            pPr.remove(existing)
        pPr.append(parse_xml(borders_xml))

    # ----------------------------------------------------------
    # 水平线
    # ----------------------------------------------------------
    def _render_hr(self):
        para = self.doc.add_paragraph()
        pPr = para._p.get_or_add_pPr()
        borders_xml = f"""
        <w:pBdr {nsdecls('w')}>
            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>
        </w:pBdr>
        """
        pPr.append(parse_xml(borders_xml))

    # ----------------------------------------------------------
    # 列表渲染
    # ----------------------------------------------------------
    def _render_list(self, list_elem):
        is_ordered = list_elem.name == "ol"
        for idx, li in enumerate(list_elem.find_all("li", recursive=False), 1):
            para = self.doc.add_paragraph()
            prefix = f"{idx}. " if is_ordered else "• "
            para.add_run(prefix)
            self._render_inline_children(para, li)
            para.paragraph_format.left_indent = Cm(1.0)

    # ----------------------------------------------------------
    # Inline 内容渲染
    # ----------------------------------------------------------
    def _render_inline_children(self, para, element):
        """渲染元素的所有 inline 子节点到指定段落"""
        for child in element.children:
            if isinstance(child, NavigableString):
                text = str(child)
                # 保留空格但去除纯换行
                text = text.replace("\n", " ")
                if text.strip() or text == " ":
                    self._add_inline_text(para, text)
            elif isinstance(child, Tag):
                self._render_inline(para, child)

    def _render_inline(self, para, element):
        """渲染单个 inline 元素"""
        tag = element.name.lower()

        if tag in ("strong", "b"):
            text = element.get_text()
            run = para.add_run(text)
            run.bold = True
        elif tag in ("em", "i"):
            text = element.get_text()
            run = para.add_run(text)
            run.italic = True
        elif tag == "u":
            text = element.get_text()
            run = para.add_run(text)
            run.underline = True
        elif tag == "br":
            para.add_run().add_break()
        elif tag == "span":
            style = element.get("style", "")
            text = element.get_text()
            run = para.add_run(text)
            self._apply_run_style(run, style)
        elif tag == "a":
            text = element.get_text()
            run = para.add_run(text)
            run.font.color.rgb = RGBColor(0, 0, 238)
            run.underline = True
        elif tag == "sup":
            text = element.get_text()
            run = para.add_run(text)
            run.font.superscript = True
        elif tag == "sub":
            text = element.get_text()
            run = para.add_run(text)
            run.font.subscript = True
        elif tag == "div":
            # inline 上下文中的 div → 换行后继续
            self._render_inline_children(para, element)
        else:
            # 未知 inline 标签 → 提取文本
            self._render_inline_children(para, element)

    def _add_inline_text(self, para, text: str):
        """添加纯文本 run"""
        if text:
            para.add_run(text)

    def _apply_run_style(self, run, style: str):
        """从 CSS style 字符串中提取并应用 run 样式"""
        if "font-weight: bold" in style or "font-weight:bold" in style:
            run.bold = True
        if "font-style: italic" in style or "font-style:italic" in style:
            run.italic = True
        if "text-decoration: underline" in style:
            run.underline = True

        # 字体大小
        size_match = re.search(r"font-size:\s*(\d+)px", style)
        if size_match:
            run.font.size = Pt(int(size_match.group(1)) * 0.75)

        # 字体颜色
        color_match = re.search(r"color:\s*#([0-9a-fA-F]{6})", style)
        if color_match:
            hex_color = color_match.group(1)
            run.font.color.rgb = RGBColor(
                int(hex_color[0:2], 16),
                int(hex_color[2:4], 16),
                int(hex_color[4:6], 16),
            )

    def _apply_paragraph_style(self, para, element):
        """从元素 style 属性中提取段落样式"""
        style = element.get("style", "")
        if "text-align: center" in style or "text-align:center" in style:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif "text-align: right" in style or "text-align:right" in style:
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif "text-align: left" in style or "text-align:left" in style:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT


# ============================================================
# 3. 主函数
# ============================================================
def main():
    # ========== 配置 ==========
    API_KEY = os.getenv("OPENROUTER_API_KEY", "")
    FILE_PATH = r"C:\Users\Administrator\Desktop\医疗文件.pdf"
    OUTPUT_DOCX = r"C:\Users\Administrator\Desktop\ocr_output.docx"
    MODEL = "google/gemini-3-flash-preview"
    # ==========================

    # 检查文件是否存在
    if not API_KEY:
        print("鉂?鏈厤缃?OPENROUTER_API_KEY")
        sys.exit(1)

    if not os.path.exists(FILE_PATH):
        print(f"❌ 文件不存在: {FILE_PATH}")
        sys.exit(1)

    # Step 1: OCR
    raw_output = ocr_file(FILE_PATH, API_KEY, MODEL)

    # (可选) 保存原始输出以供调试
    raw_output_path = OUTPUT_DOCX.replace(".docx", "_raw.txt")
    with open(raw_output_path, "w", encoding="utf-8") as f:
        f.write(raw_output)
    print(f"📄 原始 OCR 输出已保存: {raw_output_path}")

    # Step 2: 转换为 Word
    html_output_path = OUTPUT_DOCX.replace(".docx", ".html")
    convert_text_to_word_via_libreoffice(
        raw_output,
        OUTPUT_DOCX,
        html_output_path=html_output_path,
    )

    print(f"\n🎉 全部完成！输出文件: {OUTPUT_DOCX}")


# ============================================================
# 4. 也可以单独使用转换器（跳过 OCR，直接处理已有文本）
# ============================================================
def convert_text_to_docx(raw_text: str, output_path: str):
    """
    便捷函数：直接将混合 HTML/Markdown 文本转换为 Word 文档

    Usage:
        convert_text_to_docx(your_ocr_output, "output.docx")
    """
    convert_text_to_word_via_libreoffice(raw_text, output_path)


if __name__ == "__main__":
    main()
