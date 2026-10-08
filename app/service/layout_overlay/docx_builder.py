"""独立 DOCX 构建器：原页面底图及可编辑的绝对定位文本框。"""
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

from .imaging import resolve_font, safe_overlay
from .models import OverlayPage

VML = "urn:schemas-microsoft-com:vml"
OFFICE = "urn:schemas-microsoft-com:office:office"


def _style(x, y, width, height, z):
    return (f"position:absolute;margin-left:{x:.3f}pt;margin-top:{y:.3f}pt;"
            f"width:{width:.3f}pt;height:{height:.3f}pt;z-index:{z};"
            "visibility:visible;mso-wrap-style:none;"
            "mso-position-horizontal-relative:page;mso-position-vertical-relative:page")


def _shape(paragraph, shape_id, style):
    run, pict = OxmlElement("w:r"), OxmlElement("w:pict")
    shape = etree.Element(f"{{{VML}}}shape", nsmap={"v": VML, "o": OFFICE})
    shape.set("id", shape_id)
    shape.set("type", "#_x0000_t202")
    shape.set("style", style)
    shape.set("stroked", "f")
    shape.set("filled", "f")
    pict.append(shape)
    run.append(pict)
    paragraph._p.append(run)
    return shape


def build_docx(pages: list[OverlayPage], output_path: str | Path, target_lang="en") -> str:
    if not pages:
        raise ValueError("没有可导出的页面")
    document = Document()
    normal = document.styles["Normal"]
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    for index, page in enumerate(pages):
        section = document.sections[0] if index == 0 else document.add_section(WD_SECTION.NEW_PAGE)
        section.page_width, section.page_height = Pt(page.width_pt), Pt(page.height_pt)
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Pt(0)
        section.header_distance = section.footer_distance = Pt(0)
        anchor = document.add_paragraph()
        anchor.paragraph_format.line_spacing = Pt(1)
        anchor.paragraph_format.space_before = anchor.paragraph_format.space_after = Pt(0)
        image = page.clean_path or page.image_path
        rid, _ = anchor.part.get_or_add_image(str(image))
        background = _shape(anchor, f"OverlayBackground{index + 1}", _style(0, 0, page.width_pt, page.height_pt, -251654144))
        background.attrib.pop("type", None)
        image_data = etree.Element(f"{{{VML}}}imagedata")
        image_data.set(qn("r:id"), rid)
        image_data.set(f"{{{OFFICE}}}title", "")
        background.append(image_data)
        sx, sy = page.width_pt / page.width, page.height_pt / page.height
        for segment in page.segments:
            if not safe_overlay(segment, page):
                continue
            x1, y1, x2, y2 = segment.bbox
            shape = _shape(anchor, segment.segment_id, _style(x1 * sx, y1 * sy, (x2 - x1) * sx, (y2 - y1) * sy, 1000))
            textbox = etree.Element(f"{{{VML}}}textbox")
            textbox.set("inset", "0,0,0,0")
            content = OxmlElement("w:txbxContent")
            paragraph = OxmlElement("w:p")
            props = OxmlElement("w:pPr")
            spacing = OxmlElement("w:spacing")
            for key, value in {"before": "0", "after": "0", "line": str(round(segment.font_size * 1.15 * 20)), "lineRule": "exact"}.items():
                spacing.set(qn(f"w:{key}"), value)
            props.append(spacing)
            if target_lang in {"ar", "fa", "he", "ur"}:
                props.append(OxmlElement("w:bidi"))
            paragraph.append(props)
            for line_index, text in enumerate((segment.fitted_text or segment.translation).split("\n")):
                run = OxmlElement("w:r")
                rprops = OxmlElement("w:rPr")
                fonts = OxmlElement("w:rFonts")
                _, font_name = resolve_font(segment.bold)
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    fonts.set(qn(f"w:{key}"), font_name)
                rprops.append(fonts)
                for tag in ("sz", "szCs"):
                    element = OxmlElement(f"w:{tag}")
                    element.set(qn("w:val"), str(round(segment.font_size * 2)))
                    rprops.append(element)
                color = OxmlElement("w:color")
                color.set(qn("w:val"), segment.color)
                rprops.append(color)
                if segment.bold:
                    rprops.append(OxmlElement("w:b"))
                if target_lang in {"ar", "fa", "he", "ur"}:
                    rprops.append(OxmlElement("w:rtl"))
                run.append(rprops)
                if line_index:
                    run.append(OxmlElement("w:br"))
                element = OxmlElement("w:t")
                element.set(qn("xml:space"), "preserve")
                element.text = text
                run.append(element)
                paragraph.append(run)
            content.append(paragraph)
            textbox.append(content)
            shape.append(textbox)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
    return str(output)
