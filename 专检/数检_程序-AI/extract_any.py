"""
多格式文本提取器
支持: .docx / .xlsx / .pdf / .pptx
统一返回 list[Segment]
"""
from pathlib import Path
from dataclasses import dataclass


@dataclass
class Segment:
    source: str
    source_label: str
    text: str


def extract(path: str | Path) -> list[Segment]:
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".docx":
        return _from_docx(path)
    elif ext == ".xlsx":
        return _from_xlsx(path)
    elif ext == ".pdf":
        return _from_pdf(path)
    elif ext == ".pptx":
        return _from_pptx(path)
    else:
        raise ValueError(f"不支持的格式: {ext}")


# ── DOCX ──────────────────────────────────────────────────────────────
def _from_docx(path: Path) -> list[Segment]:
    from xml_full_content import extract_docx_in_order, SOURCE_LABELS
    raw = extract_docx_in_order(path)
    return [
        Segment(
            source=seg.source,
            source_label=SOURCE_LABELS.get(seg.source, seg.source),
            text=seg.text,
        )
        for seg in raw
    ]


# ── XLSX ──────────────────────────────────────────────────────────────
def _from_xlsx(path: Path) -> list[Segment]:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    segments: list[Segment] = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                text = str(val).strip()
                if text:
                    segments.append(Segment(
                        source="cell",
                        source_label=sheet.title,
                        text=text,
                    ))
    wb.close()
    return segments


# ── PPTX ──────────────────────────────────────────────────────────────
@dataclass
class _Placed:
    """形状及其在幻灯片坐标系下的位置（组合形状内部的子形状坐标需换算）。"""
    shape: object
    left: int
    top: int


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _child_transform(group, parent: tuple) -> tuple:
    """
    计算组合形状内部子形状 → 幻灯片坐标的换算 (sx, sy, ox, oy)：
    slide_x = ox + sx * child_x。子形状坐标系由 chOff/chExt 定义，需按 ext/chExt 缩放。
    读取失败时沿用父级换算（只影响阅读顺序，不影响是否提取）。
    """
    sx, sy, ox, oy = parent
    try:
        xfrm = group._element.grpSpPr.xfrm
        off, ext, ch_off, ch_ext = xfrm.off, xfrm.ext, xfrm.chOff, xfrm.chExt
        if None in (off, ext, ch_off, ch_ext):
            return parent
        kx = ext.cx / ch_ext.cx if ch_ext.cx else 1.0
        ky = ext.cy / ch_ext.cy if ch_ext.cy else 1.0
        return (
            sx * kx, sy * ky,
            ox + sx * (off.x - ch_off.x * kx),
            oy + sy * (off.y - ch_off.y * ky),
        )
    except Exception:
        return parent


def _flatten_shapes(shapes, transform: tuple = (1.0, 1.0, 0, 0)) -> list:
    """递归展开组合形状（含嵌套），返回带幻灯片绝对坐标的 _Placed 列表。"""
    from pptx.shapes.group import GroupShape

    sx, sy, ox, oy = transform
    placed: list[_Placed] = []
    for shape in shapes:
        if isinstance(shape, GroupShape):
            placed.extend(_flatten_shapes(shape.shapes, _child_transform(shape, transform)))
            continue
        placed.append(_Placed(
            shape=shape,
            left=int(ox + sx * (shape.left or 0)),
            top=int(oy + sy * (shape.top or 0)),
        ))
    return placed


def _paragraph_text(para) -> str:
    """
    段落文本：拼接 run，软换行(<a:br>)保留为 \\n，自动字段(页码/日期等)不计入。
    只用 run 拼接会丢掉换行，导致中文行与外文行粘连。
    """
    parts = []
    for child in para._p:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "r":
            parts.append(child.findtext(f"{{{_A_NS}}}t") or "")
        elif tag == "br":
            parts.append("\n")
    return "".join(parts).strip()


def _from_pptx(path: Path) -> list[Segment]:
    """
    按幻灯片顺序提取文本。
    每张幻灯片：标题优先，然后正文占位符，最后其他文本框和表格。
    组合形状（含嵌套）会被递归展开，其内部文本框/表格同样提取。
    source_label 格式：第N页 [标题前20字]
    """
    from pptx import Presentation

    prs = Presentation(path)
    segments: list[Segment] = []

    for slide_no, slide in enumerate(prs.slides, start=1):
        # 取幻灯片标题
        slide_title = ""
        title_shape = slide.shapes.title
        if title_shape is not None and title_shape.has_text_frame:
            slide_title = title_shape.text_frame.text.strip()
        label = f"第{slide_no}页" + (f" {slide_title[:20]}" if slide_title else "")

        # 按视觉位置排序：标题优先，其余按行分组后列内排序
        # 行分组：按 top 排序，相邻形状 top 差值在行容差内归为同一行
        # 行内再按列分组（left 差值在列容差内视为同列），列内按 top 排，列间按 left 排
        ROW_TOLERANCE = 457200   # ~1.2cm，容忍预排版的垂直微偏
        COL_TOLERANCE = 36000    # ~0.1cm，容忍水平亚像素偏差

        def _group_into_rows(items):
            by_top = sorted(items, key=lambda s: s.top)
            rows: list[list] = []
            for s in by_top:
                placed = False
                for row in rows:
                    row_top = min(rs.top for rs in row)
                    if abs(s.top - row_top) <= ROW_TOLERANCE:
                        row.append(s)
                        placed = True
                        break
                if not placed:
                    rows.append([s])
            # 行内按列分组后排序：列内按 top，列间按 left
            for row in rows:
                col_groups: list[list] = []
                for s in sorted(row, key=lambda s: s.left):
                    placed = False
                    for col in col_groups:
                        col_left = min(cs.left for cs in col)
                        if abs(s.left - col_left) <= COL_TOLERANCE:
                            col.append(s)
                            placed = True
                            break
                    if not placed:
                        col_groups.append([s])
                for col in col_groups:
                    col.sort(key=lambda s: s.top)
                row[:] = [s for col in col_groups for s in col]
            return rows

        all_placed = _flatten_shapes(slide.shapes)
        title_items = [p for p in all_placed
                       if title_shape is not None and p.shape._element is title_shape._element]
        other_items = [p for p in all_placed if p not in title_items]
        rows = _group_into_rows(other_items)
        ordered = title_items + [p for row in rows for p in row]

        for item in ordered:
            shape = item.shape
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = _paragraph_text(para)
                    if text:
                        segments.append(Segment(
                            source="body",
                            source_label=label,
                            text=text,
                        ))
            elif getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = "\n".join(
                            _paragraph_text(p) for p in cell.text_frame.paragraphs
                        ).strip()
                        if text:
                            segments.append(Segment(
                                source="table",
                                source_label=label,
                                text=text,
                            ))

    return segments




# ── PDF ───────────────────────────────────────────────────────────────
def _from_pdf(path: Path) -> list[Segment]:
    """
    用 pymupdf (fitz) 的 TextBlock 提取 PDF 文本。
    TextBlock 按坐标聚合，比按行提取更接近段落语义。
    """
    import fitz

    segments: list[Segment] = []
    doc = fitz.open(str(path))

    for i, page in enumerate(doc, start=1):
        label = f"第{i}页"
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            if b[6] != 0:  # 跳过图片块
                continue
            text = b[4].replace("\n", " ").strip()
            if not text:
                continue
            segments.append(Segment(source="body", source_label=label, text=text))

    doc.close()
    return segments


def align_pdf_segments(src: list[Segment], tgt: list[Segment]) -> list[tuple]:
    """
    按页分组后页内顺序对齐，多出的块插入 None 占位。
    适用于 PDF 和 PPTX 的跨语言对齐。
    """
    from collections import defaultdict

    def page_no(label: str) -> int:
        try:
            return int(label.split()[0].replace("第", "").replace("页", ""))
        except (ValueError, IndexError):
            return 0

    src_by_page: dict[int, list[Segment]] = defaultdict(list)
    tgt_by_page: dict[int, list[Segment]] = defaultdict(list)
    for s in src:
        src_by_page[page_no(s.source_label)].append(s)
    for s in tgt:
        tgt_by_page[page_no(s.source_label)].append(s)

    pairs: list[tuple] = []
    for p in sorted(set(src_by_page) | set(tgt_by_page)):
        sp = src_by_page.get(p, [])
        tp = tgt_by_page.get(p, [])
        for i in range(max(len(sp), len(tp))):
            pairs.append((
                sp[i] if i < len(sp) else None,
                tp[i] if i < len(tp) else None,
            ))

    return pairs
