"""
docx_dynamic_style_extractor.py
动态解析 DOCX（ZIP 包）内所有 XML 文件，自动发现并提取全部样式定义。
不依赖固定文件路径映射，完全根据 [Content_Types].xml 和 .rels 关系链动态发现。

依赖: pip install python-docx lxml
"""

from __future__ import annotations

import sys
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, Optional

from lxml import etree


# ──────────────────────────────────────────────
# 常用 Open XML 命名空间（动态发现时仍需识别标签）
# ──────────────────────────────────────────────

KNOWN_NAMESPACES: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

# Content-Type 到部件角色的映射（用于动态识别）
CONTENT_TYPE_ROLES: dict[str, str] = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml": "styles",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "document",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml": "header",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml": "footer",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml": "footnotes",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml": "endnotes",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml": "comments",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml": "numbering",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml": "settings",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml": "fontTable",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml": "webSettings",
    "application/vnd.openxmlformats-officedocument.theme+xml": "theme",
    # WPS / 旧版 Office 可能使用的变体
    "application/vnd.ms-word.stylesWithEffects+xml": "stylesWithEffects",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": "document",
}


# ──────────────────────────────────────────────
# 数据模型
# ──────────────────────────────────────────────

@dataclass
class ZipPartInfo:
    """ZIP 包内单个部件的元信息"""
    path: str                       # ZIP 内路径，如 word/styles.xml
    content_type: str               # MIME 类型
    role: str                       # 角色标识：styles / header / footer / document ...
    file_size: int = 0              # 文件大小（字节）


@dataclass
class StyleDetail:
    """单个样式的完整属性"""
    style_id: str
    name: str
    style_type: str                             # paragraph / character / table / numbering
    source_file: str = ""                       # 来自哪个 XML 文件
    is_default: bool = False
    is_custom: bool = False
    is_hidden: bool = False
    is_semi_hidden: bool = False
    is_quick_style: bool = False
    unhide_when_used: bool = False
    base_style: Optional[str] = None
    next_style: Optional[str] = None
    linked_style: Optional[str] = None          # 关联的字符/段落样式
    ui_priority: Optional[int] = None           # UI 排序优先级
    outline_level: Optional[int] = None
    # 段落属性
    alignment: Optional[str] = None
    line_spacing: Optional[str] = None
    space_before: Optional[str] = None
    space_after: Optional[str] = None
    indent_left: Optional[str] = None
    indent_right: Optional[str] = None
    indent_first_line: Optional[str] = None
    # 字符属性
    font_ascii: Optional[str] = None
    font_east_asia: Optional[str] = None
    font_hAnsi: Optional[str] = None
    font_cs: Optional[str] = None
    font_size: Optional[str] = None
    font_size_cs: Optional[str] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None
    strike: Optional[bool] = None
    color: Optional[str] = None
    highlight: Optional[str] = None
    # 使用位置追踪
    used_in: list[str] = field(default_factory=list)  # ["document", "header1", "footer2", ...]
    # 原始 XML 片段（供高级用户调试）
    raw_xml: str = ""


@dataclass
class ThemeInfo:
    """主题文件中的默认字体/颜色定义"""
    source_file: str = ""
    major_font_latin: Optional[str] = None      # 标题字体（西文）
    major_font_ea: Optional[str] = None          # 标题字体（东亚）
    minor_font_latin: Optional[str] = None       # 正文字体（西文）
    minor_font_ea: Optional[str] = None          # 正文字体（东亚）
    color_scheme_name: Optional[str] = None
    colors: dict[str, str] = field(default_factory=dict)  # dk1, lt1, accent1 ...


@dataclass
class DocDefaultsInfo:
    """文档默认段落/字符属性 (w:docDefaults)"""
    default_font_ascii: Optional[str] = None
    default_font_east_asia: Optional[str] = None
    default_font_size: Optional[str] = None
    default_spacing_after: Optional[str] = None
    default_spacing_line: Optional[str] = None


@dataclass
class ExtractionResult:
    """完整的动态提取结果"""
    file_path: str
    zip_entries: list[str] = field(default_factory=list)        # ZIP 内所有文件列表
    discovered_parts: list[ZipPartInfo] = field(default_factory=list)  # 动态发现的部件
    styles: list[StyleDetail] = field(default_factory=list)
    theme: Optional[ThemeInfo] = None
    doc_defaults: Optional[DocDefaultsInfo] = None
    # 按来源分组的样式使用情况
    usage_map: dict[str, set[str]] = field(default_factory=dict)  # {"header1.xml": {"Header", ...}}
    # 解析过程中的警告/日志
    warnings: list[str] = field(default_factory=list)


# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────

def _qn(ns_prefix: str, local_name: str) -> str:
    """构造 Clark notation 标签名，如 {http://...}style"""
    uri = KNOWN_NAMESPACES.get(ns_prefix, ns_prefix)
    return f"{{{uri}}}{local_name}"


def _get_attr(elem: etree._Element, ns_prefix: str, attr_name: str, default: str = "") -> str:
    """安全获取带命名空间的属性值"""
    key = _qn(ns_prefix, attr_name)
    return elem.get(key, default)


def _find(parent: etree._Element, ns_prefix: str, tag: str) -> Optional[etree._Element]:
    """安全查找子元素"""
    return parent.find(_qn(ns_prefix, tag))


def _findall(parent: etree._Element, ns_prefix: str, tag: str) -> list[etree._Element]:
    """安全查找所有匹配子元素"""
    return parent.findall(_qn(ns_prefix, tag))


def _bool_val(elem: Optional[etree._Element], ns_prefix: str = "w", attr: str = "val") -> Optional[bool]:
    """解析布尔型元素（存在即为 True，除非 val="0"/"false"）"""
    if elem is None:
        return None
    val = _get_attr(elem, ns_prefix, attr, "true")
    return val.lower() not in ("0", "false", "off")


def _twips_to_pt(twips_str: str) -> Optional[str]:
    """将 twips（1/20 磅）转为磅"""
    try:
        return f"{int(twips_str) / 20:.1f}磅"
    except (ValueError, TypeError):
        return None


def _half_pt_to_pt(half_pt_str: str) -> Optional[str]:
    """将半磅转为磅"""
    try:
        return f"{int(half_pt_str) / 2:.0f}磅"
    except (ValueError, TypeError):
        return None


# ──────────────────────────────────────────────
# 核心：动态 ZIP 解析器
# ──────────────────────────────────────────────

class DynamicDocxExtractor:
    """
    动态解析 DOCX ZIP 包：
    1. 读取 [Content_Types].xml 发现所有部件及其角色
    2. 读取 .rels 文件发现关系链
    3. 根据发现的部件动态解析样式、主题、页眉页脚等
    """

    ALIGNMENT_MAP: dict[str, str] = {
        "left": "左对齐", "center": "居中", "right": "右对齐",
        "both": "两端对齐", "distribute": "分散对齐",
        "start": "左对齐", "end": "右对齐",
    }

    def __init__(self, file_path: str | Path) -> None:
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")

        # 验证是否为有效 ZIP
        if not zipfile.is_zipfile(self.file_path):
            raise ValueError(f"文件不是有效的 ZIP/DOCX 格式: {self.file_path}")

        self.zf = zipfile.ZipFile(self.file_path, "r")
        self.result = ExtractionResult(file_path=str(self.file_path))

        # 缓存已解析的 XML 树
        self._xml_cache: dict[str, etree._Element] = {}

    def close(self) -> None:
        """释放 ZIP 文件句柄"""
        if self.zf:
            self.zf.close()

    def __enter__(self) -> "DynamicDocxExtractor":
        return self

    def __exit__(self, *args: Any) -> None:
        self.close()

    # ─────────────── 主流程 ───────────────

    def extract(self) -> ExtractionResult:
        """执行完整的动态提取"""
        # 第一步：列出 ZIP 内所有文件
        self.result.zip_entries = self.zf.namelist()

        # 第二步：解析 [Content_Types].xml，动态发现所有部件
        self._discover_parts()

        # 第三步：解析样式定义文件（可能有多个）
        self._parse_all_style_sources()

        # 第四步：解析主题文件
        self._parse_theme_files()

        # 第五步：扫描所有内容部件，追踪样式使用情况
        self._scan_style_usage()

        # 第六步：将使用信息回写到样式对象
        style_map = {s.style_id: s for s in self.result.styles}
        for source, style_ids in self.result.usage_map.items():
            for sid in style_ids:
                if sid in style_map:
                    if source not in style_map[sid].used_in:
                        style_map[sid].used_in.append(source)

        return self.result

    # ─────────────── 动态发现部件 ───────────────

    def _discover_parts(self) -> None:
        """解析 [Content_Types].xml 动态发现所有部件"""
        ct_path = "[Content_Types].xml"
        if ct_path not in self.result.zip_entries:
            self.result.warnings.append("未找到 [Content_Types].xml，尝试遍历所有 XML 文件")
            self._fallback_discover()
            return

        root = self._parse_xml(ct_path)
        if root is None:
            self._fallback_discover()
            return

        # 处理 <Override> 元素（精确匹配）
        for override in root:
            tag_local = etree.QName(override.tag).localname
            if tag_local == "Override":
                part_name = override.get("PartName", "")
                content_type = override.get("ContentType", "")
                # 去掉开头的 /
                normalized_path = part_name.lstrip("/")

                role = self._identify_role(content_type, normalized_path)
                if role:
                    size = self._get_entry_size(normalized_path)
                    part_info = ZipPartInfo(
                        path=normalized_path,
                        content_type=content_type,
                        role=role,
                        file_size=size,
                    )
                    self.result.discovered_parts.append(part_info)

        # 处理 <Default> 元素（按扩展名匹配，补充发现）
        default_types: dict[str, str] = {}
        for default in root:
            tag_local = etree.QName(default.tag).localname
            if tag_local == "Default":
                ext = default.get("Extension", "")
                ct = default.get("ContentType", "")
                default_types[ext] = ct

        # 检查是否有通过 Default 类型但未被 Override 覆盖的 XML 文件
        discovered_paths = {p.path for p in self.result.discovered_parts}
        for entry in self.result.zip_entries:
            if entry not in discovered_paths and entry.endswith(".xml"):
                ext = entry.rsplit(".", 1)[-1] if "." in entry else ""
                ct = default_types.get(ext, "")
                role = self._identify_role(ct, entry)
                if role:
                    size = self._get_entry_size(entry)
                    self.result.discovered_parts.append(
                        ZipPartInfo(path=entry, content_type=ct, role=role, file_size=size)
                    )

    def _fallback_discover(self) -> None:
        """兜底方案：当 Content_Types 不存在时，按文件名模式猜测角色"""
        import re

        patterns: list[tuple[str, str]] = [
            (r"word/styles\.xml$", "styles"),
            (r"word/stylesWithEffects\.xml$", "stylesWithEffects"),
            (r"word/document\.xml$", "document"),
            (r"word/header\d*\.xml$", "header"),
            (r"word/footer\d*\.xml$", "footer"),
            (r"word/footnotes\.xml$", "footnotes"),
            (r"word/endnotes\.xml$", "endnotes"),
            (r"word/comments\.xml$", "comments"),
            (r"word/numbering\.xml$", "numbering"),
            (r"word/settings\.xml$", "settings"),
            (r"word/fontTable\.xml$", "fontTable"),
            (r"word/theme/theme\d*\.xml$", "theme"),
        ]

        for entry in self.result.zip_entries:
            for pattern, role in patterns:
                if re.search(pattern, entry, re.IGNORECASE):
                    size = self._get_entry_size(entry)
                    self.result.discovered_parts.append(
                        ZipPartInfo(path=entry, content_type="(推断)", role=role, file_size=size)
                    )
                    break

    def _identify_role(self, content_type: str, file_path: str) -> str:
        """根据 Content-Type 和文件路径动态识别部件角色"""
        import re

        # 优先通过 Content-Type 识别
        role = CONTENT_TYPE_ROLES.get(content_type, "")
        if role:
            # 对于 header/footer，附加编号以区分
            if role == "header":
                match = re.search(r"header(\d+)", file_path, re.IGNORECASE)
                return f"header{match.group(1)}" if match else "header"
            if role == "footer":
                match = re.search(r"footer(\d+)", file_path, re.IGNORECASE)
                return f"footer{match.group(1)}" if match else "footer"
            return role

        # 兜底：通过文件路径模式推断
        lower_path = file_path.lower()
        if "styles" in lower_path and lower_path.endswith(".xml"):
            return "styles"
        if "theme" in lower_path and lower_path.endswith(".xml"):
            return "theme"
        if "header" in lower_path and lower_path.endswith(".xml"):
            return "header"
        if "footer" in lower_path and lower_path.endswith(".xml"):
            return "footer"

        return ""

    def _get_entry_size(self, path: str) -> int:
        """获取 ZIP 条目的解压后大小"""
        try:
            info = self.zf.getinfo(path)
            return info.file_size
        except KeyError:
            return 0

    # ─────────────── XML 解析工具 ───────────────

    def _parse_xml(self, zip_path: str) -> Optional[etree._Element]:
        """从 ZIP 中读取并解析 XML，带缓存"""
        if zip_path in self._xml_cache:
            return self._xml_cache[zip_path]

        try:
            with self.zf.open(zip_path) as f:
                data = f.read()
            # 使用 recover 模式，容错解析
            parser = etree.XMLParser(recover=True, remove_blank_text=True)
            root = etree.fromstring(data, parser)
            self._xml_cache[zip_path] = root
            return root
        except (KeyError, etree.XMLSyntaxError) as e:
            self.result.warnings.append(f"解析 {zip_path} 失败: {e}")
            return None

    # ─────────────── 样式解析 ───────────────

    def _parse_all_style_sources(self) -> None:
        """解析所有样式定义源（styles.xml / stylesWithEffects.xml 等）"""
        style_parts = [
            p for p in self.result.discovered_parts
            if p.role in ("styles", "stylesWithEffects")
        ]

        if not style_parts:
            self.result.warnings.append("⚠️ 未发现任何样式定义文件！")
            return

        seen_ids: set[str] = set()  # 去重（stylesWithEffects 可能与 styles 重复）

        for part in style_parts:
            root = self._parse_xml(part.path)
            if root is None:
                continue

            # 解析文档默认值 <w:docDefaults>
            doc_defaults = _find(root, "w", "docDefaults")
            if doc_defaults is not None and self.result.doc_defaults is None:
                self.result.doc_defaults = self._parse_doc_defaults(doc_defaults)

            # 遍历所有 <w:style> 元素
            for style_elem in _findall(root, "w", "style"):
                detail = self._parse_style_element(style_elem, source_file=part.path)
                if detail and detail.style_id not in seen_ids:
                    self.result.styles.append(detail)
                    seen_ids.add(detail.style_id)

    def _parse_doc_defaults(self, doc_defaults: etree._Element) -> DocDefaultsInfo:
        """解析 <w:docDefaults> 文档默认属性"""
        info = DocDefaultsInfo()

        # 默认字符属性
        rpr_default = _find(doc_defaults, "w", "rPrDefault")
        if rpr_default is not None:
            rpr = _find(rpr_default, "w", "rPr")
            if rpr is not None:
                rfonts = _find(rpr, "w", "rFonts")
                if rfonts is not None:
                    info.default_font_ascii = _get_attr(rfonts, "w", "ascii") or None
                    info.default_font_east_asia = _get_attr(rfonts, "w", "eastAsia") or None
                sz = _find(rpr, "w", "sz")
                if sz is not None:
                    info.default_font_size = _half_pt_to_pt(_get_attr(sz, "w", "val"))

        # 默认段落属性
        ppr_default = _find(doc_defaults, "w", "pPrDefault")
        if ppr_default is not None:
            ppr = _find(ppr_default, "w", "pPr")
            if ppr is not None:
                spacing = _find(ppr, "w", "spacing")
                if spacing is not None:
                    after = _get_attr(spacing, "w", "after")
                    if after:
                        info.default_spacing_after = _twips_to_pt(after)
                    line = _get_attr(spacing, "w", "line")
                    if line:
                        try:
                            info.default_spacing_line = f"{int(line) / 240:.2f}倍"
                        except ValueError:
                            info.default_spacing_line = line

        return info

    def _parse_style_element(self, elem: etree._Element, source_file: str) -> Optional[StyleDetail]:
        """解析单个 <w:style> 元素为 StyleDetail"""
        style_id = _get_attr(elem, "w", "styleId")
        if not style_id:
            return None

        style_type = _get_attr(elem, "w", "type", "unknown")

        # 样式名称
        name_elem = _find(elem, "w", "name")
        name = _get_attr(name_elem, "w", "val", style_id) if name_elem is not None else style_id

        detail = StyleDetail(
            style_id=style_id,
            name=name,
            style_type=style_type,
            source_file=source_file,
        )

        # 布尔属性
        detail.is_default = _get_attr(elem, "w", "default", "0") == "1"
        detail.is_custom = _get_attr(elem, "w", "customStyle", "0") == "1"

        # 子元素布尔属性
        detail.is_semi_hidden = _find(elem, "w", "semiHidden") is not None
        detail.unhide_when_used = _find(elem, "w", "unhideWhenUsed") is not None
        hidden_elem = _find(elem, "w", "hidden")
        detail.is_hidden = hidden_elem is not None or detail.is_semi_hidden
        detail.is_quick_style = _find(elem, "w", "qFormat") is not None

        # UI 优先级
        ui_priority = _find(elem, "w", "uiPriority")
        if ui_priority is not None:
            val = _get_attr(ui_priority, "w", "val")
            try:
                detail.ui_priority = int(val)
            except ValueError:
                pass

        # 基于 / 后续 / 关联样式
        based_on = _find(elem, "w", "basedOn")
        if based_on is not None:
            detail.base_style = _get_attr(based_on, "w", "val") or None

        next_elem = _find(elem, "w", "next")
        if next_elem is not None:
            detail.next_style = _get_attr(next_elem, "w", "val") or None

        link_elem = _find(elem, "w", "link")
        if link_elem is not None:
            detail.linked_style = _get_attr(link_elem, "w", "val") or None

        # 段落属性
        ppr = _find(elem, "w", "pPr")
        if ppr is not None:
            self._parse_ppr(ppr, detail)

        # 字符属性
        rpr = _find(elem, "w", "rPr")
        if rpr is not None:
            self._parse_rpr(rpr, detail)

        # 保存原始 XML 片段
        try:
            detail.raw_xml = etree.tostring(elem, encoding="unicode", pretty_print=True)
        except Exception:
            detail.raw_xml = ""

        return detail

    def _parse_ppr(self, ppr: etree._Element, detail: StyleDetail) -> None:
        """解析段落属性 <w:pPr>"""
        # 大纲级别
        outline_lvl = _find(ppr, "w", "outlineLvl")
        if outline_lvl is not None:
            val = _get_attr(outline_lvl, "w", "val")
            try:
                detail.outline_level = int(val)
            except ValueError:
                pass

        # 对齐
        jc = _find(ppr, "w", "jc")
        if jc is not None:
            raw = _get_attr(jc, "w", "val")
            detail.alignment = self.ALIGNMENT_MAP.get(raw, raw)

        # 间距
        spacing = _find(ppr, "w", "spacing")
        if spacing is not None:
            before = _get_attr(spacing, "w", "before")
            after = _get_attr(spacing, "w", "after")
            line = _get_attr(spacing, "w", "line")
            line_rule = _get_attr(spacing, "w", "lineRule", "auto")

            if before:
                detail.space_before = _twips_to_pt(before)
            if after:
                detail.space_after = _twips_to_pt(after)
            if line:
                try:
                    line_val = int(line)
                    if line_rule == "auto":
                        detail.line_spacing = f"{line_val / 240:.2f}倍"
                    else:
                        detail.line_spacing = f"{line_val / 20:.1f}磅"
                except ValueError:
                    detail.line_spacing = line

        # 缩进
        ind = _find(ppr, "w", "ind")
        if ind is not None:
            left = _get_attr(ind, "w", "left") or _get_attr(ind, "w", "start")
            right = _get_attr(ind, "w", "right") or _get_attr(ind, "w", "end")
            first_line = _get_attr(ind, "w", "firstLine")
            hanging = _get_attr(ind, "w", "hanging")

            if left:
                detail.indent_left = _twips_to_pt(left)
            if right:
                detail.indent_right = _twips_to_pt(right)
            if first_line:
                detail.indent_first_line = f"首行 {_twips_to_pt(first_line)}"
            elif hanging:
                detail.indent_first_line = f"悬挂 {_twips_to_pt(hanging)}"

    def _parse_rpr(self, rpr: etree._Element, detail: StyleDetail) -> None:
        """解析字符属性 <w:rPr>"""
        # 字体
        rfonts = _find(rpr, "w", "rFonts")
        if rfonts is not None:
            detail.font_ascii = _get_attr(rfonts, "w", "ascii") or None
            detail.font_east_asia = _get_attr(rfonts, "w", "eastAsia") or None
            detail.font_hAnsi = _get_attr(rfonts, "w", "hAnsi") or None
            detail.font_cs = _get_attr(rfonts, "w", "cs") or None

        # 字号
        sz = _find(rpr, "w", "sz")
        if sz is not None:
            detail.font_size = _half_pt_to_pt(_get_attr(sz, "w", "val"))

        sz_cs = _find(rpr, "w", "szCs")
        if sz_cs is not None:
            detail.font_size_cs = _half_pt_to_pt(_get_attr(sz_cs, "w", "val"))

        # 加粗
        detail.bold = _bool_val(_find(rpr, "w", "b"))

        # 斜体
        detail.italic = _bool_val(_find(rpr, "w", "i"))

        # 下划线
        u = _find(rpr, "w", "u")
        if u is not None:
            detail.underline = _get_attr(u, "w", "val") or "single"

        # 删除线
        detail.strike = _bool_val(_find(rpr, "w", "strike"))

        # 颜色
        color = _find(rpr, "w", "color")
        if color is not None:
            val = _get_attr(color, "w", "val")
            if val and val.lower() != "auto":
                detail.color = f"#{val.upper()}"

        # 高亮
        highlight = _find(rpr, "w", "highlight")
        if highlight is not None:
            detail.highlight = _get_attr(highlight, "w", "val") or None

    # ─────────────── 主题解析 ───────────────

    def _parse_theme_files(self) -> None:
        """解析所有动态发现的主题文件"""
        theme_parts = [p for p in self.result.discovered_parts if p.role == "theme"]

        for part in theme_parts:
            root = self._parse_xml(part.path)
            if root is None:
                continue

            theme = ThemeInfo(source_file=part.path)

            # 主题元素可能在不同层级，动态搜索
            # 查找 <a:fontScheme> 下的字体定义
            for elem in root.iter():
                local = etree.QName(elem.tag).localname

                match local:
                    case "majorFont":
                        for child in elem:
                            child_local = etree.QName(child.tag).localname
                            if child_local == "latin":
                                theme.major_font_latin = child.get("typeface")
                            elif child_local == "ea":
                                theme.major_font_ea = child.get("typeface")
                    case "minorFont":
                        for child in elem:
                            child_local = etree.QName(child.tag).localname
                            if child_local == "latin":
                                theme.minor_font_latin = child.get("typeface")
                            elif child_local == "ea":
                                theme.minor_font_ea = child.get("typeface")
                    case "clrScheme":
                        theme.color_scheme_name = elem.get("name") if elem is not None else None
                        for color_elem in elem:
                            color_name = etree.QName(color_elem.tag).localname
                            # 颜色值在子元素中
                            for val_elem in color_elem:
                                val = val_elem.get("val") or val_elem.get("lastClr", "")
                                if val:
                                    theme.colors[color_name] = f"#{val.upper()}"
                                    break

            self.result.theme = theme

    # ─────────────── 使用追踪 ───────────────

    def _scan_style_usage(self) -> None:
        """扫描所有内容部件，追踪每个样式的实际使用位置"""
        # 需要扫描的部件角色
        scan_roles = {"document", "footnotes", "endnotes", "comments"}
        # header/footer 也需要扫描（角色名带编号如 header1, footer2）
        content_parts = [
            p for p in self.result.discovered_parts
            if p.role in scan_roles or p.role.startswith("header") or p.role.startswith("footer")
        ]

        for part in content_parts:
            root = self._parse_xml(part.path)
            if root is None:
                continue

            used_styles: set[str] = set()

            # 遍历所有元素，查找样式引用
            for elem in root.iter():
                local = etree.QName(elem.tag).localname

                # 段落样式引用 <w:pStyle w:val="...">
                if local == "pStyle":
                    val = _get_attr(elem, "w", "val")
                    if val:
                        used_styles.add(val)

                # 字符样式引用 <w:rStyle w:val="...">
                elif local == "rStyle":
                    val = _get_attr(elem, "w", "val")
                    if val:
                        used_styles.add(val)

                # 表格样式引用 <w:tblStyle w:val="...">
                elif local == "tblStyle":
                    val = _get_attr(elem, "w", "val")
                    if val:
                        used_styles.add(val)

            # 使用文件名（不含路径前缀）作为来源标识
            source_label = Path(part.path).name
            self.result.usage_map[source_label] = used_styles

    # ─────────────── 便捷查询方法 ───────────────

    def get_styles_by_type(self, style_type: str) -> list[StyleDetail]:
        """按类型筛选样式"""
        return [s for s in self.result.styles if s.style_type == style_type]

    def get_heading_styles(self) -> list[StyleDetail]:
        """获取所有标题样式（有大纲级别的段落样式）"""
        return [
            s for s in self.result.styles
            if s.outline_level is not None and s.outline_level < 9
        ]


# ──────────────────────────────────────────────
# 报告输出
# ──────────────────────────────────────────────

class DynamicReporter:
    """动态提取结果的格式化报告"""

    @staticmethod
    def print_report(result: ExtractionResult) -> None:
        """打印完整的动态分析报告"""
        W = 110
        print(f"\n{'═' * W}")
        print(f"  📦 DOCX 动态样式提取报告")
        print(f"  📁 文件: {result.file_path}")
        print(f"{'═' * W}")

        # ── ZIP 包结构概览 ──
        print(f"\n  📂 ZIP 包内文件总数: {len(result.zip_entries)}")
        print(f"  🔍 动态发现的有效部件: {len(result.discovered_parts)}")
        print()
        for part in result.discovered_parts:
            size_kb = part.file_size / 1024
            print(f"     [{part.role:>20}]  {part.path:<45}  ({size_kb:.1f} KB)  {part.content_type}")

        # ── 文档默认值 ──
        if result.doc_defaults:
            dd = result.doc_defaults
            print(f"\n  📐 文档默认值 (docDefaults):")
            if dd.default_font_ascii:
                print(f"     默认西文字体: {dd.default_font_ascii}")
            if dd.default_font_east_asia:
                print(f"     默认东亚字体: {dd.default_font_east_asia}")
            if dd.default_font_size:
                print(f"     默认字号: {dd.default_font_size}")
            if dd.default_spacing_after:
                print(f"     默认段后间距: {dd.default_spacing_after}")
            if dd.default_spacing_line:
                print(f"     默认行距: {dd.default_spacing_line}")

        # ── 主题信息 ──
        if result.theme:
            t = result.theme
            print(f"\n  🎨 主题信息 (来自 {t.source_file}):")
            if t.major_font_latin or t.major_font_ea:
                print(f"     标题字体: {t.major_font_latin or '—'} / {t.major_font_ea or '—'}")
            if t.minor_font_latin or t.minor_font_ea:
                print(f"     正文字体: {t.minor_font_latin or '—'} / {t.minor_font_ea or '—'}")
            if t.colors:
                color_str = "  ".join(f"{k}={v}" for k, v in list(t.colors.items())[:8])
                print(f"     配色方案: {t.color_scheme_name or '—'}")
                print(f"     颜色: {color_str}")

        # ── 样式使用追踪 ──
        if result.usage_map:
            print(f"\n  📊 各部件样式使用情况:")
            for source, sids in sorted(result.usage_map.items()):
                if sids:
                    print(f"     {source}: {len(sids)} 个样式 → {', '.join(sorted(sids)[:10])}"
                          f"{'...' if len(sids) > 10 else ''}")

        # ── 按类型分组输出所有样式 ──
        type_groups: dict[str, list[StyleDetail]] = {}
        for s in result.styles:
            type_groups.setdefault(s.style_type, []).append(s)

        type_labels = {
            "paragraph": "📝 段落样式", "character": "🔤 字符样式",
            "table": "📊 表格样式", "numbering": "📋 列表/编号样式",
        }

        print(f"\n  📋 样式总数: {len(result.styles)}")

        for stype, styles in sorted(type_groups.items()):
            label = type_labels.get(stype, f"❓ {stype} 样式")
            print(f"\n{'─' * W}")
            print(f"  {label} （共 {len(styles)} 个）")
            print(f"{'─' * W}")

            # 按 UI 优先级排序，无优先级的排后面
            styles.sort(key=lambda x: (x.ui_priority or 999, x.name))

            for i, s in enumerate(styles, 1):
                # 标签
                tags: list[str] = []
                if s.is_default:
                    tags.append("🔵默认")
                if s.is_quick_style:
                    tags.append("⚡快速")
                if s.is_hidden:
                    tags.append("👁️隐藏")
                if s.is_custom:
                    tags.append("✏️自定义")
                else:
                    tags.append("📦内置")
                if s.outline_level is not None and s.outline_level < 9:
                    tags.append(f"📑标题{s.outline_level + 1}")
                if s.used_in:
                    usage_str = ",".join(s.used_in)
                    tags.append(f"✅用于:{usage_str}")

                print(f"\n  {i:>3}. 【{s.name}】 (ID: {s.style_id})")
                print(f"       {' '.join(tags)}")

                # 关系
                rels: list[str] = []
                if s.base_style:
                    rels.append(f"基于→{s.base_style}")
                if s.next_style:
                    rels.append(f"后续→{s.next_style}")
                if s.linked_style:
                    rels.append(f"关联→{s.linked_style}")
                if s.ui_priority is not None:
                    rels.append(f"优先级={s.ui_priority}")
                if rels:
                    print(f"       关系: {' | '.join(rels)}")

                # 格式属性
                fmt: list[str] = []
                fonts = [f for f in [s.font_east_asia, s.font_ascii, s.font_hAnsi] if f]
                if fonts:
                    fmt.append(f"字体: {'/'.join(dict.fromkeys(fonts))}")
                if s.font_size:
                    fmt.append(f"字号: {s.font_size}")
                if s.bold is not None:
                    fmt.append(f"加粗: {'是' if s.bold else '否'}")
                if s.italic is not None:
                    fmt.append(f"斜体: {'是' if s.italic else '否'}")
                if s.underline:
                    fmt.append(f"下划线: {s.underline}")
                if s.color:
                    fmt.append(f"颜色: {s.color}")
                if s.alignment:
                    fmt.append(f"对齐: {s.alignment}")
                if s.line_spacing:
                    fmt.append(f"行距: {s.line_spacing}")
                if s.space_before:
                    fmt.append(f"段前: {s.space_before}")
                if s.space_after:
                    fmt.append(f"段后: {s.space_after}")
                if s.indent_first_line:
                    fmt.append(f"缩进: {s.indent_first_line}")
                if fmt:
                    print(f"       格式: {' | '.join(fmt)}")

        # ── 警告信息 ──
        if result.warnings:
            print(f"\n{'─' * W}")
            print(f"  ⚠️ 解析警告 ({len(result.warnings)} 条):")
            for w in result.warnings:
                print(f"     • {w}")

        print(f"\n{'═' * W}")
        print(f"  ✅ 动态解析完毕")
        print(f"{'═' * W}\n")

    @staticmethod
    def export_to_csv(result: ExtractionResult, output_path: str | Path) -> None:
        """导出 CSV"""
        import csv

        output_path = Path(output_path)
        headers = [
            "序号", "样式ID", "样式名称", "类型", "来源文件",
            "内置", "隐藏", "快速样式", "默认", "自定义", "UI优先级",
            "基于", "后续样式", "关联样式", "大纲级别",
            "东亚字体", "西文字体", "字号", "加粗", "斜体", "下划线", "颜色",
            "对齐", "行距", "段前", "段后", "首行缩进",
            "使用位置",
        ]

        try:
            with output_path.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for idx, s in enumerate(result.styles, 1):
                    writer.writerow([
                        idx, s.style_id, s.name, s.style_type,
                        Path(s.source_file).name,
                        "是" if not s.is_custom else "否",
                        "是" if s.is_hidden else "否",
                        "是" if s.is_quick_style else "否",
                        "是" if s.is_default else "否",
                        "是" if s.is_custom else "否",
                        s.ui_priority if s.ui_priority is not None else "",
                        s.base_style or "", s.next_style or "", s.linked_style or "",
                        s.outline_level if s.outline_level is not None else "",
                        s.font_east_asia or "", s.font_ascii or "",
                        s.font_size or "",
                        "是" if s.bold else ("否" if s.bold is not None else ""),
                        "是" if s.italic else ("否" if s.italic is not None else ""),
                        s.underline or "", s.color or "",
                        s.alignment or "", s.line_spacing or "",
                        s.space_before or "", s.space_after or "",
                        s.indent_first_line or "",
                        ", ".join(s.used_in) if s.used_in else "",
                    ])
            print(f"\n  📁 CSV 已导出: {output_path.resolve()}")
        except IOError as e:
            print(f"\n  ❌ CSV 导出失败: {e}")


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────

def main(file_path: str, export_csv: bool = False, show_xml: bool = False) -> ExtractionResult:
    """
    主函数

    Args:
        file_path: DOCX 文件路径
        export_csv: 是否导出 CSV
        show_xml: 是否在报告中显示原始 XML 片段
    """
    with DynamicDocxExtractor(file_path) as extractor:
        result = extractor.extract()

    DynamicReporter.print_report(result)

    if export_csv:
        csv_path = Path(file_path).with_suffix(".styles.csv")
        DynamicReporter.export_to_csv(result, csv_path)

    return result


if __name__ == "__main__":
    target = r"C:\Users\H\Desktop\数检_程序-AI\测试文件\原文-含不可编辑_01 (2026-007)2025年年度报告.docx"
    result = main(target, export_csv=False)
    # if len(sys.argv) < 2:
    #     print("用法: python docx_dynamic_style_extractor.py <文件.docx> [--csv]")
    #     print()
    #     print("选项:")
    #     print("  --csv    同时导出 CSV 文件")
    #     print()
    #     print("示例:")
    #     print("  python docx_dynamic_style_extractor.py report.docx")
    #     print("  python docx_dynamic_style_extractor.py report.docx --csv")
    #     sys.exit(1)
    #
    # target = sys.argv[1]
    # do_csv = "--csv" in sys.argv
    #
    # try:
    #     result = main(target, export_csv=do_csv)
    #     print(f"  共发现 {len(result.discovered_parts)} 个部件，提取 {len(result.styles)} 个样式。")
    # except (FileNotFoundError, ValueError, RuntimeError) as e:
    #     print(f"❌ {e}")
    #     sys.exit(1)

