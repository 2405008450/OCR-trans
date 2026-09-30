"""
xml_scan.py — 线性扫描方式提取 .docx 文本

与 xml_full_content.py 的区别：
  - 对 word/document.xml 做一次从头到尾的线性 iter() 遍历
  - 根据节点位置（祖先链）实时判断来源标签（正文/文本框/表格/脚注/尾注）
  - mc:AlternateContent 只取 mc:Choice 分支，跳过 mc:Fallback
  - 文本框内容紧跟其所在段落，顺序更自然
  - 完整支持自动编号（有序列表序号前缀）
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile
from typing import Optional
from lxml import etree
import copy

# ── 命名空间 ────────────────────────────────────────────────────────────────
_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC  = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_NS_MAP = {
    "w":   _W,
    "mc":  _MC,
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
}

def _qn(tag: str) -> str:
    """'w:p' → '{http://...}p'"""
    prefix, local = tag.split(":", 1)
    return f"{{{_NS_MAP[prefix]}}}{local}"


# ── 数据结构 ────────────────────────────────────────────────────────────────
@dataclass
class Segment:
    source: str           # body / textbox / table / footnote / endnote / header / footer
    text: str
    para_index: int = 0   # 全局顺序计数
    row_context: str = "" # 表格行：同行所有单元格 tab 拼接

SOURCE_LABELS = {
    "body":     "正文",
    "toc":      "目录",
    "textbox":  "文本框",
    "table":    "表格",
    "footnote": "脚注",
    "endnote":  "尾注",
    "header":   "页眉",
    "footer":   "页脚",
    "chart":    "图表",
}


# ── 自动编号常量 ────────────────────────────────────────────────────────────
_CN_NUMS = [
    "〇","一","二","三","四","五","六","七","八","九","十",
    "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
    "二十一","二十二","二十三","二十四","二十五","二十六","二十七","二十八","二十九","三十",
]
_FULLWIDTH_DIGITS = "０１２３４５６７８９"


# ── 编号格式化函数 ───────────────────────────────────────────────────────────
def _to_chinese_upper(n: int) -> str:
    _map = {1:"壹",2:"贰",3:"叁",4:"肆",5:"伍",6:"陆",7:"柒",8:"捌",9:"玖",
            10:"拾",20:"贰拾",30:"叁拾",100:"佰"}
    if n <= 0: return "零"
    if n <= 10: return _map.get(n, str(n))
    if n < 20: return "拾" + _map.get(n - 10, str(n - 10))
    if n < 100:
        tens, ones = (n // 10) * 10, n % 10
        return _map.get(tens, str(tens)) + (_map.get(ones, str(ones)) if ones else "")
    return str(n)


def _to_chinese(n: int) -> str:
    if 0 <= n < len(_CN_NUMS): return _CN_NUMS[n]
    return str(n)


def _to_roman(n: int) -> str:
    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
    result = ""
    for v, s in vals:
        while n >= v:
            result += s; n -= v
    return result


def _to_alpha(n: int, upper: bool = True) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r if upper else 97 + r) + result
    return result


def _format_num(n: int, num_fmt: str) -> str:
    if num_fmt in ("bullet", "none", ""): return ""
    if num_fmt == "decimal": return str(n)
    if num_fmt == "decimalFullWidth":
        return "".join(_FULLWIDTH_DIGITS[int(c)] for c in str(n))
    if num_fmt in ("chineseCounting", "chineseCountingThousand",
                   "japaneseCounting", "japaneseDigital",
                   "ideographTraditional", "ideographZodiac",
                   "taiwaneseCountingThousand", "taiwaneseCounting"):
        return _to_chinese(n)
    if num_fmt in ("chineseLegalSimplified", "chineseLegal",
                   "ideographLegalTraditional"):
        return _to_chinese_upper(n)
    if num_fmt == "upperLetter": return _to_alpha(n, upper=True)
    if num_fmt == "lowerLetter": return _to_alpha(n, upper=False)
    if num_fmt == "upperRoman":  return _to_roman(n)
    if num_fmt == "lowerRoman":  return _to_roman(n).lower()
    if num_fmt == "ordinal":
        suffixes = {1:"st", 2:"nd", 3:"rd"}
        return str(n) + suffixes.get(n if n < 20 else n % 10, "th")
    return str(n)


def _expand_num_template(lvl_text: str, level_counts: list[int],
                          level_fmts: list[str]) -> str:
    """展开 w:lvlText 模板，如 '%1.%2.' → '1.2.'"""
    result = lvl_text
    for ref in range(len(level_counts), 0, -1):
        ph = f"%{ref}"
        if ph in result:
            idx = ref - 1
            count = level_counts[idx] if idx < len(level_counts) else 0
            fmt   = level_fmts[idx]   if idx < len(level_fmts)   else "decimal"
            result = result.replace(ph, _format_num(count, fmt))
    return result


# ── 编号状态类 ───────────────────────────────────────────────────────────────
@dataclass
class _LevelDef:
    ilvl:     int
    num_fmt:  str
    lvl_text: str
    start:    int = 1


@dataclass
class _AbstractNum:
    abstract_num_id: str
    levels: dict[int, _LevelDef] = field(default_factory=dict)


@dataclass
class _NumInst:
    num_id:          str
    abstract_num_id: str
    level_overrides: dict[int, _LevelDef] = field(default_factory=dict)


class NumberingState:
    """跟踪文档编号计数，为每个有序段落生成序号前缀。"""

    def __init__(self, zf: ZipFile) -> None:
        self._abstract_nums: dict[str, _AbstractNum] = {}
        self._num_insts:     dict[str, _NumInst]     = {}
        self._counters:      dict[str, dict[int, int]] = {}
        self.style_num:      dict[str, tuple[str, int]] = {}
        self._load(zf)

    def _load(self, zf: ZipFile) -> None:
        # ── numbering.xml ────────────────────────────────────────────────────
        num_path = "word/numbering.xml"
        if num_path not in zf.namelist():
            return
        with zf.open(num_path) as f:
            root = etree.parse(f).getroot()

        for abn in root.findall(_qn("w:abstractNum")):
            aid = abn.get(_qn("w:abstractNumId"), "")
            ab  = _AbstractNum(abstract_num_id=aid)
            for lvl in abn.findall(_qn("w:lvl")):
                ilvl      = int(lvl.get(_qn("w:ilvl"), "0"))
                nf_el     = lvl.find(_qn("w:numFmt"))
                lt_el     = lvl.find(_qn("w:lvlText"))
                st_el     = lvl.find(_qn("w:start"))
                num_fmt   = nf_el.get(_qn("w:val"), "decimal") if nf_el is not None else "decimal"
                lvl_text  = lt_el.get(_qn("w:val"), "")        if lt_el is not None else ""
                start     = int(st_el.get(_qn("w:val"), "1"))  if st_el is not None else 1
                ab.levels[ilvl] = _LevelDef(ilvl=ilvl, num_fmt=num_fmt,
                                             lvl_text=lvl_text, start=start)
            self._abstract_nums[aid] = ab

        for num_el in root.findall(_qn("w:num")):
            nid     = num_el.get(_qn("w:numId"), "")
            abn_ref = num_el.find(_qn("w:abstractNumId"))
            aid     = abn_ref.get(_qn("w:val"), "") if abn_ref is not None else ""
            inst    = _NumInst(num_id=nid, abstract_num_id=aid)
            for lo in num_el.findall(_qn("w:lvlOverride")):
                olvl = int(lo.get(_qn("w:ilvl"), "0"))
                so   = lo.find(_qn("w:startOverride"))
                lo_lvl = lo.find(_qn("w:lvl"))
                if lo_lvl is not None:
                    nf_el    = lo_lvl.find(_qn("w:numFmt"))
                    lt_el    = lo_lvl.find(_qn("w:lvlText"))
                    st_el    = lo_lvl.find(_qn("w:start"))
                    num_fmt  = nf_el.get(_qn("w:val"), "decimal") if nf_el is not None else "decimal"
                    lvl_text = lt_el.get(_qn("w:val"), "")        if lt_el is not None else ""
                    start    = int(st_el.get(_qn("w:val"), "1"))  if st_el is not None else 1
                    inst.level_overrides[olvl] = _LevelDef(
                        ilvl=olvl, num_fmt=num_fmt, lvl_text=lvl_text, start=start)
                elif so is not None:
                    base = self._get_level_def(aid, olvl)
                    if base:
                        ov = copy.copy(base)
                        ov.start = int(so.get(_qn("w:val"), "1"))
                        inst.level_overrides[olvl] = ov
            self._num_insts[nid] = inst

        # ── styles.xml：样式绑定的编号（heading 1/2/3 等）────────────────────
        sty_path = "word/styles.xml"
        if sty_path not in zf.namelist():
            return
        with zf.open(sty_path) as f:
            sroot = etree.parse(f).getroot()
        for style in sroot.iter(_qn("w:style")):
            sid  = style.get(_qn("w:styleId"), "")
            ppr  = style.find(_qn("w:pPr"))
            if ppr is None: continue
            np_  = ppr.find(_qn("w:numPr"))
            if np_ is None: continue
            nid_el  = np_.find(_qn("w:numId"))
            ilvl_el = np_.find(_qn("w:ilvl"))
            if nid_el is None: continue
            snid  = nid_el.get(_qn("w:val"), "0")
            silvl = int(ilvl_el.get(_qn("w:val"), "0")) if ilvl_el is not None else 0
            if snid != "0":
                self.style_num[sid] = (snid, silvl)

    def _get_level_def(self, abstract_num_id: str, ilvl: int) -> Optional[_LevelDef]:
        ab = self._abstract_nums.get(abstract_num_id)
        return ab.levels.get(ilvl) if ab else None

    def get_level_def(self, num_id: str, ilvl: int) -> Optional[_LevelDef]:
        inst = self._num_insts.get(num_id)
        if inst is None: return None
        if ilvl in inst.level_overrides:
            return inst.level_overrides[ilvl]
        return self._get_level_def(inst.abstract_num_id, ilvl)

    def advance(self, num_id: str, ilvl: int) -> str:
        """推进计数，返回展开后的编号字符串；bullet 格式返回空字符串。"""
        inst = self._num_insts.get(num_id)
        if inst is None: return ""

        counters = self._counters.setdefault(num_id, {})

        # 重置更深级别
        for deeper in list(counters.keys()):
            if deeper > ilvl:
                del counters[deeper]

        lvl_def = self.get_level_def(num_id, ilvl)
        if lvl_def is None: return ""

        if ilvl not in counters:
            counters[ilvl] = lvl_def.start
        else:
            counters[ilvl] += 1

        if not lvl_def.lvl_text and lvl_def.num_fmt == "bullet":
            return ""

        max_ilvl = max(counters.keys())
        level_counts: list[int] = []
        level_fmts:   list[str] = []
        for i in range(max_ilvl + 1):
            ld = self.get_level_def(num_id, i)
            level_counts.append(counters.get(i, ld.start if ld else 1))
            level_fmts.append(ld.num_fmt if ld else "decimal")

        return _expand_num_template(lvl_def.lvl_text, level_counts, level_fmts)


def _get_para_num_info(
    para: etree._Element,
    style_num: dict[str, tuple[str, int]],
) -> tuple[Optional[str], Optional[int]]:
    """
    提取段落的自动编号信息 (numId, ilvl)。
    优先读段落内 w:numPr，其次从段落样式继承。
    返回 (None, None) 表示无自动编号。
    """
    ppr = para.find(_qn("w:pPr"))
    if ppr is not None:
        np_ = ppr.find(_qn("w:numPr"))
        if np_ is not None:
            ilvl_el  = np_.find(_qn("w:ilvl"))
            nid_el   = np_.find(_qn("w:numId"))
            if nid_el is not None:
                num_id = nid_el.get(_qn("w:val"), "0")
                ilvl   = int(ilvl_el.get(_qn("w:val"), "0")) if ilvl_el is not None else 0
                if num_id != "0":
                    return num_id, ilvl
                return None, None  # numId=0 显式清除

    if style_num and ppr is not None:
        pst_el = ppr.find(_qn("w:pStyle"))
        if pst_el is not None:
            sid = pst_el.get(_qn("w:val"), "")
            if sid in style_num:
                return style_num[sid]

    return None, None


# ── TOC 段落识别 ─────────────────────────────────────────────────────────────
# Word 目录条目样式名前缀（中英文均覆盖）
_TOC_STYLE_PREFIXES = ("TOC", "toc", "目录", "Contents")

def _is_toc_para(para: etree._Element) -> bool:
    """
    判断段落是否为目录条目，支持两种识别方式：
    1. 段落样式名以 TOC/toc/目录/Contents 开头（标准 Word TOC 样式）
    2. 段落内含有 HYPERLINK \\l _Toc 域代码（无论样式名是什么）
    """
    W_r         = _qn("w:r")
    W_instrText = _qn("w:instrText")
    W_hyperlink = _qn("w:hyperlink")

    # 方式1：样式名前缀
    ppr = para.find(_qn("w:pPr"))
    if ppr is not None:
        pst = ppr.find(_qn("w:pStyle"))
        if pst is not None:
            style_val = pst.get(_qn("w:val"), "")
            if any(style_val.startswith(p) for p in _TOC_STYLE_PREFIXES):
                return True

    # 方式2：段落内有 HYPERLINK/PAGEREF _Toc 域（w:instrText）
    # 检查 w:instrText
    for r in para.iter(W_r):
        instr = r.find(W_instrText)
        if instr is not None and instr.text:
            txt = instr.text.strip().upper()
            # HYPERLINK \l _Toc... 形式
            if txt.startswith("HYPERLINK") and "_TOC" in txt.upper():
                return True
            # PAGEREF _Toc... 形式（无 HYPERLINK 包装，直接用 PAGEREF 引用书签）
            if txt.startswith("PAGEREF") and "_TOC" in txt.upper():
                return True

    # 检查 w:hyperlink 的 w:anchor 属性（另一种 TOC 超链接形式）
    _R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for hl in para.iter(W_hyperlink):
        anchor = hl.get(f"{{{_R_NS}}}anchor", "") or hl.get("anchor", "")
        if anchor.upper().startswith("_TOC"):
            return True

    return False


def _collect_toc_text(para: etree._Element, skip_ids: set[int]) -> str:
    """
    收集目录条目段落的文本，剥除末尾的页码部分。

    目录条目结构（两种形态）：
      1. 未静态化：文字内容  [制表符]  PAGEREF域(begin→separate→[页码缓存run]→end)
      2. 已静态化：文字内容  [制表符]  纯数字run（由 convert_toc_to_static 展开）

    策略：
      - 跳过所有 PAGEREF 域（begin 到 end 之间的所有 run）
      - 收集其余 w:t 文本
      - rstrip 末尾连续的制表符和数字（处理静态化后的页码残留）
    """
    W_fldChar   = _qn("w:fldChar")
    W_instrText = _qn("w:instrText")
    W_t         = _qn("w:t")
    W_r         = _qn("w:r")
    fldCharType = _qn("w:fldCharType")

    # 先找出所有属于 PAGEREF 域的 run id（含域内的缓存 run）
    pageref_run_ids: set[int] = set()
    runs = [r for r in para.iter(W_r) if id(r) not in skip_ids]
    i = 0
    while i < len(runs):
        r = runs[i]
        fc = r.find(W_fldChar)
        if fc is not None and fc.get(fldCharType) == "begin":
            # 先向前快速扫 instrText（在 separate 之前），判断域类型
            # 一旦确认不是 PAGEREF/PAGE，立即跳过，避免把外层 TOC 域误标记
            is_pageref = False
            instr_found = False
            j = i + 1
            depth = 1
            while j < len(runs) and depth > 0:
                rj = runs[j]
                fc2 = rj.find(W_fldChar)
                if fc2 is not None:
                    ft = fc2.get(fldCharType)
                    if ft == "begin":
                        depth += 1
                    elif ft == "separate" and depth == 1:
                        # 只在 depth==1 的 separate 之前查 instrText
                        # 到 separate 为止都没找到 PAGEREF → 不是 PAGEREF 域
                        if not instr_found:
                            break
                    elif ft == "end":
                        depth -= 1
                if not instr_found:
                    instr = rj.find(W_instrText)
                    if instr is not None and instr.text:
                        instr_found = True
                        txt = instr.text.strip().upper()
                        if "PAGEREF" in txt or txt == "PAGE":
                            is_pageref = True
                        else:
                            # 确认不是 PAGEREF 域，立即停止，不标记任何 run
                            break
                j += 1

            if is_pageref:
                # 重新扫一遍找到 end，标记整个域的 run
                depth = 1
                j = i + 1
                while j < len(runs) and depth > 0:
                    rj = runs[j]
                    fc2 = rj.find(W_fldChar)
                    if fc2 is not None:
                        ft = fc2.get(fldCharType)
                        if ft == "begin":
                            depth += 1
                        elif ft == "end":
                            depth -= 1
                    j += 1
                for k in range(i, j):
                    if k < len(runs):
                        pageref_run_ids.add(id(runs[k]))
                i = j
                continue
        i += 1

    # 收集非 PAGEREF 域的文本
    parts: list[str] = []
    for wt in para.iter(W_t):
        if id(wt) in skip_ids or _in_fallback(wt):
            continue
        # 找到该 w:t 的父 run，若整个 run 属于 PAGEREF 域则跳过
        parent_r = wt.getparent()
        if parent_r is not None and id(parent_r) in pageref_run_ids:
            continue
        if wt.text:
            parts.append(wt.text)

    text = "".join(parts)

    # rstrip 末尾的制表符、空格和纯数字（静态化后的页码残留）
    import re
    text = re.sub(r'[\t \u00a0]*\d+\s*$', '', text)

    return text


# ── 段落文本收集 ─────────────────────────────────────────────────────────────
def _in_fallback(node: etree._Element) -> bool:
    """节点是否位于 mc:Fallback 内。"""
    for anc in node.iterancestors():
        if anc.tag == _qn("mc:Fallback"):
            return True
    return False


def _collect_para_text(para: etree._Element) -> str:
    """
    收集段落正文文本。
    跳过 w:drawing 内的内容（文本框由线性扫描单独处理）和 mc:Fallback 内容。
    使用 list() 保持所有节点引用，防止 lxml 迭代器 id 复用。
    """
    skip_ids: set[int] = set()
    for drawing in list(para.iter(_qn("w:drawing"))):
        for n in list(drawing.iter()):
            skip_ids.add(id(n))

    parts: list[str] = []
    for wt in list(para.iter(_qn("w:t"))):
        if id(wt) in skip_ids or _in_fallback(wt):
            continue
        if wt.text:
            parts.append(wt.text)
    return "".join(parts)


# ── 表格行上下文 ─────────────────────────────────────────────────────────────
def _row_context(tr: etree._Element) -> str:
    cells = []
    for tc in tr.iter(_qn("w:tc")):
        cell_text = "".join(
            wt.text for wt in tc.iter(_qn("w:t"))
            if wt.text and not _in_fallback(wt)
        )
        if cell_text.strip():
            cells.append(cell_text)
    return "\t".join(cells)


# ── 主扫描函数 ───────────────────────────────────────────────────────────────
def _scan_tree(
    root: etree._Element,
    num_state: Optional[NumberingState],
    chart_map: Optional[dict[str, etree._Element]] = None,
) -> list[Segment]:
    """
    线性扫描策略：
      1. 先用 list(root.iter()) 把所有节点固定（防止 lxml id 复用）
      2. 标记 mc:Fallback 后代为 skip
      3. 遇到 w:txbxContent → 整块收集为 textbox，子节点标记为 handled
      4. 遇到 w:tc → 整格收集为 table，子节点标记为 handled
      5. 遇到 w:p → 收集正文文本，若有编号则加前缀
      6. 遇到 c:chart → 就地插入图表文本（需传入 chart_map）
      7. 其余节点跳过（自然线性前进）
    """
    segments: list[Segment] = []
    para_idx = 0

    # 固定所有节点引用，id() 在列表存活期间稳定
    all_nodes: list[etree._Element] = list(root.iter())

    # 标记 Fallback 子树
    skip: set[int] = set()
    for node in all_nodes:
        if node.tag == _qn("mc:Fallback"):
            skip.add(id(node))
            for desc in node.iter():
                skip.add(id(desc))

    # 标记已被父节点整体处理的后代
    handled: set[int] = set()

    for node in all_nodes:
        if id(node) in skip or id(node) in handled:
            continue

        # ── w:txbxContent → textbox ─────────────────────────────────────────
        if node.tag == _qn("w:txbxContent"):
            for desc in list(node.iter()):
                handled.add(id(desc))
            parts: list[str] = []
            for wt in list(node.iter(_qn("w:t"))):
                if id(wt) in skip or not wt.text:
                    continue
                parts.append(wt.text)
            text = "".join(parts).strip()
            if text:
                segments.append(Segment(source="textbox", text=text, para_index=para_idx))
                para_idx += 1
            continue

        # ── w:tc → table ────────────────────────────────────────────────────
        if node.tag == _qn("w:tc"):
            for desc in list(node.iter()):
                handled.add(id(desc))
            tr   = node.getparent()
            rctx = _row_context(tr) if tr is not None and tr.tag == _qn("w:tr") else ""
            parts = []
            for wt in list(node.iter(_qn("w:t"))):
                if id(wt) in skip or not wt.text:
                    continue
                parts.append(wt.text)
            cell_text = "".join(parts).strip()
            if cell_text:
                segments.append(Segment(
                    source="table", text=cell_text,
                    para_index=para_idx, row_context=rctx,
                ))
                para_idx += 1
            continue

        # ── w:p → 正文 / 脚注 / 尾注 / 页眉 / 页脚 ──────────────────────────
        if node.tag == _qn("w:p"):
            # 判断来源
            source = "body"
            for anc in node.iterancestors():
                atag = anc.tag.split("}")[-1] if "}" in anc.tag else anc.tag
                if atag == "footnote": source = "footnote"; break
                if atag == "endnote":  source = "endnote";  break
                if atag == "hdr":      source = "header";   break
                if atag == "ftr":      source = "footer";   break

            # 编号前缀（仅正文段落）
            num_prefix = ""
            if num_state is not None and source == "body":
                num_id, ilvl = _get_para_num_info(node, num_state.style_num)
                if num_id is not None and ilvl is not None:
                    generated = num_state.advance(num_id, ilvl)
                    if generated:
                        num_prefix = generated + " "

            # 目录条目：剥除末尾页码再收集，source 标记为 toc
            if source == "body" and _is_toc_para(node):
                source = "toc"
                text = _collect_toc_text(node, skip)
            else:
                text = _collect_para_text(node)
            full_text = (num_prefix + text) if num_prefix else text
            if full_text.strip():
                segments.append(Segment(source=source, text=full_text, para_index=para_idx))
                para_idx += 1
            continue

        # ── c:chart → 就地插入图表文本 ─────────────────────────────────────
        _C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        _R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        if node.tag == f"{{{_C_NS}}}chart" and chart_map:
            r_id = node.get(f"{{{_R_NS}}}id")
            if r_id and r_id in chart_map:
                chart_segs = _extract_chart_segments(chart_map[r_id], para_idx)
                segments.extend(chart_segs)
                if chart_segs:
                    para_idx = chart_segs[-1].para_index + 1
            continue

        # 其他节点：继续线性前进（不做任何操作）

    return segments


# ── 公开接口 ─────────────────────────────────────────────────────────────────
def scan_docx(docx_path: str | Path) -> list[Segment]:
    """
    线性扫描 .docx，按文档顺序返回所有文本片段（含自动编号前缀）。
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    if docx_path.suffix.lower() != ".docx":
        raise ValueError(f"不是 .docx 文件: {docx_path}")

    with ZipFile(docx_path) as zf:
        # 初始化编号状态
        num_state = NumberingState(zf) if "word/numbering.xml" in zf.namelist() else None

        # 构建 rId → chart_root 映射（从 document.xml.rels 读取）
        chart_map: dict[str, etree._Element] = {}
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in zf.namelist():
            with zf.open(rels_path) as f:
                rels_root = etree.parse(f).getroot()
            _REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
            _CHART_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            for rel in rels_root.findall(f"{{{_REL_NS}}}Relationship"):
                if rel.get("Type") == _CHART_TYPE:
                    r_id = rel.get("Id")
                    target = "word/" + rel.get("Target", "")
                    if target in zf.namelist():
                        with zf.open(target) as cf:
                            chart_map[r_id] = etree.parse(cf).getroot()

        # 扫描主文档，图表就地插入
        with zf.open("word/document.xml") as f:
            doc_root = etree.parse(f).getroot()

    return _scan_tree(doc_root, num_state, chart_map)


# ── 图表文本提取 ─────────────────────────────────────────────────────────────
_C   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_C15 = "http://schemas.microsoft.com/office/drawing/2012/chart"
_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

def _extract_chart_segments(root: etree._Element, start_idx: int) -> list[Segment]:
    """
    从已解析的 chart XML root 中提取三类文本：
      1. 坐标轴标题 / 图表标题（<a:t> 直接写死在 XML 里）
      2. 数据点标签文字（<c15:dlblRangeCache> 扩展缓存，外部链接图表专用）
      3. 系列名 / 分类名（<c:ser><c:tx><c:strCache> 内嵌数据图表）
    """
    segments: list[Segment] = []
    idx = start_idx

    # ── 1. 坐标轴标题 & 图表标题（<a:t> 直接硬编码的文字）──────────────────
    # 只取 c:title 和 c:valAx/c:catAx/c:dateAx 下的 a:t，避免重复收集系列标签
    for title_node in root.findall(f".//{{{_C}}}title"):
        parts = [t.text for t in title_node.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
        text = "".join(parts).strip()
        if text:
            segments.append(Segment(source="chart", text=text, para_index=idx))
            idx += 1

    # ── 2. 数据点标签缓存（c15:dlblRangeCache，外部链接散点/折线图）─────────
    for cache in root.iter(f"{{{_C15}}}dlblRangeCache"):
        labels = []
        for pt in cache.findall(f"{{{_C}}}pt"):
            v = pt.find(f"{{{_C}}}v")
            if v is not None and v.text and v.text.strip():
                labels.append(v.text.strip())
        if labels:
            # 合并为一个 segment，用顿号分隔，方便对齐检查
            segments.append(Segment(
                source="chart",
                text="、".join(labels),
                para_index=idx,
            ))
            idx += 1

    # ── 3. 系列名 & 分类名（内嵌数据图表的 strCache）────────────────────────
    for str_cache in root.iter(f"{{{_C}}}strCache"):
        labels = []
        for pt in str_cache.findall(f"{{{_C}}}pt"):
            v = pt.find(f"{{{_C}}}v")
            if v is not None and v.text and v.text.strip():
                labels.append(v.text.strip())
        if labels:
            segments.append(Segment(
                source="chart",
                text="、".join(labels),
                para_index=idx,
            ))
            idx += 1

    return segments


# ── CLI 入口 ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    from collections import Counter

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    path = r"D:\project\数检_程序-AI\测试文件\雅本化学2025ESG报告文字稿-20260409.docx"

    try:
        segs = scan_docx(path)
        counts = Counter(s.source for s in segs)

        print(f"{'='*60}")
        print(f"  线性扫描完成，共 {len(segs)} 个片段")
        print(f"{'='*60}")
        for src, label in SOURCE_LABELS.items():
            if counts.get(src, 0):
                print(f"  {label}: {counts[src]}")
        print(f"{'='*60}\n")

        for i, seg in enumerate(segs, 1):
            label = SOURCE_LABELS.get(seg.source, seg.source)
            print(f"[{i:>5}] [{label}] {seg.text}")

    except (FileNotFoundError, ValueError) as e:
        print(f"错误: {e}")
