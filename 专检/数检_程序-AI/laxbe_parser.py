# -*- coding: utf-8 -*-
"""
合并提取Word文档中所有文字
包括：正文、表格、脚注、尾注、页眉、页脚、文本框、批注、图表(Chart)等
图表提取策略：优先提取数据标签的实际显示内容，而非硬提取原始数值
公式提取策略：将 OMML 数学公式转换为 LaTeX 格式

修复内容：
  - 中文上下标正确用 \\text{} 包裹（如 V_{\\text{总}}^{2}）
  - 上下标 base 判断逻辑增强
  - run 内中文文本自动检测并包裹
  - LaTeX 输出后处理：清理冗余空格和语法问题
"""
import os
import re
from typing import Dict, Optional, List, Tuple, Set
from zipfile import ZipFile

from docx import Document
from lxml import etree

# ---------------- XML命名空间 ----------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
V_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
C15_NS = "http://schemas.microsoft.com/office/drawing/2012/chart"
C16R_NS = "http://schemas.microsoft.com/office/drawing/2017/03/chart"

NAMESPACES = {
    "w": W_NS, "wp": WP_NS, "a": A_NS,
    "wps": WPS_NS, "v": V_NS, "r": R_NS, "m": M_NS,
    "c": C_NS, "rel": REL_NS,
    "c15": C15_NS, "c16r": C16R_NS,
}

# 需要跳过的容器标签集合
_SKIP_CONTAINER_TAGS = {
    f"{{{WPS_NS}}}txbxContent",
    f"{{{V_NS}}}textbox",
}

# 中文字符正则（CJK 统一汉字 + 扩展）
_CJK_RE = re.compile(r'[一-\u9fff㐀-\u4dbf\uf900-﫿]')

# 判断是否为纯 ASCII 数学内容（字母、数字、基本运算符）
_PURE_MATH_RE = re.compile(r'^[a-zA-Z0-9+\-*/=.,;:!?\'\"\\{}\[\]()^_ \t]+$')


# ============================================================
#  OMML → LaTeX 转换器
# ============================================================
class OmmlToLatex:
    """
    将 Word OMML (Office Math Markup Language) 递归转换为 LaTeX 字符串。

    支持的结构：
      m:f      → \\frac{}{}          分数
      m:rad    → \\sqrt{} / \\sqrt[n]{}  根号
      m:sSup   → ^{}                 上标
      m:sSub   → _{}                 下标
      m:sSubSup → _{}^{}            上下标
      m:nary   → \\sum / \\int 等    N元运算符
      m:d      → \\left(\\right)      定界符
      m:func   → \\sin{} 等          函数
      m:eqArr  → 方程组              aligned
      m:m      → 矩阵               matrix
      m:acc    → \\hat{} 等          重音符
      m:bar    → \\overline{}        上划线/下划线
      m:limLow → 下极限              \\underset
      m:limUpp → 上极限              \\overset
      m:groupChr → 花括号组          \\underbrace / \\overbrace
      m:borderBox → 边框盒           直接提取内容
      m:box    → 盒子                直接提取内容
      m:r      → 文本 run            直接文本（含字符映射 + 中文检测）
    """

    # OMML 字符 → LaTeX 映射（常见特殊字符）
    CHAR_MAP: Dict[str, str] = {
        # 希腊字母
        'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta',
        'ε': r'\epsilon', 'ζ': r'\zeta', 'η': r'\eta', 'θ': r'\theta',
        'ι': r'\iota', 'κ': r'\kappa', 'λ': r'\lambda', 'μ': r'\mu',
        'ν': r'\nu', 'ξ': r'\xi', 'π': r'\pi', 'ρ': r'\rho',
        'σ': r'\sigma', 'τ': r'\tau', 'υ': r'\upsilon', 'φ': r'\varphi',
        'χ': r'\chi', 'ψ': r'\psi', 'ω': r'\omega',
        'Α': r'\Alpha', 'Β': r'\Beta', 'Γ': r'\Gamma', 'Δ': r'\Delta',
        'Ε': r'\Epsilon', 'Ζ': r'\Zeta', 'Η': r'\Eta', 'Θ': r'\Theta',
        'Ι': r'\Iota', 'Κ': r'\Kappa', 'Λ': r'\Lambda', 'Μ': r'\Mu',
        'Ν': r'\Nu', 'Ξ': r'\Xi', 'Π': r'\Pi', 'Ρ': r'\Rho',
        'Σ': r'\Sigma', 'Τ': r'\Tau', 'Υ': r'\Upsilon', 'Φ': r'\Phi',
        'Χ': r'\Chi', 'Ψ': r'\Psi', 'Ω': r'\Omega',
        # 运算符
        '±': r'\pm', '∓': r'\mp', '×': r'\times', '÷': r'\div',
        '·': r'\cdot', '∗': r'\ast', '∘': r'\circ',
        # 关系符
        '≤': r'\leq', '≥': r'\geq', '≠': r'\neq', '≈': r'\approx',
        '≡': r'\equiv', '∼': r'\sim', '≅': r'\cong', '∝': r'\propto',
        '≪': r'\ll', '≫': r'\gg', '⊂': r'\subset', '⊃': r'\supset',
        '⊆': r'\subseteq', '⊇': r'\supseteq', '∈': r'\in', '∉': r'\notin',
        # 箭头
        '→': r'\rightarrow', '←': r'\leftarrow', '↔': r'\leftrightarrow',
        '⇒': r'\Rightarrow', '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow',
        # 杂项
        '∞': r'\infty', '∂': r'\partial', '∇': r'\nabla',
        '∀': r'\forall', '∃': r'\exists', '∅': r'\emptyset',
        '√': r'\sqrt', '∑': r'\sum', '∏': r'\prod',
        '∫': r'\int', '∬': r'\iint', '∭': r'\iiint',
        '…': r'\ldots', '⋯': r'\cdots', '⋮': r'\vdots', '⋱': r'\ddots',
        '′': "'", '″': "''",
        # 括号类
        '⟨': r'\langle', '⟩': r'\rangle',
        '{': r'\{', '}': r'\}',
        # 其他
        'ℝ': r'\mathbb{R}', 'ℤ': r'\mathbb{Z}', 'ℕ': r'\mathbb{N}',
        'ℂ': r'\mathbb{C}', 'ℚ': r'\mathbb{Q}',
    }

    # nary 运算符字符 → LaTeX 命令
    NARY_MAP: Dict[str, str] = {
        '∑': r'\sum', '∏': r'\prod', '∐': r'\coprod',
        '∫': r'\int', '∬': r'\iint', '∭': r'\iiint',
        '∮': r'\oint', '∯': r'\oiint', '∰': r'\oiiint',
        '⋃': r'\bigcup', '⋂': r'\bigcap',
        '⋁': r'\bigvee', '⋀': r'\bigwedge',
        '⨁': r'\bigoplus', '⨂': r'\bigotimes',
    }

    # 重音符映射
    ACC_MAP: Dict[str, str] = {
        '\u0302': 'hat',  # 尖帽 ^
        '̃': 'tilde',  # 波浪 ~
        '̀': 'grave',  # 反引号 `
        '́': 'acute',  # 正引号 '
        '\u0304': 'bar',  # 横线 -
        '̇': 'dot',  # 单点
        '̈': 'ddot',  # 双点
        '\u20d7': 'vec',  # 向量箭头
        '\u0306': 'breve',  # 短音符
        '\u030c': 'check',  # 倒尖帽
    }

    # 定界符映射
    DELIM_MAP: Dict[str, Tuple[str, str]] = {
        '(': (r'\left(', r'\right)'),
        '[': (r'\left[', r'\right]'),
        '{': (r'\left\{', r'\right\}'),
        '|': (r'\left|', r'\right|'),
        '‖': (r'\left\|', r'\right\|'),
        '⟨': (r'\left\langle', r'\right\rangle'),
        '⌊': (r'\left\lfloor', r'\right\rfloor'),
        '⌈': (r'\left\lceil', r'\right\rceil'),
    }

    def __init__(self):
        pass

    def convert(self, math_elem) -> str:
        """将 m:oMath 或 m:oMathPara 元素转换为 LaTeX"""
        local = self._local_name(math_elem)
        if local == "oMathPara":
            parts = []
            for child in math_elem:
                if self._local_name(child) == "oMath":
                    parts.append(self._process_node(child))
            raw = " \\\\ ".join(parts) if len(parts) > 1 else (parts[0] if parts else "")
        elif local == "oMath":
            raw = self._process_node(math_elem)
        else:
            raw = self._process_node(math_elem)
        return self._postprocess(raw)

    @staticmethod
    def _postprocess(latex: str) -> str:
        """后处理：清理 LaTeX 输出中的冗余内容"""
        s = latex
        # 合并多余空格
        s = re.sub(r' {2,}', ' ', s)
        # 移除 {} 空花括号（非 LaTeX 命令参数的）
        s = re.sub(r'(?<![\\_^{]){}(?![_^])', '', s)
        # 清理 \text{} 中只有空格的情况
        s = re.sub(r'\\text\{\s*\}', '', s)
        # 修复 \text 嵌套：\text{\text{x}} → \text{x}
        while r'\text{\text{' in s:
            s = s.replace(r'\text{\text{', r'\text{')
            # 需要同时去掉多余的 }
            # 简单处理：逐层替换
        # 移除首尾空格
        s = s.strip()
        return s

    def _local_name(self, elem) -> str:
        """获取元素的本地名称（去掉命名空间）"""
        if not isinstance(elem.tag, str):
            return ""
        return etree.QName(elem.tag).localname

    def _get_val(self, elem, attr_local: str) -> str:
        """获取 m:xxx 或 w:xxx 属性的 val 值"""
        val = elem.get(f"{{{M_NS}}}{attr_local}")
        if val is not None:
            return val
        val = elem.get(f"{{{W_NS}}}{attr_local}")
        if val is not None:
            return val
        val = elem.get(attr_local)
        return val or ""

    def _find_child(self, parent, local_name: str):
        """在直接子元素中查找指定本地名称的元素"""
        for child in parent:
            if self._local_name(child) == local_name:
                return child
        return None

    def _find_child_val(self, parent, local_name: str) -> str:
        """查找子元素并获取其 val 属性"""
        child = self._find_child(parent, local_name)
        if child is not None:
            return self._get_val(child, "val")
        return ""

    def _process_children(self, elem) -> str:
        """递归处理所有子元素并拼接结果"""
        parts: List[str] = []
        for child in elem:
            result = self._process_node(child)
            if result:
                parts.append(result)
        return "".join(parts)

    @staticmethod
    def _has_cjk(text: str) -> bool:
        """检测文本是否包含中日韩字符"""
        return bool(_CJK_RE.search(text))

    @staticmethod
    def _is_pure_cjk(text: str) -> bool:
        """检测文本是否全部为中日韩字符（允许空格）"""
        cleaned = text.replace(' ', '')
        return bool(cleaned) and all(
            _CJK_RE.match(ch) for ch in cleaned
        )

    def _wrap_text_if_needed(self, text: str) -> str:
        """
        如果文本包含中文或是多字符非数学标识符，用 \\text{} 包裹。

        规则：
        - 纯中文 → \\text{中文}
        - 中文混合英文 → \\text{混合内容}
        - 纯英文单字母 → 保持原样（数学变量）
        - 纯英文多字母且非 LaTeX 命令 → \\text{} 包裹
        - 纯数字 → 保持原样
        """
        if not text or not text.strip():
            return text

        stripped = text.strip()

        # 已经被 \text{} 包裹的，直接返回
        if stripped.startswith(r'\text{'):
            return text

        # 包含中文 → 必须用 \text{} 包裹
        if self._has_cjk(stripped):
            return f'\\text{{{stripped}}}'

        # 纯数字 → 保持原样
        if re.match(r'^[0-9.]+$', stripped):
            return text

        # 单个英文字母 → 数学变量，保持原样
        if re.match(r'^[a-zA-Z]$', stripped):
            return text

        # 包含 LaTeX 命令 → 保持原样
        if '\\' in stripped:
            return text

        return text

    def _map_char(self, text: str) -> str:
        """将文本中的特殊字符映射为 LaTeX"""
        result: List[str] = []
        for ch in text:
            if ch in self.CHAR_MAP:
                mapped = self.CHAR_MAP[ch]
                if mapped.startswith('\\'):
                    result.append(mapped + ' ')
                else:
                    result.append(mapped)
            else:
                result.append(ch)
        return "".join(result).rstrip()

    def _process_node(self, elem) -> str:
        """根据节点类型分派到对应的处理方法"""
        local = self._local_name(elem)

        match local:
            case "oMath":
                return self._process_children(elem)
            case "r":
                return self._handle_run(elem)
            case "f":
                return self._handle_fraction(elem)
            case "rad":
                return self._handle_radical(elem)
            case "sSup":
                return self._handle_superscript(elem)
            case "sSub":
                return self._handle_subscript(elem)
            case "sSubSup":
                return self._handle_sub_superscript(elem)
            case "nary":
                return self._handle_nary(elem)
            case "d":
                return self._handle_delimiter(elem)
            case "func":
                return self._handle_function(elem)
            case "eqArr":
                return self._handle_eq_array(elem)
            case "m":
                return self._handle_matrix(elem)
            case "acc":
                return self._handle_accent(elem)
            case "bar":
                return self._handle_bar(elem)
            case "limLow":
                return self._handle_lim_low(elem)
            case "limUpp":
                return self._handle_lim_upper(elem)
            case "groupChr":
                return self._handle_group_char(elem)
            case "borderBox" | "box":
                return self._handle_box(elem)
            case "sPre":
                return self._handle_pre_sub_sup(elem)
            case "phant":
                return self._handle_phantom(elem)
            # 属性节点，跳过
            case "rPr" | "ctrlPr" | "fPr" | "radPr" | "sSupPr" | "sSubPr" | \
                 "sSubSupPr" | "naryPr" | "dPr" | "funcPr" | "eqArrPr" | \
                 "mPr" | "accPr" | "barPr" | "limLowPr" | "limUppPr" | \
                 "groupChrPr" | "borderBoxPr" | "boxPr" | "sPrePr" | \
                 "phantPr" | "oMathParaPr" | "mcs" | "mc" | "mcPr":
                return ""
            # 容器节点：递归处理子元素
            case "e" | "num" | "den" | "deg" | "sup" | "sub" | "lim" | \
                 "fName" | "mr":
                return self._process_children(elem)
            case _:
                return self._process_children(elem)

    # ---------- 具体结构处理 ----------

    def _handle_run(self, elem) -> str:
        """
        处理 m:r（数学文本 run）

        关键改进：
        1. 检测 rPr 中的 m:nor（普通文本标记）→ 强制 \\text{}
        2. 检测 rPr 中的 w:rFonts 字体信息辅助判断
        3. 自动检测中文内容并用 \\text{} 包裹
        """
        texts: List[str] = []

        # 检查 run 属性
        rpr = self._find_child(elem, "rPr")
        is_normal_text = False  # m:nor 标记的普通文本
        is_literal = False  # 非斜体（可能是单位/文字）

        if rpr is not None:
            # m:nor → 明确标记为普通文本
            nor = self._find_child(rpr, "nor")
            if nor is not None:
                is_normal_text = True

            # 检查 w:rPr 子节点中的字体信息
            w_rpr = rpr.find(f"{{{W_NS}}}rPr")
            if w_rpr is None:
                w_rpr = rpr  # 有时属性直接在 m:rPr 下

            # 检查是否有 sty="p"（plain/正体）
            sty = self._find_child(rpr, "sty")
            if sty is not None:
                sty_val = self._get_val(sty, "val")
                if sty_val == "p":
                    is_literal = True

        for child in elem:
            cl = self._local_name(child)
            if cl == "t" and child.text:
                raw_text = child.text
                # 先做特殊字符映射
                mapped = self._map_char(raw_text)

                # 判断是否需要 \text{} 包裹
                if is_normal_text:
                    # 明确标记为普通文本
                    mapped = f"\\text{{{mapped}}}"
                elif self._has_cjk(raw_text):
                    # 包含中文字符 → 用 \text{} 包裹中文部分
                    mapped = self._wrap_cjk_segments(mapped)
                elif is_literal and len(raw_text) > 1 and raw_text.isalpha():
                    # 正体 + 多字母 → 可能是单位或函数名
                    # 检查是否是已知函数
                    known_funcs = {
                        "sin", "cos", "tan", "cot", "sec", "csc",
                        "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
                        "ln", "log", "lg", "exp", "lim", "max", "min",
                        "sup", "inf", "det", "dim", "ker", "gcd",
                    }
                    if raw_text.lower() not in known_funcs:
                        mapped = f"\\text{{{mapped}}}"

                texts.append(mapped)

        return "".join(texts)

    def _wrap_cjk_segments(self, text: str) -> str:
        """
        智能包裹中文片段：将连续中文用 \\text{} 包裹，英文/数字保持原样。

        示例：
            "V总2"  →  不会到这里（上下标由结构处理）
            "总功"  →  "\\text{总功}"
            "a的值"  →  "a\\text{的值}"
            "总"    →  "\\text{总}"
        """
        # 如果整段都是中文（含标点），整体包裹
        if self._is_pure_cjk(text):
            return f'\\text{{{text}}}'

        # 如果已经包含 LaTeX 命令，不做分段处理
        if '\\' in text:
            return f'\\text{{{text}}}'

        # 分段处理：连续中文 / 连续非中文
        segments: List[str] = []
        current = ""
        current_is_cjk = False

        for ch in text:
            ch_is_cjk = bool(_CJK_RE.match(ch))
            if current and ch_is_cjk != current_is_cjk:
                # 切换类型，输出当前段
                if current_is_cjk:
                    segments.append(f'\\text{{{current}}}')
                else:
                    segments.append(current)
                current = ""
            current += ch
            current_is_cjk = ch_is_cjk

        # 输出最后一段
        if current:
            if current_is_cjk:
                segments.append(f'\\text{{{current}}}')
            else:
                segments.append(current)

        return "".join(segments)

    def _handle_fraction(self, elem) -> str:
        """处理 m:f（分数）"""
        fpr = self._find_child(elem, "fPr")
        frac_type = ""
        if fpr is not None:
            frac_type = self._find_child_val(fpr, "type")

        num_elem = self._find_child(elem, "num")
        den_elem = self._find_child(elem, "den")
        numerator = self._process_node(num_elem) if num_elem is not None else ""
        denominator = self._process_node(den_elem) if den_elem is not None else ""

        match frac_type:
            case "skw":
                return f"{{{numerator}}}/{{{denominator}}}"
            case "lin":
                return f"{numerator}/{denominator}"
            case "noBar":
                return f"\\binom{{{numerator}}}{{{denominator}}}"
            case _:
                return f"\\frac{{{numerator}}}{{{denominator}}}"

    def _handle_radical(self, elem) -> str:
        """处理 m:rad（根号）"""
        radpr = self._find_child(elem, "radPr")
        deg_hide = False
        if radpr is not None:
            dh = self._find_child(radpr, "degHide")
            if dh is not None:
                deg_hide = self._get_val(dh, "val") in ("1", "true", "on")

        deg_elem = self._find_child(elem, "deg")
        e_elem = self._find_child(elem, "e")
        degree = self._process_node(deg_elem).strip() if deg_elem is not None else ""
        content = self._process_node(e_elem) if e_elem is not None else ""

        if deg_hide or not degree or degree == "2":
            return f"\\sqrt{{{content}}}"
        else:
            return f"\\sqrt[{degree}]{{{content}}}"

    def _wrap_base(self, base: str) -> str:
        """
        智能处理上下标的 base 部分：
        - 单个 ASCII 字母 → 不加花括号
        - 单个中文字符 → \\text{字} 并加花括号
        - 多字符 → 加花括号
        """
        stripped = base.strip()
        if not stripped:
            return base

        # 单个 ASCII 字母或数字
        if len(stripped) == 1 and stripped.isascii():
            return stripped

        # 内容已经是 LaTeX 命令/结构（如 \frac{}{} ）→ 加花括号
        if len(stripped) > 1:
            return f"{{{stripped}}}"

        # 单个非 ASCII 字符（如中文）→ 用 \text{} 包裹
        if self._has_cjk(stripped):
            return f"{{\\text{{{stripped}}}}}"

        return stripped

    def _wrap_script_content(self, content: str) -> str:
        """
        包裹上标/下标内容：如果包含中文，用 \\text{} 包裹。
        确保内容中的中文字符被正确处理。
        """
        stripped = content.strip()
        if not stripped:
            return content

        # 已经被 \text{} 包裹
        if stripped.startswith(r'\text{'):
            return content

        # 包含中文 → 包裹
        if self._has_cjk(stripped):
            # 如果是纯中文
            if self._is_pure_cjk(stripped):
                return f'\\text{{{stripped}}}'
            # 混合内容 → 分段包裹
            return self._wrap_cjk_segments(stripped)

        return content

    def _handle_superscript(self, elem) -> str:
        """处理 m:sSup（上标）"""
        e_elem = self._find_child(elem, "e")
        sup_elem = self._find_child(elem, "sup")
        base = self._process_node(e_elem) if e_elem is not None else ""
        superscript = self._process_node(sup_elem) if sup_elem is not None else ""

        # 包裹中文内容
        superscript = self._wrap_script_content(superscript)
        base_wrapped = self._wrap_base(base)

        return f"{base_wrapped}^{{{superscript}}}"

    def _handle_subscript(self, elem) -> str:
        """处理 m:sSub（下标）"""
        e_elem = self._find_child(elem, "e")
        sub_elem = self._find_child(elem, "sub")
        base = self._process_node(e_elem) if e_elem is not None else ""
        subscript = self._process_node(sub_elem) if sub_elem is not None else ""

        # 包裹中文内容
        subscript = self._wrap_script_content(subscript)
        base_wrapped = self._wrap_base(base)

        return f"{base_wrapped}_{{{subscript}}}"

    def _handle_sub_superscript(self, elem) -> str:
        """处理 m:sSubSup（同时有上下标）"""
        e_elem = self._find_child(elem, "e")
        sub_elem = self._find_child(elem, "sub")
        sup_elem = self._find_child(elem, "sup")
        base = self._process_node(e_elem) if e_elem is not None else ""
        subscript = self._process_node(sub_elem) if sub_elem is not None else ""
        superscript = self._process_node(sup_elem) if sup_elem is not None else ""

        # 包裹中文内容
        subscript = self._wrap_script_content(subscript)
        superscript = self._wrap_script_content(superscript)
        base_wrapped = self._wrap_base(base)

        return f"{base_wrapped}_{{{subscript}}}^{{{superscript}}}"

    def _handle_nary(self, elem) -> str:
        """处理 m:nary（N元运算符：求和、积分等）"""
        narypr = self._find_child(elem, "naryPr")
        op_char = "∫"
        sub_hide = False
        sup_hide = False

        if narypr is not None:
            chr_elem = self._find_child(narypr, "chr")
            if chr_elem is not None:
                op_char = self._get_val(chr_elem, "val") or op_char

            sh = self._find_child(narypr, "subHide")
            if sh is not None:
                sub_hide = self._get_val(sh, "val") in ("1", "true", "on")

            sph = self._find_child(narypr, "supHide")
            if sph is not None:
                sup_hide = self._get_val(sph, "val") in ("1", "true", "on")

        sub_elem = self._find_child(elem, "sub")
        sup_elem = self._find_child(elem, "sup")
        e_elem = self._find_child(elem, "e")

        latex_op = self.NARY_MAP.get(op_char, op_char)
        subscript = self._process_node(sub_elem).strip() if sub_elem is not None else ""
        superscript = self._process_node(sup_elem).strip() if sup_elem is not None else ""
        content = self._process_node(e_elem) if e_elem is not None else ""

        result = latex_op
        if subscript and not sub_hide:
            result += f"_{{{subscript}}}"
        if superscript and not sup_hide:
            result += f"^{{{superscript}}}"
        result += f" {content}"

        return result

    def _handle_delimiter(self, elem) -> str:
        """处理 m:d（定界符/括号）"""
        dpr = self._find_child(elem, "dPr")
        beg_chr = "("
        end_chr = ")"
        sep_chr = "|"

        if dpr is not None:
            bc = self._find_child(dpr, "begChr")
            if bc is not None:
                beg_chr = self._get_val(bc, "val") or beg_chr
            ec = self._find_child(dpr, "endChr")
            if ec is not None:
                end_chr = self._get_val(ec, "val") or end_chr
            sc = self._find_child(dpr, "sepChr")
            if sc is not None:
                sep_chr = self._get_val(sc, "val") or sep_chr

        e_parts: List[str] = []
        for child in elem:
            if self._local_name(child) == "e":
                e_parts.append(self._process_node(child))

        # 映射定界符
        if beg_chr in self.DELIM_MAP:
            left_delim, right_delim = self.DELIM_MAP[beg_chr]
        else:
            left_delim = f"\\left{beg_chr}"
            right_delim = f"\\right{end_chr}"

        # 自定义结束符
        if end_chr in self.DELIM_MAP:
            _, right_delim = self.DELIM_MAP.get(end_chr, (None, f"\\right{end_chr}"))
        elif beg_chr not in self.DELIM_MAP:
            right_delim = f"\\right{end_chr}"

        latex_sep = f" {self._map_char(sep_chr).strip()} " if sep_chr else ", "
        content = latex_sep.join(e_parts)

        return f"{left_delim}{content}{right_delim}"

    def _handle_function(self, elem) -> str:
        """处理 m:func（数学函数：sin, cos, lim 等）"""
        fname_elem = self._find_child(elem, "fName")
        e_elem = self._find_child(elem, "e")

        func_name = self._process_node(fname_elem).strip() if fname_elem is not None else ""
        content = self._process_node(e_elem) if e_elem is not None else ""

        # 清理函数名中可能的 \text{} 包裹
        func_clean = re.sub(r'\\text\{([^}]*)\}', r'\1', func_name).strip()

        known_funcs = {
            "sin", "cos", "tan", "cot", "sec", "csc",
            "arcsin", "arccos", "arctan",
            "sinh", "cosh", "tanh", "coth",
            "ln", "log", "lg", "exp",
            "lim", "sup", "inf", "max", "min",
            "det", "dim", "ker", "gcd", "lcm",
            "deg", "hom", "arg",
        }

        if func_clean.lower() in known_funcs:
            return f"\\{func_clean.lower()} {content}"
        else:
            return f"\\operatorname{{{func_clean}}} {content}"

    def _handle_eq_array(self, elem) -> str:
        """处理 m:eqArr（方程组/对齐方程）"""
        rows: List[str] = []
        for child in elem:
            if self._local_name(child) == "e":
                rows.append(self._process_node(child))

        if len(rows) <= 1:
            return rows[0] if rows else ""

        content = " \\\\ ".join(rows)
        return f"\\begin{{aligned}} {content} \\end{{aligned}}"

    def _handle_matrix(self, elem) -> str:
        """处理 m:m（矩阵）"""
        rows: List[List[str]] = []
        for child in elem:
            if self._local_name(child) == "mr":
                cells: List[str] = []
                for cell in child:
                    if self._local_name(cell) == "e":
                        cells.append(self._process_node(cell))
                rows.append(cells)

        if not rows:
            return ""

        row_strs = [" & ".join(row) for row in rows]
        content = " \\\\ ".join(row_strs)
        return f"\\begin{{matrix}} {content} \\end{{matrix}}"

    def _handle_accent(self, elem) -> str:
        """处理 m:acc（重音符号：hat, tilde, vec 等）"""
        accpr = self._find_child(elem, "accPr")
        acc_char = "\u0302"

        if accpr is not None:
            chr_elem = self._find_child(accpr, "chr")
            if chr_elem is not None:
                acc_char = self._get_val(chr_elem, "val") or acc_char

        e_elem = self._find_child(elem, "e")
        content = self._process_node(e_elem) if e_elem is not None else ""

        latex_cmd = self.ACC_MAP.get(acc_char, "hat")
        return f"\\{latex_cmd}{{{content}}}"

    def _handle_bar(self, elem) -> str:
        """处理 m:bar（上划线/下划线）"""
        barpr = self._find_child(elem, "barPr")
        pos = "top"
        if barpr is not None:
            pos_elem = self._find_child(barpr, "pos")
            if pos_elem is not None:
                pos = self._get_val(pos_elem, "val") or "top"

        e_elem = self._find_child(elem, "e")
        content = self._process_node(e_elem) if e_elem is not None else ""

        if pos == "bot":
            return f"\\underline{{{content}}}"
        return f"\\overline{{{content}}}"

    def _handle_lim_low(self, elem) -> str:
        """处理 m:limLow（下极限）"""
        e_elem = self._find_child(elem, "e")
        lim_elem = self._find_child(elem, "lim")
        base = self._process_node(e_elem) if e_elem is not None else ""
        limit = self._process_node(lim_elem) if lim_elem is not None else ""

        base_clean = re.sub(r'\\text\{([^}]*)\}', r'\1', base).strip()
        if base_clean == "lim":
            return f"\\lim_{{{limit}}}"
        return f"\\underset{{{limit}}}{{{base}}}"

    def _handle_lim_upper(self, elem) -> str:
        """处理 m:limUpp（上极限）"""
        e_elem = self._find_child(elem, "e")
        lim_elem = self._find_child(elem, "lim")
        base = self._process_node(e_elem) if e_elem is not None else ""
        limit = self._process_node(lim_elem) if lim_elem is not None else ""

        return f"\\overset{{{limit}}}{{{base}}}"

    def _handle_group_char(self, elem) -> str:
        """处理 m:groupChr（花括号组）"""
        gcpr = self._find_child(elem, "groupChrPr")
        pos = "bot"
        chr_val = "⏟"

        if gcpr is not None:
            pos_elem = self._find_child(gcpr, "pos")
            if pos_elem is not None:
                pos = self._get_val(pos_elem, "val") or "bot"
            chr_elem = self._find_child(gcpr, "chr")
            if chr_elem is not None:
                chr_val = self._get_val(chr_elem, "val") or chr_val

        e_elem = self._find_child(elem, "e")
        content = self._process_node(e_elem) if e_elem is not None else ""

        if pos == "top":
            return f"\\overbrace{{{content}}}"
        return f"\\underbrace{{{content}}}"

    def _handle_box(self, elem) -> str:
        """处理 m:borderBox / m:box（盒子容器）"""
        e_elem = self._find_child(elem, "e")
        return self._process_node(e_elem) if e_elem is not None else ""

    def _handle_pre_sub_sup(self, elem) -> str:
        """处理 m:sPre（前置上下标，如同位素符号）"""
        sub_elem = self._find_child(elem, "sub")
        sup_elem = self._find_child(elem, "sup")
        e_elem = self._find_child(elem, "e")

        subscript = self._process_node(sub_elem) if sub_elem is not None else ""
        superscript = self._process_node(sup_elem) if sup_elem is not None else ""
        base = self._process_node(e_elem) if e_elem is not None else ""

        subscript = self._wrap_script_content(subscript)
        superscript = self._wrap_script_content(superscript)

        return f"{{}}_{{{subscript}}}^{{{superscript}}}{base}"

    def _handle_phantom(self, elem) -> str:
        """处理 m:phant（幻影/占位符）"""
        e_elem = self._find_child(elem, "e")
        content = self._process_node(e_elem) if e_elem is not None else ""
        return f"\\phantom{{{content}}}"


# 全局单例
_omml_converter = OmmlToLatex()


def _convert_omml_to_latex(math_elem) -> str:
    """将 OMML 数学元素转换为 LaTeX 字符串"""
    try:
        result = _omml_converter.convert(math_elem)
        # 清理多余空格
        result = re.sub(r'\s+', ' ', result).strip()
        return result
    except Exception as e:
        # 转换失败时回退到纯文本
        texts = []
        for node in math_elem.iter():
            local = etree.QName(node.tag).localname if isinstance(node.tag, str) else ""
            if local == "t" and node.text:
                texts.append(node.text)
        return "".join(texts)


# ============================================================
#  XML 文本提取（集成 LaTeX 公式）
# ============================================================
def _get_xml_text(element, *, skip_chart_drawing: bool = True,
                  skip_textbox: bool = True) -> str:
    """
    从 XML 元素中提取所有可见文字。
    数学公式转换为 LaTeX 格式，用 $ 包裹。
    """
    SYMBOL_CHAR_MAP = {
        'F0B4': '×', 'F0B8': '÷', 'F0B1': '±', 'F0B3': '≥',
        'F0A3': '≤', 'F0B9': '≠', 'F0BB': '≈',
        '\uf052': '☑', '\uf0a3': '☑', '': '☑', '': '☑',
        '\uf0a1': '☐', '\uf071': '☐', '\u25a1': '☐', '\u25fb': '☐',
        'F052': '☑', 'F0A1': '☐',
    }
    W14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"

    # 预先收集需要跳过的子树根节点 id
    skip_subtree_ids: Set[int] = set()

    if skip_chart_drawing:
        for drawing in element.iter(f"{{{W_NS}}}drawing"):
            if drawing.find(f".//{{{C_NS}}}chart") is not None:
                skip_subtree_ids.add(id(drawing))

    if skip_textbox:
        for tag in _SKIP_CONTAINER_TAGS:
            for container in element.iter(tag):
                skip_subtree_ids.add(id(container))

    # 收集所有 oMath/oMathPara 的 id，避免在 iter 时重复处理其内部节点
    math_elem_ids: Set[int] = set()
    for math_tag in (f"{{{M_NS}}}oMathPara", f"{{{M_NS}}}oMath"):
        for math_elem in element.iter(math_tag):
            math_elem_ids.add(id(math_elem))
            for desc in math_elem.iter():
                math_elem_ids.add(id(desc))

    def _should_skip(node) -> bool:
        if not skip_subtree_ids:
            return False
        if id(node) in skip_subtree_ids:
            return True
        parent = node.getparent()
        while parent is not None:
            if id(parent) in skip_subtree_ids:
                return True
            parent = parent.getparent()
        return False

    def _is_top_level_math(node) -> bool:
        """判断是否为顶层数学元素（不是另一个 oMath 的子孙）"""
        local = etree.QName(node.tag).localname if isinstance(node.tag, str) else ""
        if local not in ("oMath", "oMathPara"):
            return False
        parent = node.getparent()
        while parent is not None:
            p_local = etree.QName(parent.tag).localname if isinstance(parent.tag, str) else ""
            if p_local in ("oMath", "oMathPara"):
                return False
            parent = parent.getparent()
        return True

    texts: List[str] = []

    for node in element.iter():
        # 跳过文本框/图表内部
        if (skip_chart_drawing or skip_textbox) and _should_skip(node):
            continue

        node_local = etree.QName(node.tag).localname if isinstance(node.tag, str) else ""

        # --- 数学公式：顶层 oMath / oMathPara → LaTeX ---
        if node_local in ("oMath", "oMathPara") and node.tag.startswith(f"{{{M_NS}}}"):
            if _is_top_level_math(node):
                latex = _convert_omml_to_latex(node)
                if latex:
                    if node_local == "oMathPara":
                        texts.append(f" $${latex}$$ ")
                    else:
                        texts.append(f" ${latex}$ ")
            continue

        # 跳过数学元素的子孙
        if id(node) in math_elem_ids:
            continue

        # --- A. 标准文本 w:t（含父 w:r 的 vertAlign 上下标检测）---
        if node.tag == f"{{{W_NS}}}t":
            if node.text:
                t = node.text
                for raw, char in SYMBOL_CHAR_MAP.items():
                    if len(raw) > 2:
                        t = t.replace(raw, char)
                # 检查父 w:r 是否有 vertAlign
                parent = node.getparent()
                vert = ""
                if parent is not None and parent.tag == f"{{{W_NS}}}r":
                    rpr = parent.find(f"{{{W_NS}}}rPr")
                    if rpr is not None:
                        va = rpr.find(f"{{{W_NS}}}vertAlign")
                        if va is not None:
                            vert = va.get(f"{{{W_NS}}}val", "")
                if vert == "superscript":
                    texts.append(f"$^{{{t}}}$")
                elif vert == "subscript":
                    texts.append(f"$_{{{t}}}$")
                else:
                    texts.append(t)
            if node.tail:
                texts.append(node.tail)

        # --- B. DrawingML 文本 a:t ---
        elif node.tag == f"{{{A_NS}}}t":
            if node.text:
                texts.append(node.text)
            if node.tail:
                texts.append(node.tail)

        # --- C. 图片描述 ---
        elif node.tag == f"{{{WP_NS}}}docPr":
            alt = node.get("descr") or node.get("title")
            if alt:
                texts.append(f"[图片描述: {alt}]")

        # --- D. Symbol 符号 ---
        elif node.tag == f"{{{W_NS}}}sym":
            char_code = node.get(f"{{{W_NS}}}char", "").upper()
            if char_code in SYMBOL_CHAR_MAP:
                texts.append(SYMBOL_CHAR_MAP[char_code])
            elif char_code:
                try:
                    char_val = chr(int(char_code, 16))
                    texts.append(SYMBOL_CHAR_MAP.get(char_val, char_val))
                except (ValueError, OverflowError):
                    texts.append(f"[{char_code}]")

        # --- E. Checkbox ---
        elif node.tag.endswith("sdt"):
            checkbox = node.find(f".//{{{W14_NS}}}checkbox")
            if checkbox is not None:
                checked = checkbox.find(f".//{{{W14_NS}}}checked")
                val = checked.get(f"{{{W14_NS}}}val") if checked is not None else "0"
                texts.append("☑" if val in ["1", "true"] else "☐")

    return "".join(texts).strip()


# ============================================================
#  图表(Chart)提取系统
# ============================================================
class ChartExtractor:
    SEP = "\t"

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.rid_to_chart_path: Dict[str, str] = {}
        self.chart_texts: Dict[str, List[str]] = {}
        self._load_chart_rels()
        self._parse_all_charts()

    def _load_chart_rels(self) -> None:
        rels_path = "word/_rels/document.xml.rels"
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if rels_path not in zf.namelist():
                    return
                with zf.open(rels_path) as f:
                    tree = etree.parse(f)
                    for rel in tree.getroot():
                        rel_type = rel.get("Type", "")
                        target = rel.get("Target", "")
                        rid = rel.get("Id", "")
                        if "chart" in rel_type.lower() and rid:
                            chart_path = target if target.startswith("word/") else f"word/{target}"
                            self.rid_to_chart_path[rid] = chart_path
        except Exception as e:
            print(f"加载图表关系失败: {e}")

    def _parse_all_charts(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                for rid, chart_path in self.rid_to_chart_path.items():
                    if chart_path in zf.namelist():
                        with zf.open(chart_path) as f:
                            tree = etree.parse(f)
                            self.chart_texts[chart_path] = self._extract_chart_texts(tree.getroot())
                for name in zf.namelist():
                    if re.match(r"word/charts/chart\d*\.xml$", name) and name not in self.chart_texts:
                        with zf.open(name) as f:
                            tree = etree.parse(f)
                            self.chart_texts[name] = self._extract_chart_texts(tree.getroot())
        except Exception as e:
            print(f"解析图表文件失败: {e}")

    def _extract_datalabels_range_cache(self, ser_elem) -> List[str]:
        for dlbl_range in ser_elem.iter(f"{{{C15_NS}}}datalabelsRange"):
            cache = dlbl_range.find(f"{{{C15_NS}}}dlblRangeCache")
            if cache is not None:
                result = self._extract_cache_values(cache)
                if result:
                    return result
        for ext in ser_elem.iter(f"{{{C_NS}}}ext"):
            for child in ext.iter():
                tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
                if tag_local == "datalabelsRange":
                    for cache_child in child.iter():
                        cache_local = etree.QName(cache_child.tag).localname if isinstance(cache_child.tag, str) else ""
                        if cache_local == "dlblRangeCache":
                            result = self._extract_cache_values(cache_child)
                            if result:
                                return result
        return []

    def _extract_cache_values(self, cache_elem) -> List[str]:
        pt_map: Dict[int, str] = {}
        for pt in cache_elem.iter():
            tag_local = etree.QName(pt.tag).localname if isinstance(pt.tag, str) else ""
            if tag_local == "pt":
                idx_str = pt.get("idx", "") or pt.get(f"{{{C_NS}}}idx", "")
                try:
                    idx = int(idx_str) if idx_str else -1
                except ValueError:
                    idx = -1
                for v_child in pt:
                    if etree.QName(v_child.tag).localname == "v" and v_child.text:
                        if idx >= 0:
                            pt_map[idx] = v_child.text.strip()
                        else:
                            pt_map[len(pt_map)] = v_child.text.strip()
                        break
        if not pt_map:
            return []
        return [pt_map.get(i, "") for i in range(max(pt_map.keys()) + 1)]

    def _extract_chart_texts(self, root) -> List[str]:
        texts: List[str] = []
        title = self._get_chart_title(root)
        if title:
            texts.append(f"[图表标题] {title}")
        texts.extend(self._get_axis_texts(root))
        texts.extend(self._get_series_with_labels(root))
        fallback = self._get_all_drawingml_text(root)
        existing_set: Set[str] = set()
        for t in texts:
            for word in re.split(r'[\[\]\s]+', t):
                if word:
                    existing_set.add(word)
        for ft in fallback:
            if ft in existing_set or ft == "[CELLRANGE]":
                continue
            try:
                float(ft)
                continue
            except ValueError:
                pass
            texts.append(ft)
            existing_set.add(ft)
        return texts

    def _get_series_with_labels(self, root) -> List[str]:
        results: List[str] = []
        chart_type_tags = [
            "barChart", "bar3DChart", "lineChart", "line3DChart",
            "pieChart", "pie3DChart", "doughnutChart",
            "areaChart", "area3DChart", "scatterChart", "bubbleChart",
            "radarChart", "surfaceChart", "surface3DChart",
            "stockChart", "ofPieChart",
        ]
        for chart_type in chart_type_tags:
            for chart_elem in root.iter(f"{{{C_NS}}}{chart_type}"):
                for ser in chart_elem.iter(f"{{{C_NS}}}ser"):
                    results.extend(self._parse_series_smart(ser))
        return results

    def _parse_series_smart(self, ser_elem) -> List[str]:
        parts: List[str] = []
        sep = self.SEP
        tx = ser_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            name = self._get_str_or_ref(tx)
            if name:
                parts.append(f"  [系列名] {name}")
        cat = ser_elem.find(f"{{{C_NS}}}cat")
        if cat is not None:
            cat_vals = self._get_values(cat)
            if cat_vals:
                parts.append(f"  [分类] {sep.join(cat_vals)}")
        range_labels = self._extract_datalabels_range_cache(ser_elem)
        inline_labels = self._extract_inline_data_labels(ser_elem)
        has_labels = False
        if range_labels:
            non_empty = [lb for lb in range_labels if lb]
            if non_empty:
                has_labels = True
                parts.append(f"  [数据标签] {sep.join(range_labels)}")
        if not has_labels and inline_labels:
            meaningful = [lb for lb in inline_labels if lb and lb != "[CELLRANGE]"]
            if meaningful:
                has_labels = True
                parts.append(f"  [数据标签] {sep.join(meaningful)}")
        if not has_labels:
            for tag_name, label in [("val", "数据"), ("xVal", "X值"),
                                    ("yVal", "Y值"), ("bubbleSize", "气泡大小")]:
                elem = ser_elem.find(f"{{{C_NS}}}{tag_name}")
                if elem is not None:
                    vals = self._get_values(elem)
                    if vals:
                        parts.append(f"  [{label}] {sep.join(vals)}")
        return parts

    def _extract_inline_data_labels(self, ser_elem) -> List[str]:
        label_map: Dict[int, str] = {}
        for dlbl in ser_elem.iter(f"{{{C_NS}}}dLbl"):
            idx_elem = dlbl.find(f"{{{C_NS}}}idx")
            if idx_elem is None:
                continue
            try:
                idx = int(idx_elem.get("val", idx_elem.get(f"{{{C_NS}}}val", "-1")))
            except (ValueError, TypeError):
                continue
            parts = [at.text.strip() for at in dlbl.iter(f"{{{A_NS}}}t")
                     if at.text and at.text.strip()]
            if parts:
                label_map[idx] = "".join(parts)
        if not label_map:
            return []
        return [label_map.get(i, "") for i in range(max(label_map.keys()) + 1)]

    def _get_chart_title(self, root) -> Optional[str]:
        chart_elem = root.find(f"{{{C_NS}}}chart")
        if chart_elem is None:
            return None
        title_elem = chart_elem.find(f"{{{C_NS}}}title")
        if title_elem is None:
            return None
        tx = title_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            rich = tx.find(f"{{{C_NS}}}rich")
            if rich is not None:
                parts: List[str] = []
                for para in rich.findall(f"{{{A_NS}}}p"):
                    para_parts = [at.text for run in para.findall(f"{{{A_NS}}}r")
                                  for at in run.findall(f"{{{A_NS}}}t") if at.text]
                    if para_parts:
                        parts.append("".join(para_parts))
                if parts:
                    return "\n".join(parts).strip()
            str_ref = tx.find(f"{{{C_NS}}}strRef")
            if str_ref is not None:
                cache = str_ref.find(f"{{{C_NS}}}strCache")
                if cache is not None:
                    for pt in cache.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            return v.text.strip()
        parts = [at.text for at in title_elem.iter(f"{{{A_NS}}}t") if at.text]
        return "".join(parts).strip() if parts else None

    def _get_axis_texts(self, root) -> List[str]:
        results: List[str] = []
        axis_tags = [
            (f"{{{C_NS}}}catAx", "分类轴"), (f"{{{C_NS}}}valAx", "值轴"),
            (f"{{{C_NS}}}dateAx", "日期轴"), (f"{{{C_NS}}}serAx", "系列轴"),
        ]
        for axis_tag, axis_name in axis_tags:
            for axis in root.iter(axis_tag):
                title_elem = axis.find(f"{{{C_NS}}}title")
                if title_elem is not None:
                    parts = [at.text.strip() for at in title_elem.iter(f"{{{A_NS}}}t") if at.text]
                    if parts:
                        results.append(f"[{axis_name}标题] {''.join(parts)}")
        return results

    def _get_str_or_ref(self, elem) -> str:
        v = elem.find(f"{{{C_NS}}}v")
        if v is not None and v.text:
            return v.text.strip()
        str_ref = elem.find(f"{{{C_NS}}}strRef")
        if str_ref is not None:
            cache = str_ref.find(f"{{{C_NS}}}strCache")
            if cache is not None:
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    pv = pt.find(f"{{{C_NS}}}v")
                    if pv is not None and pv.text:
                        return pv.text.strip()
        parts = [at.text.strip() for at in elem.iter(f"{{{A_NS}}}t") if at.text]
        return "".join(parts)

    def _get_values(self, elem) -> List[str]:
        values: List[str] = []
        for cache in elem.iter(f"{{{C_NS}}}strCache"):
            for pt in cache.findall(f"{{{C_NS}}}pt"):
                v = pt.find(f"{{{C_NS}}}v")
                if v is not None and v.text:
                    values.append(v.text.strip())
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}numCache"):
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    v = pt.find(f"{{{C_NS}}}v")
                    if v is not None and v.text:
                        values.append(v.text.strip())
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}multiLvlStrCache"):
                for lvl in cache.findall(f"{{{C_NS}}}lvl"):
                    for pt in lvl.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            values.append(v.text.strip())
        if not values:
            for lit_tag in (f"{{{C_NS}}}strLit", f"{{{C_NS}}}numLit"):
                for lit in elem.iter(lit_tag):
                    for pt in lit.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            values.append(v.text.strip())
        return values

    def _get_all_drawingml_text(self, root) -> List[str]:
        texts: List[str] = []
        seen: Set[str] = set()
        for at in root.iter(f"{{{A_NS}}}t"):
            if at.text and at.text.strip():
                t = at.text.strip()
                if t not in seen:
                    texts.append(t)
                    seen.add(t)
        return texts

    def get_chart_text_by_rid(self, rid: str) -> List[str]:
        chart_path = self.rid_to_chart_path.get(rid, "")
        return self.chart_texts.get(chart_path, [])

    def get_all_chart_texts(self) -> Dict[str, List[str]]:
        return self.chart_texts


# ============================================================
#  页眉页脚
# ============================================================
def _extract_header_footer(doc_path: str) -> Tuple[List[str], List[str]]:
    headers: List[str] = []
    footers: List[str] = []
    try:
        with ZipFile(doc_path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("word/header"):
                    with zf.open(name) as f:
                        tree = etree.parse(f)
                        text = _get_xml_text(tree.getroot(), skip_chart_drawing=False,
                                             skip_textbox=False)
                        if text.strip():
                            headers.append(text)
                elif name.startswith("word/footer"):
                    with zf.open(name) as f:
                        tree = etree.parse(f)
                        text = _get_xml_text(tree.getroot(), skip_chart_drawing=False,
                                             skip_textbox=False)
                        if text.strip():
                            footers.append(text)
    except Exception as e:
        print(f"提取页眉页脚失败: {e}")
    return headers, footers


# ============================================================
#  锚点加载（脚注/尾注/批注）
# ============================================================
class DocAnchorsLoader:
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.footnotes: Dict[str, str] = {}
        self.endnotes: Dict[str, str] = {}
        self.comments: Dict[str, str] = {}
        self._load_all()

    def _load_xml_map(self, zf: ZipFile, filename: str, tag: str) -> Dict[str, str]:
        data: Dict[str, str] = {}
        if filename not in zf.namelist():
            return data
        with zf.open(filename) as f:
            tree = etree.parse(f)
            for elem in tree.findall(f".//w:{tag}", NAMESPACES):
                eid = elem.get(f"{{{W_NS}}}id")
                elem_type = elem.get(f"{{{W_NS}}}type")
                if elem_type in ("separator", "continuationSeparator"):
                    continue
                text = _get_xml_text(elem, skip_chart_drawing=False, skip_textbox=False)
                if eid and text:
                    data[eid] = text
        return data

    def _load_all(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                self.footnotes = self._load_xml_map(zf, "word/footnotes.xml", "footnote")
                self.endnotes = self._load_xml_map(zf, "word/endnotes.xml", "endnote")
                self.comments = self._load_xml_map(zf, "word/comments.xml", "comment")
        except Exception as e:
            print(f"加载锚点内容失败: {e}")


# ============================================================
#  锚定内容处理（文本框去重）
# ============================================================
def _process_anchored_content(p_element, loader: DocAnchorsLoader) -> List[str]:
    extras: List[str] = []
    txbx_content_ids: Set[int] = set()

    for txbx in p_element.iter(f"{{{WPS_NS}}}txbxContent"):
        txbx_content_ids.add(id(txbx))
        t = _get_xml_text(txbx, skip_chart_drawing=False, skip_textbox=False)
        if t:
            extras.append(t)

    for vtextbox in p_element.iter(f"{{{V_NS}}}textbox"):
        has_processed_child = any(
            id(child) in txbx_content_ids
            for child in vtextbox.iter(f"{{{WPS_NS}}}txbxContent")
        )
        if has_processed_child:
            continue
        t = _get_xml_text(vtextbox, skip_chart_drawing=False, skip_textbox=False)
        if t:
            extras.append(t)

    for ref in p_element.findall(".//w:footnoteReference", NAMESPACES):
        fid = ref.get(f"{{{W_NS}}}id")
        if fid and fid in loader.footnotes:
            extras.append(loader.footnotes[fid])

    for ref in p_element.findall(".//w:endnoteReference", NAMESPACES):
        eid = ref.get(f"{{{W_NS}}}id")
        if eid and eid in loader.endnotes:
            extras.append(loader.endnotes[eid])

    for ref in p_element.findall(".//w:commentReference", NAMESPACES):
        cid = ref.get(f"{{{W_NS}}}id")
        if cid and cid in loader.comments:
            extras.append(loader.comments[cid])

    return extras


# ============================================================
#  图表引用提取
# ============================================================
def _extract_chart_rids_from_paragraph(p_element) -> List[str]:
    rids: List[str] = []
    for chart_ref in p_element.iter(f"{{{C_NS}}}chart"):
        rid = chart_ref.get(f"{{{R_NS}}}id")
        if rid:
            rids.append(rid)
    return rids


# ============================================================
#  编号系统
# ============================================================
class NumberingSystem:
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map: Dict[str, Dict[str, Dict]] = {}
        self.abstract_num_map: Dict[str, Dict[str, Dict]] = {}
        self.level_counters: Dict[Tuple[str, str], int] = {}
        self.style_num_map: Dict[str, Tuple[str, str]] = {}
        self._load_numbering()
        self._load_style_numbering()

    def _load_numbering(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if "word/numbering.xml" not in zf.namelist():
                    return
                with zf.open("word/numbering.xml") as f:
                    tree = etree.parse(f)
                    for abstract_num in tree.findall(".//w:abstractNum", NAMESPACES):
                        abstract_num_id = abstract_num.get(f"{{{W_NS}}}abstractNumId")
                        self.abstract_num_map[abstract_num_id] = {}
                        for lvl in abstract_num.findall(".//w:lvl", NAMESPACES):
                            ilvl = lvl.get(f"{{{W_NS}}}ilvl")
                            num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                            lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                            start = lvl.find(".//w:start", NAMESPACES)
                            fmt_val = num_fmt.get(f"{{{W_NS}}}val") if num_fmt is not None else "decimal"
                            text_val = lvl_text.get(f"{{{W_NS}}}val") if lvl_text is not None else "%1."
                            start_val = int(start.get(f"{{{W_NS}}}val", "1")) if start is not None else 1
                            self.abstract_num_map[abstract_num_id][ilvl] = {
                                "format": fmt_val, "text": text_val, "start": start_val,
                            }
                    for num in tree.findall(".//w:num", NAMESPACES):
                        num_id = num.get(f"{{{W_NS}}}numId")
                        abstract_num_id_elem = num.find(".//w:abstractNumId", NAMESPACES)
                        if abstract_num_id_elem is None:
                            continue
                        abstract_num_id = abstract_num_id_elem.get(f"{{{W_NS}}}val")
                        if abstract_num_id in self.abstract_num_map:
                            level_map = {}
                            for k, v in self.abstract_num_map[abstract_num_id].items():
                                level_map[k] = v.copy()
                            for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                                ilvl = override.get(f"{{{W_NS}}}ilvl")
                                if ilvl is None:
                                    continue
                                lvl = override.find(f"{{{W_NS}}}lvl")
                                if lvl is not None:
                                    num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                                    lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                                    start = lvl.find(".//w:start", NAMESPACES)
                                    override_info = level_map.get(ilvl, {}).copy()
                                    if num_fmt is not None:
                                        override_info["format"] = num_fmt.get(f"{{{W_NS}}}val")
                                    if lvl_text is not None:
                                        override_info["text"] = lvl_text.get(f"{{{W_NS}}}val")
                                    if start is not None:
                                        override_info["start"] = int(start.get(f"{{{W_NS}}}val", "1"))
                                    level_map[ilvl] = override_info
                                start_override = override.find(f"{{{W_NS}}}startOverride")
                                if start_override is not None and ilvl in level_map:
                                    level_map[ilvl]["start"] = int(
                                        start_override.get(f"{{{W_NS}}}val", "1")
                                    )
                            self.numbering_map[num_id] = level_map
        except Exception as e:
            print(f"加载编号系统失败: {e}")

    def _load_style_numbering(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if "word/styles.xml" not in zf.namelist():
                    return
                with zf.open("word/styles.xml") as f:
                    tree = etree.parse(f)
                    for style in tree.findall(".//w:style", NAMESPACES):
                        style_id = style.get(f"{{{W_NS}}}styleId")
                        if not style_id:
                            continue
                        ppr = style.find(".//w:pPr", NAMESPACES)
                        if ppr is None:
                            continue
                        num_pr = ppr.find(".//w:numPr", NAMESPACES)
                        if num_pr is None:
                            continue
                        num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
                        ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
                        num_id = num_id_elem.get(f"{{{W_NS}}}val") if num_id_elem is not None else None
                        ilvl = ilvl_elem.get(f"{{{W_NS}}}val") if ilvl_elem is not None else "0"
                        if num_id:
                            self.style_num_map[style_id] = (num_id, ilvl)
        except Exception as e:
            print(f"加载样式编号失败: {e}")

    def _format_number(self, num: int, fmt: str) -> str:
        match fmt:
            case "decimal":
                return str(num)
            case "upperRoman":
                return self._to_roman(num).upper()
            case "lowerRoman":
                return self._to_roman(num).lower()
            case "upperLetter":
                return self._to_letter(num).upper()
            case "lowerLetter":
                return self._to_letter(num).lower()
            case "chineseCountingThousand" | "chineseCounting" | "ideographTraditional":
                return self._to_chinese(num)
            case "japaneseCounting" | "japaneseDigitalTenThousand":
                return self._to_chinese(num)
            case "bullet":
                return "•"
            case _:
                return str(num)

    @staticmethod
    def _to_roman(num: int) -> str:
        val_map = [
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        ]
        result = ""
        for value, letter in val_map:
            while num >= value:
                result += letter
                num -= value
        return result

    @staticmethod
    def _to_letter(num: int) -> str:
        result = ""
        while num > 0:
            num -= 1
            result = chr(65 + num % 26) + result
            num //= 26
        return result

    @staticmethod
    def _to_chinese(num: int) -> str:
        chinese_nums = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
        units = ["", "十", "百", "千", "万"]
        if num == 0:
            return chinese_nums[0]
        result = ""
        unit_idx = 0
        while num > 0:
            digit = num % 10
            if digit != 0:
                result = chinese_nums[digit] + units[unit_idx] + result
            elif result and result[0] != "零":
                result = chinese_nums[0] + result
            num //= 10
            unit_idx += 1
        if result.startswith("一十"):
            result = result[1:]
        return result.rstrip("零")

    def get_paragraph_number(self, p_element) -> Optional[str]:
        try:
            num_pr = p_element.find(".//w:numPr", NAMESPACES)
            if num_pr is None:
                ppr = p_element.find(".//w:pPr", NAMESPACES)
                if ppr is not None:
                    pstyle = ppr.find(".//w:pStyle", NAMESPACES)
                    if pstyle is not None:
                        style_id = pstyle.get(f"{{{W_NS}}}val")
                        if style_id and style_id in self.style_num_map:
                            num_id, ilvl = self.style_num_map[style_id]
                            return self._resolve_number(num_id, ilvl)
                return None
            num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
            ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
            if num_id_elem is None or ilvl_elem is None:
                return None
            num_id = num_id_elem.get(f"{{{W_NS}}}val")
            ilvl = ilvl_elem.get(f"{{{W_NS}}}val")
            if num_id == "0":
                return None
            return self._resolve_number(num_id, ilvl)
        except Exception as e:
            print(f"解析段落编号失败: {e}")
            return None

    def _resolve_number(self, num_id: str, ilvl: str) -> Optional[str]:
        if num_id not in self.numbering_map or ilvl not in self.numbering_map[num_id]:
            return None
        ilvl_int = int(ilvl)
        level_info = self.numbering_map[num_id][ilvl]
        counter_key = (num_id, ilvl)
        if counter_key not in self.level_counters:
            self.level_counters[counter_key] = level_info["start"]
        else:
            self.level_counters[counter_key] += 1
        for other_ilvl_str in self.numbering_map[num_id]:
            if int(other_ilvl_str) > ilvl_int:
                other_key = (num_id, other_ilvl_str)
                if other_key in self.level_counters:
                    del self.level_counters[other_key]
        text_template = level_info["text"]
        for lvl_idx in range(ilvl_int + 1):
            placeholder = f"%{lvl_idx + 1}"
            if placeholder not in text_template:
                continue
            lvl_str = str(lvl_idx)
            lvl_key = (num_id, lvl_str)
            if lvl_str in self.numbering_map[num_id] and lvl_key in self.level_counters:
                lvl_info = self.numbering_map[num_id][lvl_str]
                lvl_num = self.level_counters[lvl_key]
                formatted = self._format_number(lvl_num, lvl_info["format"])
                text_template = text_template.replace(placeholder, formatted)
        return text_template


# ============================================================
#  主函数
# ============================================================
def extract_body_text(doc_path: str) -> str:
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    headers, footers = _extract_header_footer(doc_path)
    loader = DocAnchorsLoader(doc_path)
    numbering = NumberingSystem(doc_path)
    chart_extractor = ChartExtractor(doc_path)

    doc = Document(doc_path)
    body = doc.element.body

    body_lines: List[str] = []
    inserted_chart_paths: Set[str] = set()

    def _process_paragraph(p_elem) -> List[str]:
        lines: List[str] = []
        num = numbering.get_paragraph_number(p_elem)
        text = _get_xml_text(p_elem, skip_chart_drawing=True, skip_textbox=True)
        extras = _process_anchored_content(p_elem, loader)

        if num and text:
            lines.append(f"{num} {text}")
        elif text:
            lines.append(text)

        for e in extras:
            lines.append(e)

        chart_rids = _extract_chart_rids_from_paragraph(p_elem)
        for rid in chart_rids:
            chart_path = chart_extractor.rid_to_chart_path.get(rid, "")
            if chart_path:
                inserted_chart_paths.add(chart_path)
            chart_texts = chart_extractor.get_chart_text_by_rid(rid)
            if chart_texts:
                lines.append("[图表内容开始]")
                lines.extend(chart_texts)
                lines.append("[图表内容结束]")

        return lines

    for child in body.iterchildren():
        if child.tag.endswith("p"):
            body_lines.extend(_process_paragraph(child))
        elif child.tag.endswith("tbl"):
            for row in child.iter(f"{{{W_NS}}}tr"):
                row_text: List[str] = []
                for cell in row.iter(f"{{{W_NS}}}tc"):
                    cell_content: List[str] = []
                    for p in cell.iter(f"{{{W_NS}}}p"):
                        cell_content.extend(_process_paragraph(p))
                    row_text.append("\t".join(cell_content))
                body_lines.append("\t".join(row_text))

    all_chart_texts = chart_extractor.get_all_chart_texts()
    for chart_path, texts in all_chart_texts.items():
        if chart_path not in inserted_chart_paths and texts:
            body_lines.append("[未关联图表内容开始]")
            body_lines.extend(texts)
            body_lines.append("[未关联图表内容结束]")

    result: List[str] = []
    if headers:
        result.append("[HEADER]")
        result.extend(headers)
    result.append("\n[BODY]")
    result.extend(body_lines)
    if footers:
        result.append("\n[FOOTER]")
        result.extend(footers)

    return "\n".join(result)


# ============================================================
#  测试
# ============================================================
if __name__ == "__main__":
    path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\雅本化学2025ESG报告文字稿-20260409.docx"
    try:
        result = extract_body_text(path)
        print(result)
    except FileNotFoundError as e:
        print(f"错误: {e}")
    except Exception as e:
        print(f"提取失败: {e}")