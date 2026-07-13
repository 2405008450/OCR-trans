#!/usr/bin/env python3
"""
LaTeX 公式 → 人类可读纯文本（Unicode）转换器

从包含 LaTeX 公式的文本（或 HTML 源码）中提取公式，
并将其还原为可读的 Unicode 数学表达式。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# ============================================================
# 1. Unicode 映射表
# ============================================================

# 上标字符映射
SUPERSCRIPT_MAP: dict[str, str] = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
    'n': 'ⁿ', 'i': 'ⁱ', 'x': 'ˣ', 'y': 'ʸ',
    'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
    'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'j': 'ʲ', 'k': 'ᵏ',
    'l': 'ˡ', 'm': 'ᵐ', 'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ',
    's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ', 'v': 'ᵛ', 'w': 'ʷ',
    'z': 'ᶻ',
}

# 下标字符映射
SUBSCRIPT_MAP: dict[str, str] = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
    'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ',
    'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
    'p': 'ₚ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ',
    'v': 'ᵥ', 'x': 'ₓ',
}

# 希腊字母映射
GREEK_MAP: dict[str, str] = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
    'epsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ',
    'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
    'nu': 'ν', 'xi': 'ξ', 'pi': 'π', 'rho': 'ρ',
    'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ',
    'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    # 大写
    'Alpha': 'Α', 'Beta': 'Β', 'Gamma': 'Γ', 'Delta': 'Δ',
    'Epsilon': 'Ε', 'Zeta': 'Ζ', 'Eta': 'Η', 'Theta': 'Θ',
    'Iota': 'Ι', 'Kappa': 'Κ', 'Lambda': 'Λ', 'Mu': 'Μ',
    'Nu': 'Ν', 'Xi': 'Ξ', 'Pi': 'Π', 'Rho': 'Ρ',
    'Sigma': 'Σ', 'Tau': 'Τ', 'Upsilon': 'Υ', 'Phi': 'Φ',
    'Chi': 'Χ', 'Psi': 'Ψ', 'Omega': 'Ω',
    # 变体
    'varepsilon': 'ε', 'varphi': 'φ', 'varpi': 'ϖ',
    'varrho': 'ϱ', 'varsigma': 'ς', 'vartheta': 'ϑ',
}

# 特殊符号映射
SYMBOL_MAP: dict[str, str] = {
    'pm': '±', 'mp': '∓', 'times': '×', 'div': '÷',
    'cdot': '·', 'cdots': '⋯', 'ldots': '…', 'vdots': '⋮',
    'ddots': '⋱', 'ast': '∗',
    'leq': '≤', 'geq': '≥', 'neq': '≠', 'approx': '≈',
    'equiv': '≡', 'sim': '∼', 'propto': '∝',
    'le': '≤', 'ge': '≥', 'ne': '≠', 'll': '≪', 'gg': '≫',
    'infty': '∞', 'partial': '∂', 'nabla': '∇',
    'forall': '∀', 'exists': '∃', 'nexists': '∄',
    'in': '∈', 'notin': '∉', 'subset': '⊂', 'supset': '⊃',
    'subseteq': '⊆', 'supseteq': '⊇',
    'cup': '∪', 'cap': '∩', 'emptyset': '∅',
    'land': '∧', 'lor': '∨', 'lnot': '¬', 'neg': '¬',
    'Rightarrow': '⇒', 'Leftarrow': '⇐', 'Leftrightarrow': '⇔',
    'rightarrow': '→', 'leftarrow': '←', 'leftrightarrow': '↔',
    'to': '→', 'mapsto': '↦',
    'uparrow': '↑', 'downarrow': '↓',
    'angle': '∠', 'triangle': '△', 'perp': '⊥', 'parallel': '∥',
    'star': '⋆', 'circ': '∘', 'bullet': '•',
    'hbar': 'ℏ', 'ell': 'ℓ', 'Re': 'ℜ', 'Im': 'ℑ',
    'aleph': 'ℵ',
    'quad': '  ', 'qquad': '    ',
    ',': ' ', ';': '  ', '!': '', ':': ' ',  # 间距命令
    'left': '', 'right': '',  # 定界符修饰
    'Big': '', 'big': '', 'bigg': '', 'Bigg': '',
    'Bigl': '', 'Bigr': '', 'bigl': '', 'bigr': '',
    'biggl': '', 'biggr': '', 'Biggl': '', 'Biggr': '',
    'displaystyle': '', 'textstyle': '', 'scriptstyle': '',
    'mathrm': '', 'mathbf': '', 'mathit': '', 'mathsf': '',
    'mathbb': '', 'mathcal': '', 'text': '', 'textbf': '',
}

# 函数名（保持原样输出，不带反斜杠）
FUNCTION_NAMES: set[str] = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth',
    'ln', 'log', 'exp', 'lim', 'max', 'min',
    'sup', 'inf', 'det', 'dim', 'ker', 'deg',
    'arg', 'gcd', 'lcm', 'mod', 'Pr',
}


# ============================================================
# 2. LaTeX 解析器核心
# ============================================================

@dataclass
class LaTeXToUnicode:
    """将 LaTeX 数学公式转换为人类可读的 Unicode 纯文本。"""

    # 内部状态：记录递归深度，防止无限递归
    _max_depth: int = field(default=50, repr=False)

    def convert(self, latex: str) -> str:
        """主入口：将 LaTeX 字符串转换为 Unicode 文本。"""
        # 预处理：去除首尾空白、处理换行
        text = latex.strip()
        text = self._preprocess(text)
        # 递归转换
        result = self._parse(text, depth=0)
        # 后处理：清理多余空格
        result = re.sub(r' {2,}', ' ', result).strip()
        return result

    def _preprocess(self, text: str) -> str:
        """预处理 LaTeX 源码。"""
        # 移除 \displaystyle 等纯格式命令
        text = re.sub(r'\\(displaystyle|textstyle|scriptstyle)\b\s*', '', text)
        # 移除 \left \right（但保留后面的括号）
        text = re.sub(r'\\(left|right)\s*([|.\[\](){}\\])', r'\2', text)
        text = re.sub(r'\\(left|right)\s*\\([|{}])', r'\2', text)
        # 处理 \\ 换行为换行符
        text = text.replace('\\\\', '\n')
        # 处理 & 对齐符
        text = text.replace('&', ' ')
        return text

    def _parse(self, text: str, depth: int = 0) -> str:
        """核心解析：逐字符扫描并转换。"""
        if depth > self._max_depth:
            return text  # 防止无限递归

        result: list[str] = []
        i = 0
        n = len(text)

        while i < n:
            ch = text[i]

            # --- 反斜杠命令 ---
            if ch == '\\':
                cmd, consumed = self._parse_command(text, i, depth)
                result.append(cmd)
                i += consumed
                continue

            # --- 上标 ---
            if ch == '^':
                content, consumed = self._extract_group(text, i + 1)
                sup_text = self._parse(content, depth + 1)
                result.append(self._to_superscript(sup_text))
                i += 1 + consumed
                continue

            # --- 下标 ---
            if ch == '_':
                content, consumed = self._extract_group(text, i + 1)
                sub_text = self._parse(content, depth + 1)
                result.append(self._to_subscript(sub_text))
                i += 1 + consumed
                continue

            # --- 花括号组（直接展开） ---
            if ch == '{':
                content, consumed = self._extract_brace_content(text, i)
                result.append(self._parse(content, depth + 1))
                i += consumed
                continue

            # --- 忽略右花括号（不应单独出现） ---
            if ch == '}':
                i += 1
                continue

            # --- 波浪号（空格） ---
            if ch == '~':
                result.append(' ')
                i += 1
                continue

            # --- 普通字符 ---
            result.append(ch)
            i += 1

        return ''.join(result)

    def _parse_command(self, text: str, pos: int, depth: int) -> tuple[str, int]:
        """
        解析从 pos 位置开始的 \\command。
        返回 (转换后的文本, 消耗的字符数)。
        """
        n = len(text)
        # pos 指向 '\\'
        if pos + 1 >= n:
            return ('\\', 1)

        next_ch = text[pos + 1]

        # 单字符命令：\\ \{ \} \| \  等
        if not next_ch.isalpha():
            if next_ch in SYMBOL_MAP:
                return (SYMBOL_MAP[next_ch], 2)
            if next_ch == '|':
                return ('‖', 2)
            if next_ch in '{}':
                return (next_ch, 2)
            if next_ch == '\\':
                return ('\n', 2)
            if next_ch == ' ':
                return (' ', 2)
            return (next_ch, 2)

        # 提取命令名
        cmd_start = pos + 1
        j = cmd_start
        while j < n and text[j].isalpha():
            j += 1
        cmd_name = text[cmd_start:j]
        consumed = j - pos  # 从 \\ 到命令名结束

        # 跳过命令名后的空格
        while j < n and text[j] == ' ':
            j += 1

        # ---- 特殊命令处理 ----

        # \frac{a}{b} → (a) / (b)
        if cmd_name == 'frac':
            num_content, num_consumed = self._extract_group(text, j)
            j += num_consumed
            # 跳过空格
            while j < n and text[j] == ' ':
                j += 1
            den_content, den_consumed = self._extract_group(text, j)
            j += den_consumed
            num_text = self._parse(num_content, depth + 1)
            den_text = self._parse(den_content, depth + 1)
            # 简单内容不加括号
            num_str = num_text if self._is_simple(num_text) else f"({num_text})"
            den_str = den_text if self._is_simple(den_text) else f"({den_text})"
            return (f"{num_str}/{den_str}", j - pos)

        # \dfrac 同 \frac
        if cmd_name == 'dfrac':
            num_content, num_consumed = self._extract_group(text, j)
            j += num_consumed
            while j < n and text[j] == ' ':
                j += 1
            den_content, den_consumed = self._extract_group(text, j)
            j += den_consumed
            num_text = self._parse(num_content, depth + 1)
            den_text = self._parse(den_content, depth + 1)
            num_str = num_text if self._is_simple(num_text) else f"({num_text})"
            den_str = den_text if self._is_simple(den_text) else f"({den_text})"
            return (f"{num_str}/{den_str}", j - pos)

        # \sqrt[n]{x} 或 \sqrt{x}
        if cmd_name == 'sqrt':
            # 检查可选参数 [n]
            index_text = ''
            if j < n and text[j] == '[':
                end_bracket = text.find(']', j)
                if end_bracket != -1:
                    index_text = text[j + 1:end_bracket]
                    j = end_bracket + 1
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            if index_text:
                idx = self._parse(index_text, depth + 1)
                return (f"{idx}√({inner})", j - pos)
            return (f"√({inner})", j - pos)

        # \sum, \prod, \int 等大型运算符
        if cmd_name in ('sum', 'prod', 'int', 'iint', 'iiint', 'oint',
                         'bigcup', 'bigcap', 'bigoplus', 'bigotimes'):
            op_map = {
                'sum': '∑', 'prod': '∏',
                'int': '∫', 'iint': '∬', 'iiint': '∭', 'oint': '∮',
                'bigcup': '⋃', 'bigcap': '⋂',
                'bigoplus': '⊕', 'bigotimes': '⊗',
            }
            return (op_map[cmd_name], consumed)

        # \lim 特殊处理（可能跟 _{x \to 0}）
        if cmd_name == 'lim':
            return ('lim', consumed)

        # \vec{x} → x⃗
        if cmd_name == 'vec':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}\u20D7", j - pos)

        # \hat{x} → x̂
        if cmd_name == 'hat':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}\u0302", j - pos)

        # \bar{x} → x̄
        if cmd_name == 'bar':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}̄", j - pos)

        # \dot{x} → ẋ
        if cmd_name == 'dot':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}\u0307", j - pos)

        # \ddot{x} → ẍ
        if cmd_name == 'ddot':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}̈", j - pos)

        # \tilde{x} → x̃
        if cmd_name == 'tilde':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}\u0303", j - pos)

        # \overline{x} → x̅  (用上划线组合字符)
        if cmd_name == 'overline':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}\u0305", j - pos)

        # \underline{x}
        if cmd_name == 'underline':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (f"{inner}̲", j - pos)

        # \underbrace / \overbrace → 直接输出内容
        if cmd_name in ('underbrace', 'overbrace'):
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            inner = self._parse(content, depth + 1)
            return (inner, j - pos)

        # \mathrm{}, \mathbf{}, \text{} 等 → 直接输出内容
        if cmd_name in ('mathrm', 'mathbf', 'mathit', 'mathsf', 'mathbb',
                         'mathcal', 'text', 'textbf', 'textrm', 'operatorname',
                         'boldsymbol', 'bm'):
            if j < n and text[j] == '{':
                content, grp_consumed = self._extract_group(text, j)
                j += grp_consumed
                inner = self._parse(content, depth + 1)
                return (inner, j - pos)
            return ('', consumed)

        # \begin{...} ... \end{...} → 提取内部内容
        if cmd_name == 'begin':
            content, grp_consumed = self._extract_group(text, j)
            j += grp_consumed
            env_name = content
            # 找到对应的 \end{env_name}
            end_tag = f'\\end{{{env_name}}}'
            end_pos = text.find(end_tag, j)
            if end_pos != -1:
                env_body = text[j:end_pos]
                j = end_pos + len(end_tag)
                inner = self._parse(env_body, depth + 1)
                return (inner, j - pos)
            # 找不到 \end，返回剩余部分
            return ('', consumed)

        if cmd_name == 'end':
            # 跳过 \end{...}
            if j < n and text[j] == '{':
                _, grp_consumed = self._extract_group(text, j)
                j += grp_consumed
            return ('', j - pos)

        # --- 希腊字母 ---
        if cmd_name in GREEK_MAP:
            return (GREEK_MAP[cmd_name], consumed)

        # --- 函数名 ---
        if cmd_name in FUNCTION_NAMES:
            return (cmd_name, consumed)

        # --- 通用符号 ---
        if cmd_name in SYMBOL_MAP:
            return (SYMBOL_MAP[cmd_name], consumed)

        # --- 未知命令：输出命令名 ---
        return (cmd_name, consumed)

    def _extract_group(self, text: str, pos: int) -> tuple[str, int]:
        """
        从 pos 开始提取一个"组"：
        - 如果是 {content}，提取花括号内的内容
        - 否则提取单个字符
        返回 (内容, 消耗的字符数)
        """
        n = len(text)
        if pos >= n:
            return ('', 0)

        if text[pos] == '{':
            return self._extract_brace_content(text, pos)

        # 单个字符作为组
        return (text[pos], 1)

    def _extract_brace_content(self, text: str, pos: int) -> tuple[str, int]:
        """
        从 pos 位置的 '{' 开始，提取匹配花括号内的内容。
        返回 (内容字符串, 消耗的总字符数包括花括号)。
        """
        n = len(text)
        if pos >= n or text[pos] != '{':
            return ('', 0)

        depth = 1
        j = pos + 1
        while j < n and depth > 0:
            if text[j] == '{' and (j == 0 or text[j - 1] != '\\'):
                depth += 1
            elif text[j] == '}' and (j == 0 or text[j - 1] != '\\'):
                depth -= 1
            j += 1

        # j 现在指向 '}' 之后
        content = text[pos + 1:j - 1] if depth == 0 else text[pos + 1:]
        consumed = j - pos if depth == 0 else n - pos
        return (content, consumed)

    def _to_superscript(self, text: str) -> str:
        """将文本转换为 Unicode 上标字符。"""
        result: list[str] = []
        for ch in text:
            if ch in SUPERSCRIPT_MAP:
                result.append(SUPERSCRIPT_MAP[ch])
            else:
                # 无法映射的字符用 ^(...) 表示
                result.append(f'^({text})')
                return ''.join(result[:0]) + f'^({text})'
        return ''.join(result)

    def _to_subscript(self, text: str) -> str:
        """将文本转换为 Unicode 下标字符。"""
        result: list[str] = []
        for ch in text:
            if ch in SUBSCRIPT_MAP:
                result.append(SUBSCRIPT_MAP[ch])
            else:
                result.append(f'_({text})')
                return f'_({text})'
        return ''.join(result)

    @staticmethod
    def _is_simple(text: str) -> bool:
        """判断文本是否"简单"（不需要额外加括号）。"""
        # 单个字符或纯数字/字母视为简单
        if len(text) <= 1:
            return True
        if re.match(r'^[a-zA-Z0-9]+$', text):
            return True
        if re.match(r'^-?[0-9]+\.?[0-9]*$', text):
            return True
        return False


# ============================================================
# 3. HTML 文本提取器
# ============================================================

@dataclass
class HTMLFormulaExtractor:
    """从 HTML 源码中提取 LaTeX 公式和普通文本，还原为可读文本。"""

    converter: LaTeXToUnicode = field(default_factory=LaTeXToUnicode)

    def extract_and_convert(self, html_source: str) -> str:
        """
        从 HTML 源码中提取所有内容：
        1. 识别 $$...$$ (块级公式) 和 $...$ (行内公式)
        2. 将 LaTeX 转为 Unicode
        3. 将 HTML 标签剥离，保留纯文本
        """
        # 第一步：去除 HTML 标签，但保留 $...$ 公式
        text = self._strip_html_tags(html_source)
        # 第二步：处理 HTML 实体
        text = self._decode_html_entities(text)
        # 第三步：替换公式
        text = self._replace_formulas(text)
        # 第四步：清理空行
        lines = text.splitlines()
        cleaned: list[str] = []
        for line in lines:
            stripped = line.strip()
            if stripped:
                cleaned.append(stripped)
        return '\n'.join(cleaned)

    def _strip_html_tags(self, html: str) -> str:
        """去除 HTML 标签，保留文本内容。"""
        # 移除 <script>...</script> 和 <style>...</style> 块
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        # 移除 HTML 注释
        text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
        # <br> → 换行
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        # <p>, <div>, <h1>-<h6>, <li> 等块级元素 → 换行
        text = re.sub(r'</(p|div|h[1-6]|li|tr|dt|dd)>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<(p|div|h[1-6]|li|tr|dt|dd)[^>]*>', '\n', text, flags=re.IGNORECASE)
        # 移除所有剩余标签
        text = re.sub(r'<[^>]+>', '', text)
        return text

    @staticmethod
    def _decode_html_entities(text: str) -> str:
        """解码常见 HTML 实体。"""
        import html
        return html.unescape(text)

    def _replace_formulas(self, text: str) -> str:
        """将 $$...$$ 和 $...$ 中的 LaTeX 替换为 Unicode 文本。"""
        # 先处理块级公式 $$...$$
        def replace_display(match: re.Match) -> str:
            latex = match.group(1).strip()
            try:
                converted = self.converter.convert(latex)
                return f"\n  {converted}\n"
            except Exception:
                return f"\n  [{latex}]\n"

        text = re.sub(r'\$\$(.*?)\$\$', replace_display, text, flags=re.DOTALL)

        # 再处理行内公式 $...$
        def replace_inline(match: re.Match) -> str:
            latex = match.group(1).strip()
            try:
                return self.converter.convert(latex)
            except Exception:
                return f"[{latex}]"

        text = re.sub(r'\$(.+?)\$', replace_inline, text)

        return text


# ============================================================
# 4. 主程序
# ============================================================

def main() -> None:
    """演示：解析 HTML 中的 LaTeX 公式并还原为纯文本。"""

#     # 原始 HTML 源码
#     html_source = r'''<!DOCTYPE html>
# <html lang="zh-CN">
# <head>
#     <meta charset="UTF-8">
#     <title>LaTeX 公式渲染 - KaTeX</title>
#     <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.css">
#     <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.js"></script>
#     <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/contrib/auto-render.min.js"
#             onload="renderMathInElement(document.body, {
#                 delimiters: [
#                     {left: '$$', right: '$$', display: true},
#                     {left: '$', right: '$', display: false}
#                 ]
#             });"></script>
#     <style>
#         body { font-family: 'Segoe UI', Roboto, sans-serif; padding: 20px; }
#     </style>
# </head>
# <body>
#     <h2>📐 渲染效果展示</h2>
#     <div id="formula">
#         $$\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$
#         $$w=∆k=\frac{1}{2}mv^{2}−\frac{1}{2}mv_{0}^{2}$$
#     </div>
#
#     <p>这是行内公式示例：$\sin(\theta) = \frac{opp}{hyp}$，它会和文字融合在一起。</p>
#     <p>另一个行内公式：勾股定理 $a^2 + b^2 = c^2$ 非常直观。</p>
#
#     <div class="note">
#         ✅ 注意：上面的 <code>$\sin(\theta) = \frac{opp}{hyp}$</code> 已经由 KaTeX 自动渲染为数学公式。<br>
#         🔧 原代码中指数写成了 b³，已改为正确的 b²。
#     </div>
# </body>
# </html>'''

    print("=" * 60)
    print("  LaTeX → Unicode 纯文本还原器")
    print("=" * 60)

    # --- 方式一：直接转换单个公式 ---
    converter = LaTeXToUnicode()

    test_formulas = [
        r"V_{0}^{2}",
        r"\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}",
        r"w=\Delta k=\frac{1}{2}mv^{2}-\frac{1}{2}mv_{0}^{2}",
        r"\sin(\theta) = \frac{opp}{hyp}",
        r"a^2 + b^2 = c^2",
        r"E = mc^2",
        r"\int_{0}^{\infty} e^{-x^2} dx = \frac{\sqrt{\pi}}{2}",
        r"\sum_{i=1}^{n} i = \frac{n(n+1)}{2}",
        r"\lim_{x \to 0} \frac{\sin x}{x} = 1",
    ]

    print("\n📝 单公式转换演示：")
    print("-" * 60)
    for formula in test_formulas:
        result = converter.convert(formula)
        print(f"  LaTeX:   {formula}")
        print(f"  文本:    {result}")
        print()

    # --- 方式二：从 HTML 中提取并转换 ---
    print("=" * 60)
    print("📄 HTML 全文还原结果：")
    print("=" * 60)

    extractor = HTMLFormulaExtractor()
    # plain_text = extractor.extract_and_convert(html_source)
    # print(plain_text)


if __name__ == "__main__":
    main()