"""
PDF 批注和编辑模块
使用 PyMuPDF (fitz) 实现 PDF 的文本替换和批注功能
使用统一的文本匹配逻辑（与 Word 共用）
"""

import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Tuple, Optional, Dict
from datetime import datetime
from 数值检查1.llm.llm_project.replace.text_matcher import TextMatcher, generate_search_variants


class PDFAnnotator:
    """PDF 批注器 - 支持文本高亮和批注"""
    
    def __init__(self, pdf_path: str):
        """
        初始化 PDF 批注器
        
        Args:
            pdf_path: PDF 文件路径
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")
        
        self.doc = fitz.open(str(self.pdf_path))
        self.annotations_count = 0
        self.matcher = TextMatcher()  # 使用统一的文本匹配器
        
    def find_text_instances(self, search_text: str, case_sensitive: bool = False, use_smart_match: bool = True) -> List[Dict]:
        """
        在 PDF 中查找所有匹配的文本实例（使用统一的智能匹配逻辑）
        
        Args:
            search_text: 要查找的文本
            case_sensitive: 是否区分大小写
            use_smart_match: 是否使用智能匹配（包括变体和模糊匹配）
            
        Returns:
            匹配实例列表，每个实例包含页码和位置信息
        """
        instances = []
        flags = 0 if case_sensitive else fitz.TEXT_PRESERVE_WHITESPACE
        
        # 生成搜索变体
        search_variants = [search_text]
        if use_smart_match:
            search_variants = generate_search_variants(search_text)
        
        # 对每个变体进行搜索
        for variant in search_variants:
            for page_num in range(len(self.doc)):
                page = self.doc[page_num]
                
                # 使用 PyMuPDF 的搜索功能
                text_instances = page.search_for(variant, flags=flags)
                
                for rect in text_instances:
                    instances.append({
                        'page_num': page_num,
                        'rect': rect,
                        'text': variant,
                        'original_search': search_text
                    })
            
            # 如果找到了匹配，就不再尝试其他变体
            if instances:
                break
        
        return instances
    
    def add_highlight_annotation(
        self, 
        page_num: int, 
        rect: fitz.Rect, 
        comment: str = "",
        color: Tuple[float, float, float] = (1, 1, 0)  # 黄色
    ) -> bool:
        """
        在指定位置添加高亮批注
        
        Args:
            page_num: 页码（从0开始）
            rect: 文本矩形区域
            comment: 批注内容
            color: RGB 颜色值 (0-1范围)
            
        Returns:
            是否成功添加批注
        """
        try:
            if page_num < 0 or page_num >= len(self.doc):
                return False
            
            page = self.doc[page_num]
            
            # 添加高亮
            highlight = page.add_highlight_annot(rect)
            highlight.set_colors(stroke=color)
            
            # 添加批注内容
            if comment:
                highlight.set_info(content=comment)
                highlight.set_info(title="数值检查")
            
            highlight.update()
            self.annotations_count += 1
            
            return True
            
        except Exception as e:
            print(f"添加批注失败: {e}")
            return False
    
    def add_text_annotation(
        self,
        page_num: int,
        rect: fitz.Rect,
        comment: str,
        icon: str = "Comment"  # Comment, Note, Help, Insert, Key, NewParagraph, Paragraph
    ) -> bool:
        """
        添加文本批注（便签式）
        
        Args:
            page_num: 页码
            rect: 批注位置
            comment: 批注内容
            icon: 图标类型
            
        Returns:
            是否成功
        """
        try:
            if page_num < 0 or page_num >= len(self.doc):
                return False
            
            page = self.doc[page_num]
            
            # 创建文本批注
            annot = page.add_text_annot(rect.tl, comment, icon=icon)
            annot.set_info(title="数值检查")
            annot.update()
            
            self.annotations_count += 1
            return True
            
        except Exception as e:
            print(f"添加文本批注失败: {e}")
            return False
    
    def highlight_and_comment(
        self,
        search_text: str,
        comment: str,
        color: Tuple[float, float, float] = (1, 1, 0),
        max_instances: int = -1
    ) -> int:
        """
        查找文本并添加高亮和批注
        
        Args:
            search_text: 要查找的文本
            comment: 批注内容
            color: 高亮颜色
            max_instances: 最多处理多少个实例（-1表示全部）
            
        Returns:
            成功添加批注的数量
        """
        instances = self.find_text_instances(search_text)
        
        if max_instances > 0:
            instances = instances[:max_instances]
        
        success_count = 0
        for instance in instances:
            if self.add_highlight_annotation(
                instance['page_num'],
                instance['rect'],
                comment,
                color
            ):
                success_count += 1
        
        return success_count
    
    def _verify_highlighted_text(self, instance: Dict, search_text: str, debug: bool = False) -> bool:
        """
        验证高亮区域的文本是否与搜索文本匹配
        
        Args:
            instance: 文本实例字典，包含 page_num 和 rect
            search_text: 要搜索的文本
            debug: 是否输出调试信息
            
        Returns:
            bool: 如果匹配返回 True，否则返回 False
        """
        try:
            page = self.doc[instance['page_num']]
            rect = instance['rect']
            
            # 方法1: 使用 get_text() 提取矩形区域的文本
            highlighted_text = page.get_text("text", clip=rect).strip()
            
            # 方法2: 使用 get_textbox() 提取（可能更准确）
            textbox_text = page.get_textbox(rect).strip()
            
            # 标准化文本进行比较（移除空白字符）
            normalized_search = ''.join(search_text.split())
            normalized_highlight = ''.join(highlighted_text.split())
            normalized_textbox = ''.join(textbox_text.split())
            
            # 对于短文本（如单个数字），要求精确匹配
            if len(normalized_search) <= 3:
                # 检查提取的文本是否正好是搜索文本（不能有额外字符）
                match = (normalized_highlight == normalized_search or 
                        normalized_textbox == normalized_search)
                
                if debug:
                    print(f"    [验证] 搜索: '{search_text}'")
                    print(f"           get_text: '{highlighted_text}' (标准化: '{normalized_highlight}')")
                    print(f"           get_textbox: '{textbox_text}' (标准化: '{normalized_textbox}')")
                    print(f"           精确匹配: {match}")
                
                return match
            else:
                # 对于长文本，使用包含检查
                match = (normalized_search in normalized_highlight or 
                        normalized_highlight in normalized_search or
                        normalized_search == normalized_highlight or
                        normalized_search in normalized_textbox or
                        normalized_textbox in normalized_search)
                
                if debug:
                    print(f"    [验证] 搜索: '{search_text}', 高亮: '{highlighted_text}', 匹配: {match}")
                
                return match
            
        except Exception as e:
            if debug:
                print(f"    [验证错误] {e}")
            return False
    
    def replace_text_at_annotation(
        self,
        search_text: str,
        new_text: str,
        context: str = "",
        debug: bool = False
    ) -> tuple:
        """
        基于批注位置精确替换文本（保留格式和样式）
        
        这个方法利用批注功能找到的精确位置来替换文本，
        确保替换的位置、格式、样式都与原文一致
        
        Args:
            search_text: 要查找的文本
            new_text: 新文本
            context: 上下文（用于精确定位）
            debug: 是否输出调试信息
            
        Returns:
            (成功数量, 位置标识符) 元组
        """
        instances = self.find_text_instances(search_text)
        
        if debug:
            print(f"\n    [调试] 查找 '{search_text}': 找到 {len(instances)} 个实例")
        
        if not instances:
            if debug:
                print(f"    [调试] 未找到任何匹配")
            return (0, None)
        
        # 如果没有提供上下文，或者只有一个实例，直接处理第一个
        if not context or len(instances) == 1:
            if debug:
                print(f"    [调试] {'无上下文' if not context else '只有一个实例'}，使用第一个匹配")
            
            instance = instances[0]
            
            # 验证高亮文本是否与搜索文本一致
            if not self._verify_highlighted_text(instance, search_text, debug):
                if debug:
                    print(f"    [调试] 验证失败，跳过此实例")
                return (0, None)
            
            # 执行替换
            success = self._replace_text_preserve_format(
                instance['page_num'],
                instance['rect'],
                search_text,
                new_text,
                debug
            )
            
            if success:
                position_key = f"{instance['page_num']}_{instance['rect'].x0:.2f}_{instance['rect'].y0:.2f}"
                return (1, position_key)
            return (0, None)
        
        # 有多个实例且提供了上下文，需要智能匹配
        def normalize_text(text):
            """标准化文本：移除空白、标点，转小写"""
            import re
            text = ''.join(text.split())
            text = text.lower()
            text = re.sub(r'[^\w]', '', text)
            return text
        
        clean_context = normalize_text(context)
        clean_search = normalize_text(search_text)
        
        if debug:
            print(f"    [调试] 标准化上下文: {clean_context[:50]}...")
        
        # 检查上下文是否只包含搜索文本本身（无效上下文）
        if clean_context == clean_search:
            if debug:
                print(f"    [调试] ⚠️ 上下文无效：上下文只包含搜索文本本身")
            return (0, None)
        
        # 上下文包含其他字符，严格按照上下文匹配
        if debug:
            print(f"    [调试] 上下文有效，严格按照上下文匹配")
        
        best_match = None
        best_score = 0
        
        def longest_common_substring_length(s1, s2):
            """计算最长公共子串长度"""
            m, n = len(s1), len(s2)
            if m == 0 or n == 0:
                return 0
            
            dp = [[0] * (n + 1) for _ in range(m + 1)]
            max_len = 0
            
            for i in range(1, m + 1):
                for j in range(1, n + 1):
                    if s1[i-1] == s2[j-1]:
                        dp[i][j] = dp[i-1][j-1] + 1
                        max_len = max(max_len, dp[i][j])
            
            return max_len
        
        for idx, instance in enumerate(instances):
            page_num = instance['page_num']
            rect = instance['rect']
            page = self.doc[page_num]
            
            # 扩展矩形区域以获取上下文
            expanded_rect = fitz.Rect(
                max(0, rect.x0 - 150),
                max(0, rect.y0 - 20),
                min(page.rect.width, rect.x1 + 150),
                min(page.rect.height, rect.y1 + 20)
            )
            
            surrounding_text = page.get_text("text", clip=expanded_rect)
            clean_surrounding = normalize_text(surrounding_text)
            
            if debug:
                print(f"    [调试] 实例 {idx+1}: 页面 {page_num+1}, 位置 x={rect.x0:.2f}")
                print(f"           周围文本: {surrounding_text[:80]}...")
            
            # 计算匹配得分
            match_count = sum(1 for char in clean_context if char in clean_surrounding)
            char_score = match_count / len(clean_context) if clean_context else 0
            
            lcs_length = longest_common_substring_length(clean_context, clean_surrounding)
            lcs_score = lcs_length / len(clean_context) if clean_context else 0
            
            score = char_score * 0.4 + lcs_score * 0.6
            
            if debug:
                print(f"           得分 {score:.2%} (字符: {char_score:.2%}, 子串: {lcs_score:.2%})")
            
            if score > best_score:
                best_score = score
                best_match = instance
        
        if debug:
            print(f"    [调试] 最佳匹配得分: {best_score:.2%}")
        
        # 要求至少50%的匹配度
        if best_match and best_score >= 0.5:
            if debug:
                print(f"    [调试] 使用最佳匹配（得分 >= 50%）")
            
            # 验证高亮文本
            if not self._verify_highlighted_text(best_match, search_text, debug):
                if debug:
                    print(f"    [调试] 最佳匹配验证失败")
                return (0, None)
            
            # 执行替换
            success = self._replace_text_preserve_format(
                best_match['page_num'],
                best_match['rect'],
                search_text,
                new_text,
                debug
            )
            
            if success:
                position_key = f"{best_match['page_num']}_{best_match['rect'].x0:.2f}_{best_match['rect'].y0:.2f}"
                return (1, position_key)
        else:
            if debug:
                print(f"    [调试] ❌ 上下文匹配失败：最佳得分 {best_score:.2%} < 50%")
        
        return (0, None)
    
    def _replace_text_preserve_format(
        self,
        page_num: int,
        rect: fitz.Rect,
        old_text: str,
        new_text: str,
        debug: bool = False
    ) -> bool:
        """
        在指定位置替换文本，保留原格式和样式
        
        Args:
            page_num: 页码
            rect: 文本矩形区域
            old_text: 旧文本
            new_text: 新文本
            debug: 是否调试
            
        Returns:
            是否成功
        """
        try:
            page = self.doc[page_num]
            
            # 1. 提取原文本的格式信息
            blocks = page.get_text("dict", clip=rect)
            
            # 默认值
            font_size = rect.height * 0.7
            font_name = "helv"
            font_color = (0, 0, 0)  # 黑色
            
            # 尝试提取实际的格式信息
            if blocks and "blocks" in blocks:
                for block in blocks["blocks"]:
                    if block.get("type") == 0:  # 文本块
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                if "size" in span:
                                    font_size = span["size"]
                                if "color" in span:
                                    # 转换颜色格式
                                    color_int = span["color"]
                                    r = ((color_int >> 16) & 0xFF) / 255.0
                                    g = ((color_int >> 8) & 0xFF) / 255.0
                                    b = (color_int & 0xFF) / 255.0
                                    font_color = (r, g, b)
                                if "font" in span:
                                    original_font = span["font"]
                                    # 尝试映射到支持的字体
                                    if "bold" in original_font.lower():
                                        font_name = "helv-bold"
                                    elif "italic" in original_font.lower():
                                        font_name = "helv-oblique"
                                    else:
                                        font_name = "helv"
                                break
            
            if debug:
                print(f"    [格式] 字体: {font_name}, 大小: {font_size:.1f}, 颜色: {font_color}")
            
            # 2. 用白色矩形覆盖旧文本（稍微扩大以确保完全覆盖）
            cover_rect = fitz.Rect(
                rect.x0 - 0.5,
                rect.y0 - 0.5,
                rect.x1 + 0.5,
                rect.y1 + 0.5
            )
            page.draw_rect(cover_rect, color=(1, 1, 1), fill=(1, 1, 1))
            
            # 3. 计算精确的插入点（保持原位置）
            # 使用基线对齐
            insert_point = fitz.Point(rect.x0, rect.y1 - font_size * 0.2)
            
            # 4. 尝试插入新文本（使用支持的字体）
            supported_fonts = [font_name, "helv", "cour", "symbol"]
            inserted = False
            
            for try_font in supported_fonts:
                try:
                    rc = page.insert_text(
                        insert_point,
                        new_text,
                        fontsize=font_size,
                        color=font_color,
                        fontname=try_font
                    )
                    inserted = True
                    if debug:
                        print(f"    [成功] 使用字体 '{try_font}' 插入文本")
                    break
                except Exception as e:
                    if debug:
                        print(f"    [尝试] 字体 '{try_font}' 失败: {e}")
                    continue
            
            if not inserted:
                # 最后尝试：不指定字体名
                try:
                    rc = page.insert_text(
                        insert_point,
                        new_text,
                        fontsize=font_size,
                        color=font_color
                    )
                    inserted = True
                    if debug:
                        print(f"    [成功] 使用默认字体插入文本")
                except Exception as e:
                    if debug:
                        print(f"    [失败] 无法插入文本: {e}")
                    return False
            
            return inserted
            
        except Exception as e:
            if debug:
                print(f"    [错误] 替换失败: {e}")
            return False
    
    def highlight_and_comment_with_context(
        self,
        search_text: str,
        comment: str,
        context: str = "",
        color: Tuple[float, float, float] = (1, 1, 0),
        context_window: int = 150,
        debug: bool = False
    ) -> tuple:
        """
        基于上下文查找文本并添加高亮和批注（智能匹配）

        Args:
            search_text: 要查找的文本
            comment: 批注内容
            context: 上下文文本（用于精确定位）
            color: 高亮颜色
            context_window: 上下文窗口大小（字符数）
            debug: 是否输出调试信息

        Returns:
            (成功添加批注的数量, 位置标识符) 元组
        """
        instances = self.find_text_instances(search_text)

        if debug:
            print(f"\n    [调试] 查找 '{search_text}': 找到 {len(instances)} 个实例")

        if not instances:
            if debug:
                print(f"    [调试] 未找到任何匹配")
            return (0, None)

        # 如果没有提供上下文，或者只有一个实例，直接处理第一个
        if not context or len(instances) == 1:
            if debug:
                print(f"    [调试] {'无上下文' if not context else '只有一个实例'}，使用第一个匹配")

            instance = instances[0]

            # 验证高亮文本是否与搜索文本一致
            if not self._verify_highlighted_text(instance, search_text, debug):
                if debug:
                    print(f"    [调试] 验证失败，跳过此实例")
                return (0, None)

            position_key = f"{instance['page_num']}_{instance['rect'].x0:.2f}_{instance['rect'].y0:.2f}"

            if self.add_highlight_annotation(
                instance['page_num'],
                instance['rect'],
                comment,
                color
            ):
                return (1, position_key)
            return (0, None)

        # 有多个实例且提供了上下文，需要智能匹配
        def normalize_text(text):
            """标准化文本：移除空白、标点，转小写"""
            import re
            text = ''.join(text.split())
            text = text.lower()
            text = re.sub(r'[^\w]', '', text)
            return text

        clean_context = normalize_text(context)
        clean_search = normalize_text(search_text)

        if debug:
            print(f"    [调试] 标准化上下文: {clean_context[:50]}...")

        # 检查上下文是否只包含搜索文本本身（无效上下文）
        if clean_context == clean_search:
            if debug:
                print(f"    [调试] ⚠️ 上下文无效：上下文只包含搜索文本本身")
            return (0, None)

        # 上下文包含其他字符，严格按照上下文匹配
        if debug:
            print(f"    [调试] 上下文有效，严格按照上下文匹配")

        best_match = None
        best_score = 0

        def longest_common_substring_length(s1, s2):
            """计算最长公共子串长度"""
            m, n = len(s1), len(s2)
            if m == 0 or n == 0:
                return 0

            dp = [[0] * (n + 1) for _ in range(m + 1)]
            max_len = 0

            for i in range(1, m + 1):
                for j in range(1, n + 1):
                    if s1[i-1] == s2[j-1]:
                        dp[i][j] = dp[i-1][j-1] + 1
                        max_len = max(max_len, dp[i][j])

            return max_len

        for idx, instance in enumerate(instances):
            page_num = instance['page_num']
            rect = instance['rect']
            page = self.doc[page_num]

            # 扩展矩形区域以获取上下文
            expanded_rect = fitz.Rect(
                max(0, rect.x0 - context_window),
                max(0, rect.y0 - 20),
                min(page.rect.width, rect.x1 + context_window),
                min(page.rect.height, rect.y1 + 20)
            )

            surrounding_text = page.get_text("text", clip=expanded_rect)
            clean_surrounding = normalize_text(surrounding_text)

            if debug:
                print(f"    [调试] 实例 {idx+1}: 页面 {page_num+1}, 位置 x={rect.x0:.2f}")
                print(f"           周围文本: {surrounding_text[:80]}...")

            # 计算匹配得分
            match_count = sum(1 for char in clean_context if char in clean_surrounding)
            char_score = match_count / len(clean_context) if clean_context else 0

            lcs_length = longest_common_substring_length(clean_context, clean_surrounding)
            lcs_score = lcs_length / len(clean_context) if clean_context else 0

            score = char_score * 0.4 + lcs_score * 0.6

            if debug:
                print(f"           得分 {score:.2%} (字符: {char_score:.2%}, 子串: {lcs_score:.2%})")

            if score > best_score:
                best_score = score
                best_match = instance

        if debug:
            print(f"    [调试] 最佳匹配得分: {best_score:.2%}")

        # 要求至少50%的匹配度
        if best_match and best_score >= 0.5:
            if debug:
                print(f"    [调试] 使用最佳匹配（得分 >= 50%）")

            # 验证高亮文本
            if not self._verify_highlighted_text(best_match, search_text, debug):
                if debug:
                    print(f"    [调试] 最佳匹配验证失败")
                return (0, None)

            position_key = f"{best_match['page_num']}_{best_match['rect'].x0:.2f}_{best_match['rect'].y0:.2f}"

            if self.add_highlight_annotation(
                best_match['page_num'],
                best_match['rect'],
                comment,
                color
            ):
                return (1, position_key)
        else:
            if debug:
                print(f"    [调试] ❌ 上下文匹配失败：最佳得分 {best_score:.2%} < 50%")

        return (0, None)


    
    def save(self, output_path: Optional[str] = None) -> str:
        """
        保存修改后的 PDF
        
        注意：不再自动创建备份，备份应该在调用此方法前由外部管理
        
        Args:
            output_path: 输出路径（None则覆盖原文件）
            
        Returns:
            保存的文件路径
        """
        if output_path is None:
            output_path = self.pdf_path
        else:
            output_path = Path(output_path)
        
        # 判断是否保存到原文件
        is_same_file = Path(output_path).resolve() == Path(self.pdf_path).resolve()
        
        # 统一使用增量保存模式以保留批注
        self.doc.saveIncr()
        
        # 如果需要保存到不同文件，复制过去
        if not is_same_file:
            import shutil
            shutil.copy2(str(self.pdf_path), str(output_path))
        
        print(f"✓ 已保存 PDF: {output_path}")
        print(f"  批注数量: {self.annotations_count}")
        
        return str(output_path)
    
    def close(self):
        """关闭 PDF 文档"""
        if self.doc:
            self.doc.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class PDFTextReplacer:
    """PDF 文本替换器 - 通过重绘实现文本替换"""
    
    def __init__(self, pdf_path: str):
        """
        初始化 PDF 文本替换器
        
        Args:
            pdf_path: PDF 文件路径
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")
        
        self.doc = fitz.open(str(self.pdf_path))
        self.replacements_count = 0
    
    def replace_text(
        self,
        old_text: str,
        new_text: str,
        page_num: Optional[int] = None,
        case_sensitive: bool = False,
        debug: bool = False
    ) -> int:
        """
        替换 PDF 中的文本
        
        注意：PDF 文本替换比较复杂，这个方法会：
        1. 找到旧文本的位置
        2. 用白色矩形覆盖旧文本
        3. 在相同位置绘制新文本
        
        Args:
            old_text: 要替换的文本
            new_text: 新文本
            page_num: 指定页码（None表示所有页）
            case_sensitive: 是否区分大小写
            debug: 是否输出调试信息
            
        Returns:
            替换的数量
        """
        pages_to_process = [page_num] if page_num is not None else range(len(self.doc))
        replaced_count = 0
        
        for pnum in pages_to_process:
            if pnum < 0 or pnum >= len(self.doc):
                continue
            
            page = self.doc[pnum]
            
            # 查找文本
            flags = 0 if case_sensitive else fitz.TEXT_PRESERVE_WHITESPACE
            text_instances = page.search_for(old_text, flags=flags)
            
            if debug and text_instances:
                print(f"  页 {pnum+1}: 找到 {len(text_instances)} 个匹配")
            
            for idx, rect in enumerate(text_instances):
                try:
                    # 获取原始文本的字体信息
                    blocks = page.get_text("dict", clip=rect)
                    font_size = rect.height * 0.7  # 默认字体大小
                    
                    # 尝试从文本块中提取字体信息
                    original_font_name = None
                    if blocks and "blocks" in blocks:
                        for block in blocks["blocks"]:
                            if "lines" in block:
                                for line in block["lines"]:
                                    if "spans" in line:
                                        for span in line["spans"]:
                                            if "size" in span:
                                                font_size = span["size"]
                                            if "font" in span:
                                                original_font_name = span["font"]
                                            break
                    
                    # 1. 用白色矩形覆盖旧文本（稍微扩大一点以确保完全覆盖）
                    cover_rect = fitz.Rect(
                        rect.x0 - 1,
                        rect.y0 - 1,
                        rect.x1 + 1,
                        rect.y1 + 1
                    )
                    page.draw_rect(cover_rect, color=(1, 1, 1), fill=(1, 1, 1))
                    
                    # 2. 在相同位置插入新文本
                    # 计算文本插入点（稍微向下偏移以对齐基线）
                    insert_point = fitz.Point(rect.x0, rect.y1 - font_size * 0.2)
                    
                    # 尝试使用内置字体（不需要字体文件）
                    # PyMuPDF 内置字体：helv, tiro, cour, symb, zadb
                    builtin_fonts = ["helv", "tiro", "cour", "times", "courier", "helvetica"]
                    
                    inserted = False
                    for font_name in builtin_fonts:
                        try:
                            # 插入新文本
                            rc = page.insert_text(
                                insert_point,
                                new_text,
                                fontsize=font_size,
                                color=(0, 0, 0),  # 黑色
                                fontname=font_name
                            )
                            inserted = True
                            
                            if debug:
                                print(f"    [{idx+1}] 替换成功: '{old_text}' -> '{new_text}' (字体: {font_name}, 大小: {font_size:.1f})")
                            break
                        except:
                            continue
                    
                    if not inserted:
                        # 如果所有内置字体都失败，尝试不指定字体名
                        try:
                            rc = page.insert_text(
                                insert_point,
                                new_text,
                                fontsize=font_size,
                                color=(0, 0, 0)
                            )
                            inserted = True
                            if debug:
                                print(f"    [{idx+1}] 替换成功: '{old_text}' -> '{new_text}' (默认字体, 大小: {font_size:.1f})")
                        except Exception as e:
                            if debug:
                                print(f"    [{idx+1}] 插入文本失败: {e}")
                            raise
                    
                    if inserted:
                        replaced_count += 1
                    
                except Exception as e:
                    if debug:
                        print(f"替换文本失败 (页{pnum+1}, 实例{idx+1}): {e}")
                        import traceback
                        traceback.print_exc()
                    else:
                        print(f"替换文本失败 (页{pnum+1}, 实例{idx+1}): {e}")
        
        self.replacements_count += replaced_count
        return replaced_count
    
    def save(self, output_path: Optional[str] = None) -> str:
        """
        保存修改后的 PDF
        
        注意：不再自动创建备份，备份应该在调用此方法前由外部管理
        
        Args:
            output_path: 输出路径
            
        Returns:
            保存的文件路径
        """
        if output_path is None:
            output_path = self.pdf_path
        else:
            output_path = Path(output_path)
        
        # 判断是否保存到原文件
        is_same_file = Path(output_path).resolve() == Path(self.pdf_path).resolve()
        
        try:
            if is_same_file:
                # 保存到原文件，使用增量模式
                self.doc.saveIncr()
            else:
                # 保存到新文件，使用完整保存
                self.doc.save(str(output_path), garbage=4, deflate=True, clean=True)
            
            print(f"✓ 已保存 PDF: {output_path}")
            print(f"  替换数量: {self.replacements_count}")
            
        except Exception as e:
            print(f"✗ 保存 PDF 失败: {e}")
            raise
        
        return str(output_path)
    
    def close(self):
        """关闭 PDF 文档"""
        if self.doc:
            self.doc.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# 使用示例
if __name__ == "__main__":
    import sys

    # 检查命令行参数
    if len(sys.argv) < 2:
        print("用法: python pdf_annotator.py <PDF文件路径>")
        print("\n示例:")
        print('  python pdf_annotator.py "example.pdf"')
        sys.exit(1)

    pdf_file = sys.argv[1]

    if not Path(pdf_file).exists():
        print(f"错误：文件不存在: {pdf_file}")
        sys.exit(1)

    if not pdf_file.lower().endswith('.pdf'):
        print(f"错误：不是 PDF 文件: {pdf_file}")
        sys.exit(1)

    # 示例1：添加批注
    print("=== 示例1：添加批注 ===")
    try:
        with PDFAnnotator(pdf_file) as annotator:
            # 查找并高亮所有 "2024" 文本
            count = annotator.highlight_and_comment(
                search_text="2024",
                comment="年份需要更新为2025",
                color=(1, 1, 0)  # 黄色
            )
            print(f"添加了 {count} 个批注")
            
            # 保存
            output_file = pdf_file.replace(".pdf", "_annotated.pdf")
            annotator.save(output_file)
    except Exception as e:
        print(f"批注失败: {e}")
        import traceback
        traceback.print_exc()
    
    # 示例2：替换文本
    print("\n=== 示例2：替换文本 ===")
    try:
        with PDFTextReplacer(pdf_file) as replacer:
            # 替换所有 "2024" 为 "2025"
            count = replacer.replace_text("2024", "2025")
            print(f"替换了 {count} 处文本")
            
            # 保存
            output_file = pdf_file.replace(".pdf", "_replaced.pdf")
            replacer.save(output_file)
    except Exception as e:
        print(f"替换失败: {e}")
        import traceback
        traceback.print_exc()
