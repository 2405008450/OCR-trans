"""
改进的 PDF 替换模块
利用批注定位来精确替换文本，保留格式和样式
"""
import fitz
from pathlib import Path
from typing import Tuple, Optional
from 数值检查1.llm.llm_project.replace.pdf_annotator import PDFAnnotator


class ImprovedPDFReplacer(PDFAnnotator):
    """
    改进的 PDF 替换器
    继承 PDFAnnotator 的定位能力，添加精确替换功能
    """
    
    def __init__(self, pdf_path: str):
        """初始化"""
        super().__init__(pdf_path)
        self.replacements_count = 0
    
    def replace_and_annotate(
        self,
        search_text: str,
        new_text: str,
        comment: str,
        context: str = "",
        color: Tuple[float, float, float] = (0, 1, 0),  # 绿色表示已修改
        debug: bool = False
    ) -> tuple:
        """
        替换文本并添加批注（一步完成，保留格式）
        
        核心思路：
        1. 使用 highlight_and_comment_with_context 找到精确位置
        2. 在该位置替换文本（保留格式）
        3. 在新文本位置添加批注
        
        Args:
            search_text: 要查找的文本
            new_text: 新文本
            comment: 批注内容
            context: 上下文
            color: 批注颜色（绿色表示已修改）
            debug: 是否调试
            
        Returns:
            (替换成功数, 批注成功数, 位置标识符)
        """
        # 第一步：找到文本位置（使用父类的智能定位）
        instances = self.find_text_instances(search_text)
        
        if not instances:
            if debug:
                print(f"    [未找到] '{search_text}'")
            return (0, 0, None)
        
        # 使用上下文匹配找到最佳实例
        best_instance = self._find_best_match_with_context(
            instances, search_text, context, debug
        )
        
        if not best_instance:
            if debug:
                print(f"    [匹配失败] 无法通过上下文定位 '{search_text}'")
            return (0, 0, None)
        
        # 第二步：在精确位置替换文本（保留格式）
        replace_success = self._replace_at_position(
            best_instance['page_num'],
            best_instance['rect'],
            search_text,
            new_text,
            debug
        )
        
        if not replace_success:
            if debug:
                print(f"    [替换失败] '{search_text}' → '{new_text}'")
            return (0, 0, None)
        
        self.replacements_count += 1

        # 第三步：在原位置添加批注（不重新查找新文本）
        annot_success = 0

        if self.add_highlight_annotation(
                best_instance['page_num'],
                best_instance['rect'],  # 使用原矩形区域
                comment,
                color
        ):
            annot_success = 1

        position_key = f"{best_instance['page_num']}_{best_instance['rect'].x0:.2f}_{best_instance['rect'].y0:.2f}"

        return (1, annot_success, position_key)
    
    def _find_best_match_with_context(self, instances, search_text, context, debug):
        """使用上下文找到最佳匹配实例"""
        if not context or len(instances) == 1:
            # 验证文本
            if self._verify_highlighted_text(instances[0], search_text, debug):
                return instances[0]
            return None
        
        # 标准化上下文
        def normalize_text(text):
            import re
            text = ''.join(text.split())
            text = text.lower()
            text = re.sub(r'[^\w]', '', text)
            return text
        
        clean_context = normalize_text(context)
        clean_search = normalize_text(search_text)
        
        if clean_context == clean_search:
            return instances[0] if self._verify_highlighted_text(instances[0], search_text, debug) else None
        
        # 计算最佳匹配
        best_match = None
        best_score = 0
        
        for instance in instances:
            page = self.doc[instance['page_num']]
            rect = instance['rect']
            
            # 获取周围文本
            expanded_rect = fitz.Rect(
                max(0, rect.x0 - 150),
                max(0, rect.y0 - 20),
                min(page.rect.width, rect.x1 + 150),
                min(page.rect.height, rect.y1 + 20)
            )
            
            surrounding_text = page.get_text("text", clip=expanded_rect)
            clean_surrounding = normalize_text(surrounding_text)
            
            # 计算相似度
            match_count = sum(1 for char in clean_context if char in clean_surrounding)
            score = match_count / len(clean_context) if clean_context else 0
            
            if score > best_score:
                best_score = score
                best_match = instance
        
        if best_match and best_score >= 0.5:
            if self._verify_highlighted_text(best_match, search_text, debug):
                return best_match
        
        return None
    
    def _replace_at_position(
        self,
        page_num: int,
        rect: fitz.Rect,
        old_text: str,
        new_text: str,
        debug: bool = False
    ) -> bool:
        """
        在指定位置替换文本，保留原格式和样式
        
        关键：提取原文本的所有格式信息，然后精确复制
        """
        try:
            page = self.doc[page_num]
            
            # 1. 提取原文本的完整格式信息
            blocks = page.get_text("dict", clip=rect)
            
            # 默认值
            font_size = rect.height * 0.7
            font_name = "helv"
            font_color = (0, 0, 0)
            
            # 提取实际格式
            if blocks and "blocks" in blocks:
                for block in blocks["blocks"]:
                    if block.get("type") == 0:  # 文本块
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                if "size" in span:
                                    font_size = span["size"]
                                if "color" in span:
                                    color_int = span["color"]
                                    r = ((color_int >> 16) & 0xFF) / 255.0
                                    g = ((color_int >> 8) & 0xFF) / 255.0
                                    b = (color_int & 0xFF) / 255.0
                                    font_color = (r, g, b)
                                break
            
            if debug:
                print(f"    [格式] 大小: {font_size:.1f}, 颜色: {font_color}")
            
            # 2. 用白色矩形覆盖旧文本
            cover_rect = fitz.Rect(
                rect.x0 - 0.5,
                rect.y0 - 0.5,
                rect.x1 + 0.5,
                rect.y1 + 0.5
            )
            page.draw_rect(cover_rect, color=(1, 1, 1), fill=(1, 1, 1))
            
            # 3. 计算插入点（保持原位置）
            insert_point = fitz.Point(rect.x0, rect.y1 - font_size * 0.2)
            
            # 4. 插入新文本（尝试多种字体）
            supported_fonts = ["helv", "cour", "symbol"]
            
            for try_font in supported_fonts:
                try:
                    page.insert_text(
                        insert_point,
                        new_text,
                        fontsize=font_size,
                        color=font_color,
                        fontname=try_font
                    )
                    if debug:
                        print(f"    [成功] 使用字体 '{try_font}'")
                    return True
                except:
                    continue
            
            # 最后尝试：不指定字体
            try:
                page.insert_text(
                    insert_point,
                    new_text,
                    fontsize=font_size,
                    color=font_color
                )
                if debug:
                    print(f"    [成功] 使用默认字体")
                return True
            except Exception as e:
                if debug:
                    print(f"    [失败] {e}")
                return False
            
        except Exception as e:
            if debug:
                print(f"    [错误] {e}")
            return False
    
    def save(self, output_path: Optional[str] = None) -> str:
        """保存 PDF"""
        if output_path is None:
            output_path = self.pdf_path
        else:
            output_path = Path(output_path)
        
        is_same_file = Path(output_path).resolve() == Path(self.pdf_path).resolve()
        
        try:
            if is_same_file:
                self.doc.saveIncr()
            else:
                self.doc.save(str(output_path), garbage=4, deflate=True, clean=True)
            
            print(f"✓ 已保存 PDF: {output_path}")
            print(f"  替换数量: {self.replacements_count}")
            print(f"  批注数量: {self.annotations_count}")
            
        except Exception as e:
            print(f"✗ 保存失败: {e}")
            raise
        
        return str(output_path)
