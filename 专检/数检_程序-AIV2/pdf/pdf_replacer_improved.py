"""
改进的 PDF 替换模块
利用批注定位来精确替换文本，保留格式和样式
"""
import fitz
from pathlib import Path
from typing import Tuple, Optional
from pdf.pdf_annotator import PDFAnnotator


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
        color: Tuple[float, float, float] = (0, 1, 0),
        debug: bool = False,
        prev_tgt: str = "",
        next_tgt: str = "",
    ) -> tuple:
        """
        替换文本并添加批注。多候选时先用前后句夹逼，再用上下文得分。
        """
        instances = self.find_text_instances(search_text)
        if not instances:
            if debug:
                print(f"    [未找到] '{search_text}'")
            return (0, 0, None)

        # 多候选时夹逼定位
        if len(instances) > 1 and (prev_tgt or next_tgt):
            sandwiched = self._sandwich_by_neighbors(instances, prev_tgt, next_tgt, debug)
            if sandwiched:
                instances = sandwiched

        best_instance = self._find_best_match_with_context(instances, search_text, context, debug)
        if not best_instance:
            if debug:
                print(f"    [匹配失败] 无法定位 '{search_text}'")
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
    
    def _sandwich_by_neighbors(self, instances: list, prev_tgt: str, next_tgt: str, debug: bool = False) -> list:
        """
        用前一句/后一句译文在 PDF 全文中精确定位，夹逼出正确的候选实例。
        策略：
          1. 全文提取每页文本块，精确搜索前句/后句所在位置（页码+y坐标）
          2. 找满足 prev_pos < candidate < next_pos 的候选
          3. 唯一时返回；多个时取距离最近的
        """
        import re

        def _clean(t: str) -> str:
            return re.sub(r'\s+', '', t).lower() if t else ""

        MAX_GAP_Y = 2000   # 同页最大 y 距离（pt），跨页时用页码差代替

        # 构建全文文本块索引：[(page_num, y0, clean_text), ...]
        text_blocks = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            for block in page.get_text("blocks"):
                # block: (x0, y0, x1, y1, text, block_no, block_type)
                if block[6] == 0 and block[4].strip():  # type 0 = text
                    text_blocks.append((page_num, block[1], _clean(block[4])))

        def _find_positions(target_clean: str) -> list:
            if not target_clean:
                return []
            exact = [(pn, y) for pn, y, t in text_blocks if target_clean in t]
            if exact:
                return exact
            # 降级：字符覆盖率 >= 0.8
            result = []
            for pn, y, t in text_blocks:
                if not t:
                    continue
                overlap = sum(1 for ch in target_clean if ch in t) / len(target_clean)
                if overlap >= 0.8:
                    result.append((pn, y))
            return result

        prev_positions = _find_positions(_clean(prev_tgt))
        next_positions = _find_positions(_clean(next_tgt))

        if not prev_positions and not next_positions:
            return []

        # 每个 instance 的位置：(page_num, y0)
        def _inst_pos(inst):
            return (inst["page_num"], inst["rect"].y0)

        def _pos_less(a, b):
            """位置比较：页码优先，同页按 y0"""
            return a[0] < b[0] or (a[0] == b[0] and a[1] < b[1])

        def _pos_dist(a, b):
            """位置距离：跨页用页差*10000+y差近似"""
            return abs(a[0] - b[0]) * 10000 + abs(a[1] - b[1])

        def _is_sandwiched(inst) -> bool:
            pos = _inst_pos(inst)
            if prev_positions and not next_positions:
                return any(_pos_less(p, pos) and _pos_dist(p, pos) <= MAX_GAP_Y + 10000
                           for p in prev_positions)
            if next_positions and not prev_positions:
                return any(_pos_less(pos, n) and _pos_dist(pos, n) <= MAX_GAP_Y + 10000
                           for n in next_positions)
            closest_prev = min(
                (p for p in prev_positions if _pos_less(p, pos)),
                key=lambda p: _pos_dist(p, pos), default=None
            )
            closest_next = min(
                (n for n in next_positions if _pos_less(pos, n)),
                key=lambda n: _pos_dist(pos, n), default=None
            )
            if closest_prev is None or closest_next is None:
                return False
            return (_pos_dist(closest_prev, pos) <= MAX_GAP_Y + 10000 and
                    _pos_dist(pos, closest_next) <= MAX_GAP_Y + 10000)

        sandwiched = [inst for inst in instances if _is_sandwiched(inst)]
        if not sandwiched:
            if debug:
                print(f"    [PDF夹逼] 无夹逼命中，退化上下文匹配")
            return []

        if len(sandwiched) == 1:
            return sandwiched

        # 多个：取与前后句距离之和最小的
        def _gap(inst) -> float:
            pos = _inst_pos(inst)
            cp = min((p for p in prev_positions if _pos_less(p, pos)),
                     key=lambda p: _pos_dist(p, pos), default=pos)
            cn = min((n for n in next_positions if _pos_less(pos, n)),
                     key=lambda n: _pos_dist(pos, n), default=pos)
            return _pos_dist(cp, pos) + _pos_dist(pos, cn)

        sandwiched.sort(key=_gap)
        if len(sandwiched) > 1 and _gap(sandwiched[1]) == _gap(sandwiched[0]):
            if debug:
                print(f"    [PDF夹逼] 多候选距离相同，退化上下文匹配")
            return []
        return [sandwiched[0]]

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
