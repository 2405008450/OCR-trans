"""
PPTX 替换与批注模块
使用 python-pptx 在幻灯片中查找文本并替换。
通过操作底层 OOXML 添加真正的 PowerPoint 批注（评论气泡）。
支持：上下文验证、文本变体容错、Unicode/不可见字符清洗。
"""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.oxml.ns import qn, nsmap
from lxml import etree
from pathlib import Path
from typing import Optional, List, Tuple, Dict
from datetime import datetime
from llm.llm_project.replace.text_matcher import TextMatcher, clean_text_thoroughly, generate_search_variants


# from replace.text_matcher import TextMatcher, clean_text_thoroughly, generate_search_variants


class PPTXReplacer:
    """PPTX 替换器 — 遍历所有幻灯片的形状，执行文本替换并添加批注"""

    def __init__(self, pptx_path: str):
        self.pptx_path = Path(pptx_path)
        if not self.pptx_path.exists():
            raise FileNotFoundError(f"文件不存在: {pptx_path}")
        self.prs = Presentation(str(self.pptx_path))
        self.replacements_count = 0
        self.annotations_count = 0
        self._author_id = self._ensure_author("标点检查")
        self._matcher = TextMatcher()

    # ── 公开接口 ──

    def replace_and_annotate(
        self,
        old_text: str,
        new_text: str,
        reason: str = "",
        context: str = "",
    ) -> bool:
        """
        在所有幻灯片中查找 old_text 并替换为 new_text，
        同时在该幻灯片添加 PowerPoint 批注。
        支持上下文验证和文本变体容错。
        """
        replaced = False

        # 收集所有候选匹配位置
        candidates = self._find_all_candidates(old_text)

        if not candidates:
            return False

        # 如果有上下文，用上下文筛选最佳匹配
        if context and len(candidates) > 1:
            candidates = self._rank_by_context(candidates, context)

        # 执行替换（只替换最佳匹配）
        best = candidates[0]
        slide = best["slide"]
        text_frame = best["text_frame"]
        para_idx = best["para_idx"]

        para = text_frame.paragraphs[para_idx]
        runs = para.runs
        if not runs:
            return False

        full = "".join(r.text for r in runs)
        actual_old = best.get("matched_text", old_text)

        new_full = full.replace(actual_old, new_text, 1)
        if new_full == full:
            # 如果精确替换失败，用清洗后的方式定位
            ok = self._replace_fuzzy_in_runs(runs, old_text, new_text)
            if not ok:
                return False
        else:
            self._redistribute_text(runs, new_full)

        self._add_comment(slide, old_text, new_text, reason)
        self.replacements_count += 1
        return True

    def save(self, output_path: Optional[str] = None) -> str:
        out = Path(output_path) if output_path else self.pptx_path
        self.prs.save(str(out))
        print(f"✓ 已保存 PPTX: {out}")
        print(f"  替换数量: {self.replacements_count}")
        print(f"  批注数量: {self.annotations_count}")
        return str(out)

    # ── 智能查找逻辑 ──

    def _find_all_candidates(self, old_text: str) -> List[Dict]:
        """
        在所有幻灯片中查找 old_text 的所有候选位置。
        使用多层匹配策略：精确 → 清洗后 → 变体 → 模糊。
        """
        candidates = []
        variants = generate_search_variants(old_text)
        old_clean = clean_text_thoroughly(old_text)

        for slide_idx, slide in enumerate(self.prs.slides):
            # 文本框形状
            for tf_info in self._iter_text_frames(slide.shapes, slide_idx):
                text_frame = tf_info["text_frame"]
                for para_idx, para in enumerate(text_frame.paragraphs):
                    runs = para.runs
                    if not runs:
                        continue
                    full = "".join(r.text for r in runs)
                    if not full.strip():
                        continue

                    match_result = self._try_match(full, old_text, old_clean, variants)
                    if match_result:
                        candidates.append({
                            "slide": slide,
                            "slide_idx": slide_idx,
                            "text_frame": text_frame,
                            "para_idx": para_idx,
                            "full_text": full,
                            "matched_text": match_result[0],
                            "strategy": match_result[1],
                            "score": match_result[2],
                        })

            # 表格
            for tf_info in self._iter_table_text_frames(slide.shapes, slide_idx):
                text_frame = tf_info["text_frame"]
                for para_idx, para in enumerate(text_frame.paragraphs):
                    runs = para.runs
                    if not runs:
                        continue
                    full = "".join(r.text for r in runs)
                    if not full.strip():
                        continue

                    match_result = self._try_match(full, old_text, old_clean, variants)
                    if match_result:
                        candidates.append({
                            "slide": slide,
                            "slide_idx": slide_idx,
                            "text_frame": text_frame,
                            "para_idx": para_idx,
                            "full_text": full,
                            "matched_text": match_result[0],
                            "strategy": match_result[1],
                            "score": match_result[2],
                        })

        # 按匹配质量排序（精确 > 清洗 > 变体 > 模糊）
        candidates.sort(key=lambda c: c["score"], reverse=True)
        return candidates

    def _try_match(
        self, full_text: str, old_text: str, old_clean: str, variants: List[str]
    ) -> Optional[Tuple[str, str, float]]:
        """
        多层匹配策略，返回 (实际匹配到的文本, 策略名, 分数) 或 None。
        分数越高越优先。
        """
        # 1. 精确匹配（最高优先）
        if old_text in full_text:
            return (old_text, "精确匹配", 1.0)

        full_clean = clean_text_thoroughly(full_text)

        # 2. 清洗后匹配（去除不可见字符、统一空格/Unicode）
        if old_clean and old_clean in full_clean:
            return (old_text, "清洗后匹配", 0.9)

        # 3. 忽略空格匹配
        full_no_sp = full_text.replace(" ", "").replace("\u3000", "").replace("\xa0", "")
        old_no_sp = old_text.replace(" ", "").replace("\u3000", "").replace("\xa0", "")
        if old_no_sp and old_no_sp in full_no_sp:
            return (old_text, "忽略空格匹配", 0.85)

        # 4. 变体匹配
        for v in variants:
            if v == old_text:
                continue
            if v in full_text:
                return (v, f"变体匹配({v})", 0.7)
            v_clean = clean_text_thoroughly(v)
            if v_clean and v_clean in full_clean:
                return (v, f"变体清洗匹配({v})", 0.65)

        # 5. TextMatcher 7层匹配（模糊兜底）
        found, start, end, strategy = self._matcher.find_best_match(
            full_text, old_text, return_strategy=True
        )
        if found and strategy and "精确" not in strategy:
            return (old_text, f"TextMatcher:{strategy}", 0.5)

        return None

    def _rank_by_context(self, candidates: List[Dict], context: str) -> List[Dict]:
        """根据上下文对候选项排序，上下文相似度高的排前面"""
        context_clean = clean_text_thoroughly(context)

        for c in candidates:
            # 获取段落所在幻灯片的周围文本作为上下文
            surrounding = self._get_surrounding_text(c)
            surrounding_clean = clean_text_thoroughly(surrounding)

            # 计算上下文相似度（字符重叠率）
            if context_clean and surrounding_clean:
                overlap = self._char_overlap(context_clean, surrounding_clean)
            else:
                overlap = 0.0

            # 综合分数 = 匹配质量 * 0.4 + 上下文相似度 * 0.6
            c["context_score"] = c["score"] * 0.4 + overlap * 0.6

        candidates.sort(key=lambda c: c["context_score"], reverse=True)
        return candidates

    def _get_surrounding_text(self, candidate: Dict) -> str:
        """获取候选匹配位置周围的文本（同一 text_frame 的所有段落）"""
        tf = candidate["text_frame"]
        parts = []
        for para in tf.paragraphs:
            t = "".join(r.text for r in para.runs)
            if t.strip():
                parts.append(t)
        return "\n".join(parts)

    @staticmethod
    def _char_overlap(text1: str, text2: str) -> float:
        """计算两段文本的字符重叠率（用于上下文验证）"""
        if not text1 or not text2:
            return 0.0
        shorter, longer = (text1, text2) if len(text1) <= len(text2) else (text2, text1)
        match_count = sum(1 for ch in shorter if ch in longer)
        return match_count / len(shorter) if shorter else 0.0

    # ── 替换辅助 ──

    def _replace_fuzzy_in_runs(self, runs, old_text: str, new_text: str) -> bool:
        """当精确替换失败时，用清洗后的文本定位并替换"""
        full = "".join(r.text for r in runs)
        full_clean = clean_text_thoroughly(full)
        old_clean = clean_text_thoroughly(old_text)

        if old_clean in full_clean:
            # 在清洗后文本中找到位置，在原文中尝试定位
            # 策略：把所有 run 文本合并后整体替换
            # 用 TextMatcher 找到原文中的最佳位置
            found, start, end, _ = self._matcher.find_best_match(full, old_text)
            if found and start is not None and end is not None:
                new_full = full[:start] + new_text + full[end:]
                self._redistribute_text(runs, new_full)
                return True

            # 兜底：忽略空格后定位
            full_no_sp = full.replace(" ", "")
            old_no_sp = old_text.replace(" ", "")
            if old_no_sp in full_no_sp:
                idx = full_no_sp.index(old_no_sp)
                # 映射回原文位置
                orig_start = self._map_no_space_pos(full, idx)
                orig_end = self._map_no_space_pos(full, idx + len(old_no_sp))
                new_full = full[:orig_start] + new_text + full[orig_end:]
                self._redistribute_text(runs, new_full)
                return True

        return False

    @staticmethod
    def _map_no_space_pos(text: str, no_space_pos: int) -> int:
        """将去空格后的位置映射回原文位置"""
        count = 0
        for i, ch in enumerate(text):
            if ch != " ":
                if count == no_space_pos:
                    return i
                count += 1
        return len(text)

    # ── 形状遍历 ──

    def _iter_text_frames(self, shapes, slide_idx: int):
        """递归遍历所有形状的 text_frame"""
        for shape in shapes:
            if shape.shape_type == 6:  # GroupShape
                try:
                    yield from self._iter_text_frames(shape.shapes, slide_idx)
                except Exception:
                    pass
            if shape.has_text_frame:
                yield {"text_frame": shape.text_frame, "slide_idx": slide_idx}

    def _iter_table_text_frames(self, shapes, slide_idx: int):
        """递归遍历所有表格单元格的 text_frame"""
        for shape in shapes:
            if shape.shape_type == 6:
                try:
                    yield from self._iter_table_text_frames(shape.shapes, slide_idx)
                except Exception:
                    pass
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        yield {"text_frame": cell.text_frame, "slide_idx": slide_idx}

    @staticmethod
    def _redistribute_text(runs, new_full: str):
        for i, run in enumerate(runs):
            run.text = new_full if i == 0 else ""

    # ── 批注逻辑（操作底层 OOXML）──

    def _ensure_author(self, name: str) -> int:
        """确保 commentAuthors.xml 中存在指定作者，返回 authorId"""
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        try:
            authors_part = self.prs.part.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
            )
            authors_xml = authors_part._element
        except Exception:
            authors_xml = self._create_comment_authors_part(name)
            if authors_xml is None:
                return 0

        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        for cmAuthor in authors_xml.findall(f"{{{P_NS}}}cmAuthor"):
            if cmAuthor.get("name") == name:
                return int(cmAuthor.get("id", "0"))

        existing = authors_xml.findall(f"{{{P_NS}}}cmAuthor")
        new_id = len(existing)
        new_author = etree.SubElement(authors_xml, f"{{{P_NS}}}cmAuthor")
        new_author.set("id", str(new_id))
        new_author.set("name", name)
        new_author.set("initials", name[0] if name else "A")
        new_author.set("lastIdx", "0")
        new_author.set("clrIdx", "0")
        return new_id

    def _create_comment_authors_part(self, name: str):
        """创建 commentAuthors.xml part"""
        try:
            from pptx.opc.part import Part
            from pptx.opc.packuri import PackURI

            P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
            authors_xml = etree.Element(f"{{{P_NS}}}cmAuthorLst")
            author = etree.SubElement(authors_xml, f"{{{P_NS}}}cmAuthor")
            author.set("id", "0")
            author.set("name", name)
            author.set("initials", name[0] if name else "A")
            author.set("lastIdx", "0")
            author.set("clrIdx", "0")

            content_type = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
            part_name = PackURI("/ppt/commentAuthors.xml")
            xml_bytes = etree.tostring(authors_xml, xml_declaration=True, encoding="UTF-8", standalone=True)

            authors_part = Part(
                part_name, content_type, xml_bytes, self.prs.part.package
            )
            self.prs.part.relate_to(
                authors_part,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
            )
            authors_part._element = authors_xml
            return authors_xml
        except Exception as e:
            print(f"  [警告] 创建 commentAuthors 失败: {e}")
            return None

    def _add_comment(self, slide, old_text: str, new_text: str, reason: str):
        """在幻灯片上添加真正的 PowerPoint 批注"""
        comment_text = f"'{old_text}' → '{new_text}'"
        if reason:
            comment_text += f"\n理由: {reason}"

        try:
            self._add_comment_xml(slide, comment_text)
            self.annotations_count += 1
        except Exception as e:
            print(f"  [警告] 批注添加失败({e})，改用备注记录")
            self._append_note_fallback(slide, old_text, new_text, reason)
            self.annotations_count += 1

    def _add_comment_xml(self, slide, text: str):
        """通过底层 XML 操作添加幻灯片批注"""
        from pptx.opc.part import Part
        from pptx.opc.packuri import PackURI

        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        CM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        CM_CT = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"

        slide_part = slide.part
        cm_list = None
        cm_part = None

        for rel in slide_part.rels.values():
            if rel.reltype == CM_REL:
                cm_part = rel.target_part
                cm_list = cm_part._element
                break

        if cm_list is None:
            slide_uri = str(slide_part.partname)
            slide_num = slide_uri.split("slide")[-1].replace(".xml", "")
            cm_uri = PackURI(f"/ppt/comments/comment{slide_num}.xml")

            cm_list = etree.Element(f"{{{P_NS}}}cmLst")
            xml_bytes = etree.tostring(cm_list, xml_declaration=True, encoding="UTF-8", standalone=True)
            cm_part = Part(cm_uri, CM_CT, xml_bytes, slide_part.package)
            slide_part.relate_to(cm_part, CM_REL)
            cm_part._element = cm_list

        existing = cm_list.findall(f"{{{P_NS}}}cm")
        next_idx = len(existing) + 1

        cm = etree.SubElement(cm_list, f"{{{P_NS}}}cm")
        cm.set("authorId", str(self._author_id))
        cm.set("dt", datetime.now().strftime("%Y-%m-%dT%H:%M:%S.000"))
        cm.set("idx", str(next_idx))

        pos = etree.SubElement(cm, f"{{{P_NS}}}pos")
        pos.set("x", "0")
        pos.set("y", "0")

        cm_text = etree.SubElement(cm, f"{{{P_NS}}}text")
        cm_text.text = text

        cm_part.blob = etree.tostring(cm_list, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _append_note_fallback(self, slide, old_text, new_text, reason):
        """备注 fallback"""
        if not slide.has_notes_slide:
            slide.notes_slide
        notes_tf = slide.notes_slide.notes_text_frame
        entry = f"[修改] '{old_text}' → '{new_text}'"
        if reason:
            entry += f"  理由: {reason}"
        p = notes_tf.add_paragraph()
        run = p.add_run()
        run.text = entry
