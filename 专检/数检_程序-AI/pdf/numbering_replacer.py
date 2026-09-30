"""
Word自动编号替换模块
用于修改Word文档中的自动编号格式（如将 i., ii., iii. 改为 1), 2), 3)）
"""
import re
from typing import Optional, Tuple, Dict
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# 配置常量：是否为自动编号替换添加批注
# True: 为每个修改的编号段落添加批注（默认）
# False: 不添加批注
ENABLE_NUMBERING_COMMENTS = True


class NumberingReplacer:
    """处理Word文档中的自动编号替换"""
    
    # 编号格式映射
    FORMAT_MAP = {
        'lowerRoman': 'decimal',      # i., ii., iii. -> 1, 2, 3
        'upperRoman': 'decimal',      # I., II., III. -> 1, 2, 3
        'lowerLetter': 'decimal',     # a., b., c. -> 1, 2, 3
        'upperLetter': 'decimal',     # A., B., C. -> 1, 2, 3
    }
    
    # 编号文本格式映射
    TEXT_FORMAT_MAP = {
        '%1.': '%1)',   # 将点号改为右括号
        '%1': '%1)',    # 添加右括号
    }
    
    def __init__(self, doc: Document):
        """
        初始化编号替换器
        
        Args:
            doc: Word文档对象
        """
        self.doc = doc
        self.numbering_part = None
        self.modified_abstract_nums = set()  # 记录已修改的抽象编号
        self._load_numbering_part()
    
    def _load_numbering_part(self):
        """加载文档的编号部分"""
        try:
            self.numbering_part = self.doc.part.numbering_part
        except Exception as e:
            self.numbering_part = None
    
    def find_paragraphs_with_numbering(
        self, 
        target_format: str = 'lowerRoman',
        context_keywords: list = None
    ) -> list:
        """
        查找使用特定编号格式的所有段落
        
        Args:
            target_format: 目标编号格式（如 'lowerRoman'）
            context_keywords: 上下文关键词列表，用于过滤
        
        Returns:
            符合条件的段落列表
        """
        matching_paragraphs = []
        
        for paragraph in self.doc.paragraphs:
            try:
                # 检查段落是否有编号
                pPr = paragraph._element.pPr
                if pPr is None:
                    continue
                
                numPr = pPr.numPr
                if numPr is None:
                    continue
                
                # 获取编号ID和级别
                numId_elem = numPr.numId
                ilvl_elem = numPr.ilvl
                
                if numId_elem is None or ilvl_elem is None:
                    continue
                
                num_id = numId_elem.get(qn('w:val'))
                ilvl = ilvl_elem.get(qn('w:val'))
                
                # 检查编号格式
                num_format = self._get_numbering_format(num_id, ilvl)
                
                if num_format == target_format:
                    # 如果提供了上下文关键词，进行过滤
                    if context_keywords:
                        para_text = paragraph.text.lower()
                        if any(keyword.lower() in para_text for keyword in context_keywords):
                            matching_paragraphs.append((paragraph, num_id, ilvl))
                    else:
                        matching_paragraphs.append((paragraph, num_id, ilvl))
                        
            except Exception as e:
                continue
        
        return matching_paragraphs
    
    def _get_numbering_format(self, num_id: str, ilvl: str) -> Optional[str]:
        """
        获取指定编号的格式
        
        Args:
            num_id: 编号ID
            ilvl: 编号级别
        
        Returns:
            编号格式（如 'lowerRoman', 'decimal'）
        """
        if not self.numbering_part:
            return None
        
        try:
            numbering_xml = self.numbering_part.element
            
            # 查找编号定义
            for num in numbering_xml.findall(qn('w:num')):
                if num.get(qn('w:numId')) == num_id:
                    # 获取抽象编号ID
                    abstractNumId_elem = num.find(qn('w:abstractNumId'))
                    if abstractNumId_elem is None:
                        continue
                    
                    abstract_num_id = abstractNumId_elem.get(qn('w:val'))
                    
                    # 查找抽象编号定义
                    for abstract_num in numbering_xml.findall(qn('w:abstractNum')):
                        if abstract_num.get(qn('w:abstractNumId')) == abstract_num_id:
                            # 查找对应级别的格式
                            for lvl in abstract_num.findall(qn('w:lvl')):
                                if lvl.get(qn('w:ilvl')) == ilvl:
                                    numFmt_elem = lvl.find(qn('w:numFmt'))
                                    if numFmt_elem is not None:
                                        return numFmt_elem.get(qn('w:val'))
            
            return None
            
        except Exception as e:
            return None
    
    def modify_numbering_format(
        self, 
        num_id: str, 
        ilvl: str,
        new_format: str = 'decimal',
        new_text_format: str = '%1)'
    ) -> bool:
        """
        修改编号格式定义
        
        Args:
            num_id: 编号ID
            ilvl: 编号级别
            new_format: 新的编号格式（如 'decimal'）
            new_text_format: 新的文本格式（如 '%1)'）
        
        Returns:
            是否成功修改
        """
        if not self.numbering_part:
            return False
        
        try:
            numbering_xml = self.numbering_part.element
            
            # 查找编号定义
            for num in numbering_xml.findall(qn('w:num')):
                if num.get(qn('w:numId')) == num_id:
                    # 获取抽象编号ID
                    abstractNumId_elem = num.find(qn('w:abstractNumId'))
                    if abstractNumId_elem is None:
                        continue
                    
                    abstract_num_id = abstractNumId_elem.get(qn('w:val'))
                    
                    # 避免重复修改同一个抽象编号
                    if abstract_num_id in self.modified_abstract_nums:
                        return True
                    
                    # 查找抽象编号定义
                    for abstract_num in numbering_xml.findall(qn('w:abstractNum')):
                        if abstract_num.get(qn('w:abstractNumId')) == abstract_num_id:
                            # 查找对应级别的定义
                            for lvl in abstract_num.findall(qn('w:lvl')):
                                if lvl.get(qn('w:ilvl')) == ilvl:
                                    # 修改编号格式
                                    success = self._update_level_format(
                                        lvl, new_format, new_text_format
                                    )
                                    if success:
                                        self.modified_abstract_nums.add(abstract_num_id)
                                        return True
            
            return False
            
        except Exception as e:
            return False
    
    def _update_level_format(
        self, 
        lvl_element, 
        new_format: str,
        new_text_format: str
    ) -> bool:
        """
        更新级别格式
        
        Args:
            lvl_element: 级别元素
            new_format: 新格式（如 'decimal'）
            new_text_format: 新文本格式（如 '%1)'）
        
        Returns:
            是否成功更新
        """
        try:
            # 1. 修改编号格式类型（numFmt）
            numFmt_elem = lvl_element.find(qn('w:numFmt'))
            if numFmt_elem is not None:
                numFmt_elem.set(qn('w:val'), new_format)
            
            # 2. 修改编号文本格式（lvlText）
            lvlText_elem = lvl_element.find(qn('w:lvlText'))
            if lvlText_elem is not None:
                lvlText_elem.set(qn('w:val'), new_text_format)
            
            return True
            
        except Exception as e:
            return False
    
    def replace_all_numbering(
        self,
        old_format: str = 'lowerRoman',
        new_format: str = 'decimal',
        new_text_format: str = '%1)',
        context_keywords: list = None,
        comment_manager = None,
        old_number: str = "",
        new_number: str = "",
        reason: str = ""
    ) -> Tuple[int, int]:
        """
        批量替换所有匹配的编号格式
        
        Args:
            old_format: 旧格式（如 'lowerRoman'）
            new_format: 新格式（如 'decimal'）
            new_text_format: 新文本格式（如 '%1)'）
            context_keywords: 上下文关键词，用于过滤
            comment_manager: 批注管理器
            old_number: 旧编号示例（如 'i.'）
            new_number: 新编号示例（如 '1)'）
            reason: 修改理由
        
        Returns:
            (成功数量, 失败数量)
        """
        # 1. 查找所有使用旧格式的段落
        paragraphs = self.find_paragraphs_with_numbering(old_format, context_keywords)
        
        if not paragraphs:
            return 0, 0
        
        # 2. 收集需要修改的编号定义（去重）
        num_definitions = {}
        for para, num_id, ilvl in paragraphs:
            key = (num_id, ilvl)
            if key not in num_definitions:
                num_definitions[key] = []
            num_definitions[key].append(para)
        
        # 3. 修改编号定义
        success_count = 0
        fail_count = 0
        
        for (num_id, ilvl), paras in num_definitions.items():
            success = self.modify_numbering_format(num_id, ilvl, new_format, new_text_format)
            
            if success:
                success_count += len(paras)
                
                # 为每个段落添加批注（如果功能已开启）
                if ENABLE_NUMBERING_COMMENTS and comment_manager:
                    for para in paras:
                        # 获取段落的第一个run，如果没有则创建一个
                        if para.runs:
                            first_run = para.runs[0]
                        else:
                            # 如果段落没有run，创建一个空run
                            first_run = para.add_run("")
                        
                        # 构建批注文本
                        comment_text = f"【自动编号修改】\n"
                        comment_text += f"原格式: {old_number} (罗马数字/字母编号)\n"
                        comment_text += f"新格式: {new_number} (阿拉伯数字编号)\n"
                        if reason:
                            comment_text += f"修改理由: {reason}\n"
                        comment_text += f"批量修改: 此编号定义下的所有段落已统一修改"
                        
                        # 添加批注
                        try:
                            comment_manager.add_comment_to_run(first_run, comment_text)
                        except:
                            pass
            else:
                fail_count += len(paras)
        
        return success_count, fail_count


def replace_numbering_in_docx(
    doc: Document,
    old_number: str,
    new_number: str,
    context: str = "",
    comment_manager = None,
    reason: str = ""
) -> Tuple[bool, str]:
    """
    在Word文档中替换自动编号
    
    Args:
        doc: Word文档对象
        old_number: 旧编号（如 "i."）
        new_number: 新编号（如 "1)"）
        context: 上下文文本
        comment_manager: 批注管理器
        reason: 修改理由
    
    Returns:
        (是否成功, 策略描述)
    """
    replacer = NumberingReplacer(doc)
    
    # 确定旧格式类型
    old_format = None
    if re.match(r'^[ivxlcdm]+\.$', old_number.lower()):
        old_format = 'lowerRoman'
    elif re.match(r'^[IVXLCDM]+\.$', old_number):
        old_format = 'upperRoman'
    elif re.match(r'^[a-z]+\.$', old_number):
        old_format = 'lowerLetter'
    elif re.match(r'^[A-Z]+\.$', old_number):
        old_format = 'upperLetter'
    
    if not old_format:
        return False, f"无法识别编号格式: {old_number}"
    
    # 确定新格式
    new_format = 'decimal'  # 默认使用阿拉伯数字
    new_text_format = '%1)'  # 默认使用右括号
    
    # 从new_number中提取格式
    if ')' in new_number:
        new_text_format = '%1)'
    elif '.' in new_number:
        new_text_format = '%1.'
    
    # 提取上下文关键词
    context_keywords = []
    if context:
        # 提取前几个有意义的词
        words = context.split()[:10]
        context_keywords = [w for w in words if len(w) > 3]
    
    # 执行替换
    success_count, fail_count = replacer.replace_all_numbering(
        old_format, new_format, new_text_format, context_keywords,
        comment_manager, old_number, new_number, reason
    )
    
    if success_count > 0:
        return True, f"成功修改 {success_count} 个编号 (从 {old_number} 到 {new_number})"
    else:
        return False, f"未找到匹配的编号或修改失败"
