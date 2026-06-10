"""
Word 修订（Track Changes）管理器
在文档中以"修订模式"插入删除/新增标记，用户可在 Word 中接受或拒绝修改。
"""
import warnings
from datetime import datetime
from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

warnings.filterwarnings("ignore")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


class RevisionManager:
    """Word 修订管理器 —— 生成可撤回的 Track Changes 标记"""

    def __init__(self, doc: Document, author: str = "翻译校对"):
        self.doc = doc
        self.author = author
        self._rev_id = 0
        self._init_rev_id()
        # 确保文档设置中开启修订保护（打开时自动显示修订）
        self._enable_track_changes()

    # ------------------------------------------------------------------
    # 初始化
    # ------------------------------------------------------------------

    def _init_rev_id(self):
        """扫描文档中已有的修订 ID，避免冲突"""
        body = self.doc.element.body
        for elem in body.iter():
            rid = elem.get(qn("w:id"))
            if rid and rid.isdigit():
                self._rev_id = max(self._rev_id, int(rid))
        self._rev_id += 1

    def _next_rev_id(self) -> str:
        rid = self._rev_id
        self._rev_id += 1
        return str(rid)

    def _now_iso(self) -> str:
        return datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

    def _enable_track_changes(self):
        """在 settings 中标记文档处于修订模式"""
        try:
            settings = self.doc.settings.element
            # 如果已有 trackChanges 就不重复添加
            existing = settings.find(qn("w:trackChanges"))
            if existing is None:
                tc = OxmlElement("w:trackChanges")
                settings.append(tc)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 核心 API
    # ------------------------------------------------------------------

    def delete_run_text(self, run, reason: str = ""):
        """
        将一个 run 标记为"删除线"修订（红色删除线，可撤回）。

        Args:
            run:    python-docx 的 Run 对象
            reason: 可选的修订原因（不会显示在文档中，仅供日志）
        """
        r_elem = run._element
        p_elem = r_elem.getparent()
        if p_elem is None:
            return

        rev_id = self._next_rev_id()
        date = self._now_iso()

        # 构造 <w:del> 包裹
        del_elem = OxmlElement("w:del")
        del_elem.set(qn("w:id"), rev_id)
        del_elem.set(qn("w:author"), self.author)
        del_elem.set(qn("w:date"), date)

        # 把原 run 中的 <w:t> 改为 <w:delText>
        r_copy = self._clone_run_as_del(r_elem)
        del_elem.append(r_copy)

        # 用 del_elem 替换原 run
        p_elem.replace(r_elem, del_elem)

    def insert_text_after(self, anchor_run, new_text: str):
        """
        在 anchor_run 后面插入一段"新增"修订文本（绿色下划线，可撤回）。

        Args:
            anchor_run: python-docx 的 Run 对象，新文本插在它后面
            new_text:   要插入的文本
        """
        r_elem = anchor_run._element
        p_elem = r_elem.getparent()
        if p_elem is None:
            return

        rev_id = self._next_rev_id()
        date = self._now_iso()

        ins_elem = OxmlElement("w:ins")
        ins_elem.set(qn("w:id"), rev_id)
        ins_elem.set(qn("w:author"), self.author)
        ins_elem.set(qn("w:date"), date)

        # 构造新 run
        new_r = OxmlElement("w:r")

        # 复制原 run 的格式
        orig_rPr = r_elem.find(qn("w:rPr"))
        if orig_rPr is not None:
            new_r.append(self._deep_copy(orig_rPr))

        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = new_text
        new_r.append(t)

        ins_elem.append(new_r)

        # 插入到 anchor_run 后面
        idx = list(p_elem).index(r_elem)
        p_elem.insert(idx + 1, ins_elem)

    def insert_text_at_beginning(self, paragraph, new_text: str, reason: str = ""):
        """
        在段落最前面插入一段"新增"修订文本。

        用于替换自动编号：先删除 numPr（由调用方完成），再在段落开头插入新文本。

        Args:
            paragraph: python-docx 的 Paragraph 对象
            new_text:  要插入的文本
            reason:    修订原因（仅供日志）
        """
        p_elem = paragraph._element

        rev_id = self._next_rev_id()
        date = self._now_iso()

        ins_elem = OxmlElement("w:ins")
        ins_elem.set(qn("w:id"), rev_id)
        ins_elem.set(qn("w:author"), self.author)
        ins_elem.set(qn("w:date"), date)

        new_r = OxmlElement("w:r")

        # 尝试从段落第一个 run 复制格式
        first_r = p_elem.find(qn("w:r"))
        if first_r is not None:
            orig_rPr = first_r.find(qn("w:rPr"))
            if orig_rPr is not None:
                new_r.append(self._deep_copy(orig_rPr))

        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = new_text
        new_r.append(t)

        ins_elem.append(new_r)

        # 插入到 pPr 之后、第一个 run 之前
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(ins_elem)
        else:
            p_elem.insert(0, ins_elem)

    def replace_run_text(self, run, new_text: str, reason: str = ""):
        """
        替换一个 run 的文本：先标记删除旧文本，再插入新文本。
        在 Word 中显示为删除线 + 新增下划线，用户可逐条接受/拒绝。

        Args:
            run:      python-docx 的 Run 对象
            new_text: 替换后的文本
            reason:   修订原因（日志用）
        """
        r_elem = run._element
        p_elem = r_elem.getparent()
        if p_elem is None:
            return

        date = self._now_iso()

        # 1) 构造 <w:del> —— 删除旧文本
        del_id = self._next_rev_id()
        del_elem = OxmlElement("w:del")
        del_elem.set(qn("w:id"), del_id)
        del_elem.set(qn("w:author"), self.author)
        del_elem.set(qn("w:date"), date)
        del_elem.append(self._clone_run_as_del(r_elem))

        # 2) 构造 <w:ins> —— 插入新文本
        ins_id = self._next_rev_id()
        ins_elem = OxmlElement("w:ins")
        ins_elem.set(qn("w:id"), ins_id)
        ins_elem.set(qn("w:author"), self.author)
        ins_elem.set(qn("w:date"), date)

        new_r = OxmlElement("w:r")
        orig_rPr = r_elem.find(qn("w:rPr"))
        if orig_rPr is not None:
            new_r.append(self._deep_copy(orig_rPr))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = new_text
        new_r.append(t)
        ins_elem.append(new_r)

        # 3) 替换：先放 del，再放 ins
        idx = list(p_elem).index(r_elem)
        p_elem.remove(r_elem)
        p_elem.insert(idx, ins_elem)
        p_elem.insert(idx, del_elem)

    def replace_in_paragraph(self, paragraph, old_text: str, new_text: str, reason: str = ""):
        """
        在段落中查找 old_text 并替换为 new_text（修订模式）。
        处理文本可能跨多个 run 的情况。

        Args:
            paragraph: python-docx 的 Paragraph 对象
            old_text:  要查找的原文
            new_text:  替换后的文本
            reason:    修订原因
        Returns:
            bool: 是否成功替换
        """
        # 收集段落中所有 run 的文本和位置信息
        runs = paragraph.runs
        if not runs:
            return False

        full_text = "".join(r.text or "" for r in runs)
        start_pos = full_text.find(old_text)
        if start_pos == -1:
            return False

        end_pos = start_pos + len(old_text)

        # 找出 old_text 覆盖了哪些 run
        char_offset = 0
        affected_runs = []
        for r in runs:
            r_text = r.text or ""
            r_start = char_offset
            r_end = char_offset + len(r_text)

            if r_end > start_pos and r_start < end_pos:
                affected_runs.append((r, r_start, r_end))

            char_offset = r_end

        if not affected_runs:
            return False

        date = self._now_iso()
        p_elem = paragraph._element

        # 如果只涉及一个 run
        if len(affected_runs) == 1:
            run, r_start, r_end = affected_runs[0]
            r_text = run.text or ""
            local_start = start_pos - r_start
            local_end = end_pos - r_start

            prefix = r_text[:local_start]
            suffix = r_text[local_end:]
            r_elem = run._element
            idx = list(p_elem).index(r_elem)

            # 移除原 run
            p_elem.remove(r_elem)
            insert_at = idx

            # 前缀 run（保留部分）
            if prefix:
                pre_r = self._make_run_with_format(r_elem, prefix)
                p_elem.insert(insert_at, pre_r)
                insert_at += 1

            # 删除标记
            del_id = self._next_rev_id()
            del_elem = OxmlElement("w:del")
            del_elem.set(qn("w:id"), del_id)
            del_elem.set(qn("w:author"), self.author)
            del_elem.set(qn("w:date"), date)
            del_r = self._make_del_run_with_format(r_elem, old_text)
            del_elem.append(del_r)
            p_elem.insert(insert_at, del_elem)
            insert_at += 1

            # 插入标记
            ins_id = self._next_rev_id()
            ins_elem = OxmlElement("w:ins")
            ins_elem.set(qn("w:id"), ins_id)
            ins_elem.set(qn("w:author"), self.author)
            ins_elem.set(qn("w:date"), date)
            ins_r = self._make_run_with_format(r_elem, new_text)
            ins_elem.append(ins_r)
            p_elem.insert(insert_at, ins_elem)
            insert_at += 1

            # 后缀 run（保留部分）
            if suffix:
                suf_r = self._make_run_with_format(r_elem, suffix)
                p_elem.insert(insert_at, suf_r)

            return True

        # 多个 run 的情况：删除所有涉及的 run，整体替换
        first_run_elem = affected_runs[0][0]._element
        idx = list(p_elem).index(first_run_elem)

        # 处理第一个 run 的前缀
        first_r, first_start, first_end = affected_runs[0]
        prefix = (first_r.text or "")[:start_pos - first_start]

        # 处理最后一个 run 的后缀
        last_r, last_start, last_end = affected_runs[-1]
        suffix = (last_r.text or "")[end_pos - last_start:]

        # 移除所有涉及的 run
        for r, _, _ in affected_runs:
            p_elem.remove(r._element)

        insert_at = idx

        # 前缀
        if prefix:
            pre_r = self._make_run_with_format(first_run_elem, prefix)
            p_elem.insert(insert_at, pre_r)
            insert_at += 1

        # 删除标记
        del_id = self._next_rev_id()
        del_elem = OxmlElement("w:del")
        del_elem.set(qn("w:id"), del_id)
        del_elem.set(qn("w:author"), self.author)
        del_elem.set(qn("w:date"), date)
        del_r = self._make_del_run_with_format(first_run_elem, old_text)
        del_elem.append(del_r)
        p_elem.insert(insert_at, del_elem)
        insert_at += 1

        # 插入标记
        ins_id = self._next_rev_id()
        ins_elem = OxmlElement("w:ins")
        ins_elem.set(qn("w:id"), ins_id)
        ins_elem.set(qn("w:author"), self.author)
        ins_elem.set(qn("w:date"), date)
        ins_r = self._make_run_with_format(first_run_elem, new_text)
        ins_elem.append(ins_r)
        p_elem.insert(insert_at, ins_elem)
        insert_at += 1

        # 后缀
        if suffix:
            suf_r = self._make_run_with_format(first_run_elem, suffix)
            p_elem.insert(insert_at, suf_r)

        return True

    # ------------------------------------------------------------------
    # 内部工具方法
    # ------------------------------------------------------------------

    @staticmethod
    def _deep_copy(element):
        """深拷贝一个 lxml 元素"""
        from copy import deepcopy
        return deepcopy(element)

    @staticmethod
    def _clone_run_as_del(r_elem):
        """
        把一个 <w:r> 克隆为删除版本：<w:t> → <w:delText>
        """
        from copy import deepcopy
        new_r = deepcopy(r_elem)

        for t in new_r.findall(qn("w:t")):
            # 创建 delText 替换 t
            dt = OxmlElement("w:delText")
            dt.set(qn("xml:space"), "preserve")
            dt.text = t.text or ""
            parent = t.getparent()
            parent.replace(t, dt)

        return new_r

    @staticmethod
    def _make_run_with_format(ref_r_elem, text: str):
        """根据参考 run 的格式创建新 run"""
        new_r = OxmlElement("w:r")
        rPr = ref_r_elem.find(qn("w:rPr"))
        if rPr is not None:
            from copy import deepcopy
            new_r.append(deepcopy(rPr))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        new_r.append(t)
        return new_r

    @staticmethod
    def _make_del_run_with_format(ref_r_elem, text: str):
        """根据参考 run 的格式创建删除 run（用 delText）"""
        new_r = OxmlElement("w:r")
        rPr = ref_r_elem.find(qn("w:rPr"))
        if rPr is not None:
            from copy import deepcopy
            new_r.append(deepcopy(rPr))
        dt = OxmlElement("w:delText")
        dt.set(qn("xml:space"), "preserve")
        dt.text = text
        new_r.append(dt)
        return new_r
