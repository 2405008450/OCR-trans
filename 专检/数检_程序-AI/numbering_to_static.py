"""
将 Word 文档中的自动编号转换为静态文本（纯 Python，不依赖 COM/VBA）。

原理：
1. 从 numbering.xml 加载编号定义（含 lvlOverride）
2. 从 styles.xml 加载样式关联的编号
3. 按文档顺序遍历所有段落，计算每个段落的编号文本
4. 在段落开头插入编号文本的 run
5. 删除段落的 numPr（移除自动编号格式）
6. 保存文档
"""
import os
import copy
from zipfile import ZipFile
from typing import Dict, Optional, Tuple
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _collect_all_paragraphs_in_order(body, w_ns) -> list:
    """
    按文档顺序收集 body 下所有 w:p 元素，展开 w:sdt 内容控件包装（支持嵌套）。

    背景：Word 生成的 TOC 常被包在 w:sdt/w:sdtContent 里（"目录"内容控件），
    只扫 body 直接子元素会完全看不到这些段落，导致 TOC 域检测失败。
    不递归表格（维持原有范围：TOC 域本身不会出现在表格单元格里）。
    """
    W_p = f"{{{w_ns}}}p"
    W_sdt = f"{{{w_ns}}}sdt"
    W_sdtContent = f"{{{w_ns}}}sdtContent"

    paras = []

    def walk(container):
        for child in container:
            if child.tag == W_p:
                paras.append(child)
            elif child.tag == W_sdt:
                content = child.find(W_sdtContent)
                if content is not None:
                    walk(content)

    walk(body)
    return paras


def has_auto_numbering(doc_path: str) -> bool:
    """快速检测文档是否包含自动编号"""
    try:
        with ZipFile(doc_path, "r") as zf:
            if "word/numbering.xml" not in zf.namelist():
                return False
            with zf.open("word/document.xml") as f:
                content = f.read().decode("utf-8", errors="ignore")
                return "<w:numId " in content
    except Exception:
        return False


class _NumberingResolver:
    """从 numbering.xml 和 styles.xml 构建编号解析器"""

    def __init__(self, zf: ZipFile):
        self.numbering_map: Dict[str, Dict[str, dict]] = {}
        self.abstract_num_map: Dict[str, Dict[str, dict]] = {}
        self.style_num_map: Dict[str, Tuple[str, str]] = {}
        self.level_counters: Dict[Tuple[str, str], int] = {}
        self._load_numbering(zf)
        self._load_styles(zf)

    def _load_numbering(self, zf: ZipFile):
        if "word/numbering.xml" not in zf.namelist():
            return
        with zf.open("word/numbering.xml") as f:
            tree = etree.parse(f)

        for abs_num in tree.findall(".//w:abstractNum", NS):
            abs_id = abs_num.get(f"{{{W_NS}}}abstractNumId")
            self.abstract_num_map[abs_id] = {}
            for lvl in abs_num.findall(".//w:lvl", NS):
                ilvl = lvl.get(f"{{{W_NS}}}ilvl")
                nf = lvl.find(".//w:numFmt", NS)
                lt = lvl.find(".//w:lvlText", NS)
                st = lvl.find(".//w:start", NS)
                self.abstract_num_map[abs_id][ilvl] = {
                    "format": nf.get(f"{{{W_NS}}}val") if nf is not None else "decimal",
                    "text": lt.get(f"{{{W_NS}}}val") if lt is not None else "%1.",
                    "start": int(st.get(f"{{{W_NS}}}val", "1")) if st is not None else 1,
                }

        for num in tree.findall(".//w:num", NS):
            nid = num.get(f"{{{W_NS}}}numId")
            abs_elem = num.find(".//w:abstractNumId", NS)
            if abs_elem is None:
                continue
            abs_id = abs_elem.get(f"{{{W_NS}}}val")
            if abs_id not in self.abstract_num_map:
                continue
            lm = {}
            for k, v in self.abstract_num_map[abs_id].items():
                lm[k] = v.copy()
            # lvlOverride
            for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                ov_ilvl = override.get(f"{{{W_NS}}}ilvl")
                if ov_ilvl is None:
                    continue
                lvl = override.find(f"{{{W_NS}}}lvl")
                if lvl is not None:
                    oi = lm.get(ov_ilvl, {}).copy()
                    nf = lvl.find(".//w:numFmt", NS)
                    lt = lvl.find(".//w:lvlText", NS)
                    st = lvl.find(".//w:start", NS)
                    if nf is not None:
                        oi["format"] = nf.get(f"{{{W_NS}}}val")
                    if lt is not None:
                        oi["text"] = lt.get(f"{{{W_NS}}}val")
                    if st is not None:
                        oi["start"] = int(st.get(f"{{{W_NS}}}val", "1"))
                    lm[ov_ilvl] = oi
                so = override.find(f"{{{W_NS}}}startOverride")
                if so is not None and ov_ilvl in lm:
                    lm[ov_ilvl]["start"] = int(so.get(f"{{{W_NS}}}val", "1"))
            self.numbering_map[nid] = lm

    def _load_styles(self, zf: ZipFile):
        if "word/styles.xml" not in zf.namelist():
            return
        with zf.open("word/styles.xml") as f:
            tree = etree.parse(f)
        for style in tree.findall(".//w:style", NS):
            sid = style.get(f"{{{W_NS}}}styleId")
            if not sid:
                continue
            ppr = style.find(".//w:pPr", NS)
            if ppr is None:
                continue
            num_pr = ppr.find(".//w:numPr", NS)
            if num_pr is None:
                continue
            nid_e = num_pr.find(".//w:numId", NS)
            ilvl_e = num_pr.find(".//w:ilvl", NS)
            nid = nid_e.get(f"{{{W_NS}}}val") if nid_e is not None else None
            ilvl = ilvl_e.get(f"{{{W_NS}}}val") if ilvl_e is not None else "0"
            if nid:
                self.style_num_map[sid] = (nid, ilvl)

    @staticmethod
    def _to_chinese(num: int) -> str:
        cn = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
        units = ["", "十", "百", "千", "万"]
        if num == 0:
            return cn[0]
        result = ""
        ui = 0
        n = num
        while n > 0:
            d = n % 10
            if d != 0:
                result = cn[d] + units[ui] + result
            elif result and result[0] != "零":
                result = cn[0] + result
            n //= 10
            ui += 1
        if result.startswith("一十"):
            result = result[1:]
        return result.rstrip("零")

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

    def _format_number(self, num: int, fmt: str) -> str:
        if fmt == "decimal":
            return str(num)
        if fmt == "upperRoman":
            return self._to_roman(num).upper()
        if fmt == "lowerRoman":
            return self._to_roman(num).lower()
        if fmt == "upperLetter":
            return chr(64 + num) if 1 <= num <= 26 else str(num)
        if fmt == "lowerLetter":
            return chr(96 + num) if 1 <= num <= 26 else str(num)
        if fmt in ("chineseCountingThousand", "chineseCounting",
                    "japaneseCounting", "japaneseDigitalTenThousand",
                    "ideographTraditional"):
            return self._to_chinese(num)
        if fmt == "bullet":
            return "•"
        return str(num)

    def resolve(self, p_elem) -> Optional[str]:
        """计算段落的自动编号文本，返回 None 表示无编号"""
        pPr = p_elem.find(f"{{{W_NS}}}pPr")
        if pPr is None:
            return None

        numPr = pPr.find(f"{{{W_NS}}}numPr")
        nid = None
        ilvl = "0"

        if numPr is not None:
            nid_e = numPr.find(f"{{{W_NS}}}numId")
            ilvl_e = numPr.find(f"{{{W_NS}}}ilvl")
            if nid_e is not None:
                nid = nid_e.get(f"{{{W_NS}}}val")
            if ilvl_e is not None:
                ilvl = ilvl_e.get(f"{{{W_NS}}}val")
        else:
            # 从样式查找
            pstyle = pPr.find(f"{{{W_NS}}}pStyle")
            if pstyle is not None:
                sid = pstyle.get(f"{{{W_NS}}}val")
                if sid and sid in self.style_num_map:
                    nid, ilvl = self.style_num_map[sid]

        if nid is None or nid == "0":
            return None
        if nid not in self.numbering_map or ilvl not in self.numbering_map[nid]:
            return None

        li = self.numbering_map[nid][ilvl]
        ck = (nid, ilvl)
        if ck not in self.level_counters:
            self.level_counters[ck] = li["start"]
        else:
            self.level_counters[ck] += 1

        # 重置子级别
        ilvl_int = int(ilvl)
        for oi_str in list(self.numbering_map[nid].keys()):
            if int(oi_str) > ilvl_int:
                ok = (nid, oi_str)
                if ok in self.level_counters:
                    del self.level_counters[ok]

        # 多级占位符替换
        tmpl = li["text"]
        for li_idx in range(ilvl_int + 1):
            ph = f"%{li_idx + 1}"
            if ph not in tmpl:
                continue
            ls = str(li_idx)
            lk = (nid, ls)
            if ls in self.numbering_map[nid] and lk in self.level_counters:
                linfo = self.numbering_map[nid][ls]
                formatted = self._format_number(self.level_counters[lk], linfo["format"])
                tmpl = tmpl.replace(ph, formatted)

        return tmpl


def convert_toc_to_static(doc_path: str) -> bool:
    """
    将 docx 中目录（TOC）域代码转为静态文本。

    Word TOC 结构：
    - 外层：一个大 TOC 域（begin→instrText→separate→[段落们]→end）
    - 每个目录条目段落内：HYPERLINK 子域（begin→instrText→separate→[文本run]→end）
    - 文本实际在 HYPERLINK 子域的缓存 run 里

    策略：找到 TOC 域覆盖的所有段落，对每个段落内的所有域（HYPERLINK等）展开为静态文本。

    Args:
        doc_path: docx 文件路径
    Returns:
        是否成功
    """
    import shutil
    import tempfile

    abs_path = os.path.abspath(doc_path)
    if not os.path.exists(abs_path):
        print(f"    [目录静态化] 文件不存在: {abs_path}")
        return False

    W = W_NS
    W_r = f"{{{W}}}r"
    W_p = f"{{{W}}}p"
    W_body = f"{{{W}}}body"
    W_fldChar = f"{{{W}}}fldChar"
    W_instrText = f"{{{W}}}instrText"
    W_t = f"{{{W}}}t"
    W_rPr = f"{{{W}}}rPr"
    fldCharType = f"{{{W}}}fldCharType"

    try:
        with ZipFile(abs_path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return False
            with zf.open("word/document.xml") as f:
                content_bytes = f.read()
            content_str = content_bytes.decode("utf-8", errors="ignore")
            if "TOC" not in content_str:
                print(f"    [目录静态化] 未检测到目录域，跳过")
                return True
            all_names = zf.namelist()

        tree = etree.fromstring(content_bytes)

        def _collect_field_range(elements, start_idx):
            """从 start_idx（begin run）开始，收集整个域的范围，返回 (separate_idx, end_idx)"""
            depth = 1
            separate_idx = None
            end_idx = None
            j = start_idx + 1
            while j < len(elements):
                r = elements[j]
                if r.tag == W_r:
                    fc = r.find(W_fldChar)
                    if fc is not None:
                        ft = fc.get(fldCharType)
                        if ft == "begin":
                            depth += 1
                        elif ft == "separate" and depth == 1:
                            separate_idx = j
                        elif ft == "end":
                            depth -= 1
                            if depth == 0:
                                end_idx = j
                                break
                j += 1
            return separate_idx, end_idx

        def _is_toc_instr(elements, begin_idx, separate_idx):
            """检查域的 instrText 是否是 TOC"""
            end = separate_idx if separate_idx is not None else len(elements)
            for k in range(begin_idx + 1, end):
                r = elements[k]
                if r.tag == W_r:
                    instr = r.find(W_instrText)
                    if instr is not None and instr.text:
                        txt = instr.text.strip().upper()
                        if txt.startswith("TOC"):
                            return True
            return False

        def _unlink_all_fields_in_paragraph(p):
            """
            展开段落内所有域为静态文本：
            1. 将 w:hyperlink 子元素内的 run 提升到段落直接子元素，删除 hyperlink 包装
            2. 展开段落直接子元素中的所有 fldChar 域（PAGEREF 等）为缓存值
            """
            W_hyperlink = f"{{{W}}}hyperlink"
            changed = False

            # 第一步：展开 w:hyperlink，把内部 run 提升到段落层
            while True:
                children = list(p)
                found_hl = False
                for i, elem in enumerate(children):
                    if elem.tag != W_hyperlink:
                        continue
                    # 收集 hyperlink 内所有 run（深拷贝）
                    inner_runs = [copy.deepcopy(r) for r in elem if r.tag == W_r]
                    # 插入到 hyperlink 位置
                    for ci, r in enumerate(inner_runs):
                        if i + ci < len(children):
                            children[i + ci].addprevious(r)
                        else:
                            p.append(r)
                    p.remove(elem)
                    found_hl = True
                    changed = True
                    break
                if not found_hl:
                    break

            # 第二步：展开段落内所有 fldChar 域（PAGEREF 等）为缓存值
            while True:
                children = list(p)
                found = False
                for i, elem in enumerate(children):
                    if elem.tag != W_r:
                        continue
                    fc = elem.find(W_fldChar)
                    if fc is None or fc.get(fldCharType) != "begin":
                        continue

                    separate_idx, end_idx = _collect_field_range(children, i)
                    if end_idx is None:
                        continue

                    # 提取缓存 run
                    cached_runs = []
                    if separate_idx is not None:
                        for k in range(separate_idx + 1, end_idx):
                            r = children[k]
                            if r.tag == W_r:
                                cr = copy.deepcopy(r)
                                for bad_tag in [W_fldChar, W_instrText]:
                                    for bad in cr.findall(bad_tag):
                                        cr.remove(bad)
                                t = cr.find(W_t)
                                if t is not None and t.text:
                                    cached_runs.append(cr)

                    # 从后往前删除域元素
                    for k in range(end_idx, i - 1, -1):
                        if k < len(children):
                            p.remove(children[k])

                    # 插入缓存 run
                    new_children = list(p)
                    insert_pos = min(i, len(new_children))
                    for ci, cr in enumerate(cached_runs):
                        if insert_pos + ci < len(new_children):
                            new_children[insert_pos + ci].addprevious(cr)
                        else:
                            p.append(cr)

                    found = True
                    changed = True
                    break

                if not found:
                    break

            return changed

        # 找到 TOC 域覆盖的段落范围
        # TOC 域是跨段落的：begin 在某段落，end 在另一段落
        # 需要在 body 层面扫描
        body = tree.find(f".//{W_body}")
        if body is None:
            print(f"    [目录静态化] 未找到 body 元素")
            return False

        # body_children 展开 w:sdt 包装（Word "目录"内容控件常把整个 TOC
        # 包在 sdtContent 里），否则下面的扁平化扫描会完全看不到 TOC 域
        body_children = _collect_all_paragraphs_in_order(body, W)
        toc_para_ranges = []  # [(start_para_idx, end_para_idx)]

        # 扫描所有段落（已展开 sdt），找跨段落的 TOC 域
        # TOC 域的 begin/end run 可能在不同段落里
        # 用扁平化方式：收集所有 (para_idx, run) 对
        flat = []  # (para_idx, run_elem)
        for pi, child in enumerate(body_children):
            if child.tag == W_p:
                for r in child:
                    if r.tag == W_r:
                        flat.append((pi, r))

        i = 0
        while i < len(flat):
            pi, r = flat[i]
            fc = r.find(W_fldChar)
            if fc is None or fc.get(fldCharType) != "begin":
                i += 1
                continue

            # 检查是否 TOC
            is_toc = False
            j = i + 1
            while j < len(flat):
                pj, rj = flat[j]
                instr = rj.find(W_instrText)
                if instr is not None and instr.text:
                    if instr.text.strip().upper().startswith("TOC"):
                        is_toc = True
                fc2 = rj.find(W_fldChar)
                if fc2 is not None and fc2.get(fldCharType) in ("separate", "end"):
                    break
                j += 1

            if not is_toc:
                i += 1
                continue

            # 找 TOC 域的 end
            depth = 1
            end_pi = None
            j = i + 1
            while j < len(flat):
                pj, rj = flat[j]
                fc2 = rj.find(W_fldChar)
                if fc2 is not None:
                    ft = fc2.get(fldCharType)
                    if ft == "begin":
                        depth += 1
                    elif ft == "end":
                        depth -= 1
                        if depth == 0:
                            end_pi = pj
                            break
                j += 1

            if end_pi is not None:
                toc_para_ranges.append((pi, end_pi))
                print(f"    [目录静态化] 找到 TOC 域，覆盖段落 {pi} ~ {end_pi}")
            i = j + 1

        if not toc_para_ranges:
            print(f"    [目录静态化] 未找到 TOC 域代码")
            return True

        # 对 TOC 覆盖范围内的所有段落展开子域
        total_unlinked = 0
        for start_pi, end_pi in toc_para_ranges:
            for pi in range(start_pi, end_pi + 1):
                child = body_children[pi]
                if child.tag == W_p:
                    if _unlink_all_fields_in_paragraph(child):
                        total_unlinked += 1

        if total_unlinked == 0:
            print(f"    [目录静态化] TOC 域内未找到子域需要展开")
            return True

        # 回写 docx
        fd, tmp_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            new_xml = etree.tostring(tree, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
            with ZipFile(abs_path, "r") as zf_in, ZipFile(tmp_path, "w") as zf_out:
                for name in all_names:
                    if name == "word/document.xml":
                        zf_out.writestr(name, new_xml)
                    else:
                        zf_out.writestr(name, zf_in.read(name))
            shutil.move(tmp_path, abs_path)
            print(f"    [目录静态化] 成功展开 {total_unlinked} 个目录条目段落")
            return True
        except Exception:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            raise

    except Exception as e:
        print(f"    [目录静态化] 转换失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def convert_numbering_to_static(doc_path: str) -> bool:
    """
    纯 Python 实现：将 docx 中所有自动编号转为静态文本。

    处理 document.xml + 所有 header*.xml / footer*.xml：
    1. 计算每个段落的编号文本
    2. 在段落第一个 run 之前插入一个新 run（包含编号文本）
    3. 删除段落的 numPr
    4. 回写 docx

    Args:
        doc_path: docx 文件路径

    Returns:
        是否成功
    """
    abs_path = os.path.abspath(doc_path)
    if not os.path.exists(abs_path):
        print(f"    [编号静态化] 文件不存在: {abs_path}")
        return False

    if not has_auto_numbering(abs_path):
        print(f"    [编号静态化] 文档无自动编号，跳过")
        return True

    try:
        import shutil
        import tempfile

        # 读取 zip 内容
        with ZipFile(abs_path, "r") as zf:
            resolver = _NumberingResolver(zf)
            all_names = zf.namelist()

            # 收集需要处理的 XML 文件：document.xml + header*.xml + footer*.xml
            xml_targets = []
            for name in all_names:
                if name == "word/document.xml":
                    xml_targets.append(name)
                elif name.startswith("word/header") and name.endswith(".xml"):
                    xml_targets.append(name)
                elif name.startswith("word/footer") and name.endswith(".xml"):
                    xml_targets.append(name)

            # 解析所有目标 XML
            parsed_trees = {}
            for name in xml_targets:
                with zf.open(name) as f:
                    parsed_trees[name] = etree.parse(f)

        count = 0

        def _staticize_paragraphs(tree):
            """对一棵 XML 树中的所有段落执行编号静态化"""
            nonlocal count
            root = tree.getroot()
            for p in root.iter(f"{{{W_NS}}}p"):
                num_text = resolver.resolve(p)
                if num_text is None:
                    continue
                print(f"正在转换第 {count + 1} 处编号: {num_text}")
                pPr = p.find(f"{{{W_NS}}}pPr")

                # 构造新 run
                new_r = etree.SubElement(p, f"{{{W_NS}}}r")
                first_r = p.find(f"{{{W_NS}}}r")
                if first_r is not None and first_r is not new_r:
                    first_rPr = first_r.find(f"{{{W_NS}}}rPr")
                    if first_rPr is not None:
                        new_r.insert(0, copy.deepcopy(first_rPr))

                new_t = etree.SubElement(new_r, f"{{{W_NS}}}t")
                new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                new_t.text = num_text

                # 移到段落开头
                p.remove(new_r)
                if pPr is not None:
                    pPr.addnext(new_r)
                else:
                    p.insert(0, new_r)

                # 删除 numPr
                if pPr is not None:
                    numPr = pPr.find(f"{{{W_NS}}}numPr")
                    if numPr is not None:
                        pPr.remove(numPr)

                count += 1

        # 处理所有目标 XML
        for name, tree in parsed_trees.items():
            _staticize_paragraphs(tree)

        if count == 0:
            print(f"    [编号静态化] 未找到需要转换的编号段落")
            return True

        # 回写 docx
        fd, tmp_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)

        try:
            with ZipFile(abs_path, "r") as zf_in, ZipFile(tmp_path, "w") as zf_out:
                for name in all_names:
                    if name in parsed_trees:
                        new_xml = etree.tostring(parsed_trees[name],
                                                  xml_declaration=True,
                                                  encoding="UTF-8",
                                                  standalone=True)
                        zf_out.writestr(name, new_xml)
                    else:
                        zf_out.writestr(name, zf_in.read(name))

            shutil.move(tmp_path, abs_path)
            print(f"    [编号静态化] 成功转换 {count} 个编号段落（含页眉页脚）")
            return True

        except Exception:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            raise

    except Exception as e:
        print(f"    [编号静态化] 转换失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    e=convert_numbering_to_static(
        r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\【批註】TP260331011_007 议案19-附件：合规管理及合规文化建设情况工作报告_edited_v2.docx")
    print(e)


# def unlink_page_fields_vba(doc_path: str) -> bool:
#     """
#     用 VBA（cscript + VBS）将页眉页脚中的域代码（PAGE、NUMPAGES 等）
#     转为静态文本（Field.Unlink），使页码变成实际数字。
#
#     步骤：
#     1. 先 Fields.Update 更新所有域（确保缓存值是正确的）
#     2. 遍历所有节的页眉页脚，对每个域执行 Unlink
#     3. 保存文档
#
#     Args:
#         doc_path: docx 文件的绝对路径
#
#     Returns:
#         是否成功
#     """
#     import subprocess
#     import tempfile
#
#     abs_path = os.path.abspath(doc_path)
#     if not os.path.exists(abs_path):
#         print(f"    [域静态化] 文件不存在: {abs_path}")
#         return False
#
#     # 快速检测是否有域代码
#     try:
#         with ZipFile(abs_path, "r") as zf:
#             has_field = False
#             for name in zf.namelist():
#                 if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
#                     with zf.open(name) as f:
#                         content = f.read().decode("utf-8", errors="ignore")
#                         if "fldChar" in content or "instrText" in content:
#                             has_field = True
#                             break
#         if not has_field:
#             print(f"    [域静态化] 页眉页脚无域代码，跳过")
#             return True
#     except Exception:
#         pass
#
#     escaped = abs_path.replace("\\", "\\\\")
#
#     vbs_content = f'''
# Dim objWord
# Dim objDoc
# Dim sec
# Dim hf
# Dim fld
# Dim count
#
# On Error Resume Next
#
# Set objWord = CreateObject("Word.Application")
# objWord.Visible = False
# objWord.DisplayAlerts = 0
#
# Set objDoc = objWord.Documents.Open("{escaped}")
#
# If Err.Number <> 0 Then
#     WScript.Echo "ERROR: " & Err.Description
#     objWord.Quit False
#     WScript.Quit 1
# End If
#
# On Error Resume Next
#
# ' 先更新所有域，确保页码缓存值正确
# objDoc.Fields.Update
#
# count = 0
#
# ' 遍历所有节的页眉页脚
# For Each sec In objDoc.Sections
#     ' 页眉
#     Dim i
#     For i = 1 To 3
#         Dim hfObj
#         On Error Resume Next
#         Set hfObj = sec.Headers(i)
#         If Err.Number = 0 And Not hfObj Is Nothing Then
#             If hfObj.Exists Then
#                 Dim j
#                 ' 从后往前 Unlink，避免索引偏移
#                 For j = hfObj.Range.Fields.Count To 1 Step -1
#                     On Error Resume Next
#                     hfObj.Range.Fields(j).Unlink
#                     If Err.Number = 0 Then count = count + 1
#                     Err.Clear
#                 Next j
#             End If
#         End If
#         Err.Clear
#     Next i
#
#     ' 页脚
#     For i = 1 To 3
#         Dim ffObj
#         On Error Resume Next
#         Set ffObj = sec.Footers(i)
#         If Err.Number = 0 And Not ffObj Is Nothing Then
#             If ffObj.Exists Then
#                 For j = ffObj.Range.Fields.Count To 1 Step -1
#                     On Error Resume Next
#                     ffObj.Range.Fields(j).Unlink
#                     If Err.Number = 0 Then count = count + 1
#                     Err.Clear
#                 Next j
#             End If
#         End If
#         Err.Clear
#     Next i
# Next
#
# objDoc.Save
# objDoc.Close False
# objWord.Quit False
#
# WScript.Echo "OK: " & count & " fields unlinked"
# '''
#
#     vbs_path = None
#     try:
#         fd, vbs_path = tempfile.mkstemp(suffix=".vbs", prefix="unlink_fields_")
#         os.close(fd)
#         with open(vbs_path, "w", encoding="utf-8") as f:
#             f.write(vbs_content)
#
#         print(f"    [域静态化] 正在调用 Word VBA 转换页眉页脚域代码...")
#
#         result = subprocess.run(
#             ["cscript", "//Nologo", vbs_path],
#             capture_output=True, text=True, timeout=120,
#             encoding="utf-8", errors="replace"
#         )
#
#         output = (result.stdout or "").strip()
#         error = (result.stderr or "").strip()
#
#         if result.returncode != 0:
#             print(f"    [域静态化] 执行失败 (code={result.returncode})")
#             if error:
#                 print(f"    [域静态化] stderr: {error}")
#             if output:
#                 print(f"    [域静态化] stdout: {output}")
#             return False
#
#         if output.startswith("OK:"):
#             print(f"    [域静态化] {output}")
#             return True
#         elif output.startswith("ERROR:"):
#             print(f"    [域静态化] {output}")
#             return False
#         else:
#             print(f"    [域静态化] 输出: {output}")
#             return True
#
#     except subprocess.TimeoutExpired:
#         print(f"    [域静态化] 执行超时（120秒）")
#         return False
#     except Exception as e:
#         print(f"    [域静态化] 异常: {e}")
#         return False
#     finally:
#         if vbs_path and os.path.exists(vbs_path):
#             try:
#                 os.remove(vbs_path)
#             except:
#                 pass
#
#
# def unlink_page_fields_xml(doc_path: str) -> bool:
#     """
#     纯 XML 方案：将页眉页脚中的域代码（PAGE、NUMPAGES 等）转为静态文本。
#
#     使用域代码的缓存值（上次 Word 保存时写入的实际页码数字）。
#     把域代码的 run 结构简化为只保留缓存值的普通 run，删除 fldChar/instrText。
#
#     注意：缓存值是上次 Word 保存时的值。如果文档内容改动过但未重新用 Word 打开，
#     缓存值可能不准确。如需准确页码请先用 Word 打开保存一次再调用此函数。
#
#     Args:
#         doc_path: docx 文件路径
#
#     Returns:
#         是否成功
#     """
#     import shutil
#     import tempfile
#
#     abs_path = os.path.abspath(doc_path)
#     if not os.path.exists(abs_path):
#         print(f"    [域静态化XML] 文件不存在: {abs_path}")
#         return False
#
#     W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
#
#     def _has_field_in_hf(zf):
#         for name in zf.namelist():
#             if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
#                 with zf.open(name) as f:
#                     content = f.read().decode("utf-8", errors="ignore")
#                     if "fldChar" in content:
#                         return True
#         return False
#
#     try:
#         with ZipFile(abs_path, "r") as zf:
#             if not _has_field_in_hf(zf):
#                 print(f"    [域静态化XML] 页眉页脚无域代码，跳过")
#                 return True
#
#             all_names = zf.namelist()
#
#             # 收集页眉页脚 XML
#             hf_targets = []
#             for name in all_names:
#                 if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
#                     hf_targets.append(name)
#
#             parsed_trees = {}
#             for name in hf_targets:
#                 with zf.open(name) as f:
#                     parsed_trees[name] = etree.parse(f)
#
#         total_unlinked = 0
#
#         def _unlink_fields_in_tree(tree):
#             """把树中所有域代码替换为缓存值的静态 run"""
#             nonlocal total_unlinked
#             root = tree.getroot()
#
#             for p in root.iter(f"{{{W}}}p"):
#                 _unlink_fields_in_paragraph(p)
#
#         def _unlink_fields_in_paragraph(p):
#             """处理单个段落中的域代码"""
#             nonlocal total_unlinked
#             W_r = f"{{{W}}}r"
#             W_fldChar = f"{{{W}}}fldChar"
#             W_instrText = f"{{{W}}}instrText"
#             W_t = f"{{{W}}}t"
#             W_rPr = f"{{{W}}}rPr"
#             fldCharType = f"{{{W}}}fldCharType"
#
#             # 收集段落直接子元素（只处理直接子 run，不深入文本框）
#             children = list(p)
#
#             i = 0
#             while i < len(children):
#                 elem = children[i]
#                 if elem.tag != W_r:
#                     i += 1
#                     continue
#
#                 fld = elem.find(W_fldChar)
#                 if fld is None or fld.get(fldCharType) != "begin":
#                     i += 1
#                     continue
#
#                 # 找到域的开始，收集整个域的 run 范围
#                 begin_idx = i
#                 separate_idx = None
#                 end_idx = None
#
#                 j = i + 1
#                 while j < len(children):
#                     r = children[j]
#                     if r.tag != W_r:
#                         j += 1
#                         continue
#                     fc = r.find(W_fldChar)
#                     if fc is not None:
#                         ft = fc.get(fldCharType)
#                         if ft == "separate":
#                             separate_idx = j
#                         elif ft == "end":
#                             end_idx = j
#                             break
#                     j += 1
#
#                 if end_idx is None:
#                     i += 1
#                     continue
#
#                 # 提取缓存值（separate 和 end 之间的 w:t 文本）
#                 cached_text = ""
#                 cached_rPr = None
#                 if separate_idx is not None:
#                     for k in range(separate_idx + 1, end_idx):
#                         r = children[k]
#                         if r.tag != W_r:
#                             continue
#                         t = r.find(W_t)
#                         if t is not None and t.text:
#                             cached_text += t.text
#                         if cached_rPr is None:
#                             cached_rPr = r.find(W_rPr)
#
#                 # 删除域的所有 run（begin 到 end，含两端）
#                 for k in range(end_idx, begin_idx - 1, -1):
#                     if k < len(children) and children[k].tag == W_r:
#                         p.remove(children[k])
#
#                 # 如果有缓存值，插入静态 run
#                 if cached_text:
#                     new_r = etree.Element(W_r)
#                     if cached_rPr is not None:
#                         new_r.append(copy.deepcopy(cached_rPr))
#                     new_t = etree.SubElement(new_r, W_t)
#                     new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
#                     new_t.text = cached_text
#
#                     # 插入到 begin_idx 位置
#                     # 重新获取 children（已经删除了一些元素）
#                     new_children = list(p)
#                     insert_pos = min(begin_idx, len(new_children))
#                     if insert_pos < len(new_children):
#                         new_children[insert_pos].addprevious(new_r)
#                     else:
#                         p.append(new_r)
#
#                     total_unlinked += 1
#
#                 # 重新获取 children 继续处理
#                 children = list(p)
#                 i = begin_idx  # 从当前位置重新检查
#
#         for name, tree in parsed_trees.items():
#             _unlink_fields_in_tree(tree)
#
#         if total_unlinked == 0:
#             print(f"    [域静态化XML] 未找到可处理的域代码")
#             return True
#
#         # 回写 docx
#         fd, tmp_path = tempfile.mkstemp(suffix=".docx")
#         os.close(fd)
#
#         try:
#             with ZipFile(abs_path, "r") as zf_in, ZipFile(tmp_path, "w") as zf_out:
#                 for name in all_names:
#                     if name in parsed_trees:
#                         new_xml = etree.tostring(parsed_trees[name],
#                                                   xml_declaration=True,
#                                                   encoding="UTF-8",
#                                                   standalone=True)
#                         zf_out.writestr(name, new_xml)
#                     else:
#                         zf_out.writestr(name, zf_in.read(name))
#
#             shutil.move(tmp_path, abs_path)
#             print(f"    [域静态化XML] 成功转换 {total_unlinked} 个域代码为静态文本")
#             return True
#
#         except Exception:
#             if os.path.exists(tmp_path):
#                 os.remove(tmp_path)
#             raise
#
#     except Exception as e:
#         print(f"    [域静态化XML] 转换失败: {e}")
#         import traceback
#         traceback.print_exc()
#         return False
