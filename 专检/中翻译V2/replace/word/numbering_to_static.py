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
