"""
读取 docx 的 Track Changes 修订，显示每处修订的前后上下文。
"""
import sys
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def qn(tag): return f"{{{W}}}{tag}"

def get_para_text_with_revisions(para_elem):
    """
    返回段落的完整文本，以及其中所有修订信息。
    修订信息格式：[(old_text, new_text, position_in_full_text), ...]
    """
    # 先收集所有文本片段及标记
    parts = []  # (text, is_del, is_ins)
    
    for child in para_elem.iter():
        tag = child.tag
        if tag == qn("t"):
            parent = child.getparent()
            if parent is None:
                continue
            gp = parent.getparent()
            if gp is None:
                continue
            text = child.text or ""
            
            # 判断是删除还是插入还是普通
            in_del = any(a.tag == qn("del") for a in _ancestors(child))
            in_ins = any(a.tag == qn("ins") for a in _ancestors(child))
            parts.append((text, in_del, in_ins))
        
        elif tag == qn("delText"):
            parts.append((child.text or "", True, False))
    
    return parts


def _ancestors(elem):
    ancestors = []
    p = elem.getparent()
    while p is not None:
        ancestors.append(p)
        p = p.getparent()
    return ancestors


def extract_revisions_from_docx(docx_path: str):
    """提取所有段落中的修订及上下文。"""
    from zipfile import ZipFile
    
    with ZipFile(docx_path) as zf:
        with zf.open("word/document.xml") as f:
            tree = etree.parse(f)
    
    results = []
    
    for para in tree.iter(qn("p")):
        # 检查段落是否有修订
        has_del = para.find(f".//{qn('del')}") is not None
        has_ins = para.find(f".//{qn('ins')}") is not None
        if not has_del and not has_ins:
            continue
        
        # 收集各部分文本
        normal_parts = []
        del_parts = []
        ins_parts = []
        
        # 遍历 para 的直接子元素（run/del/ins/...）
        def collect(elem, depth=0):
            tag = elem.tag
            if tag == qn("del"):
                texts = [t.text or "" for t in elem.iter(qn("delText"))]
                del_parts.append("".join(texts))
            elif tag == qn("ins"):
                texts = [t.text or "" for t in elem.iter(qn("t"))]
                ins_parts.append("".join(texts))
            elif tag == qn("r"):
                # 普通 run（不在 del/ins 内）
                # 检查是否在 del/ins 内
                in_rev = any(a.tag in (qn("del"), qn("ins")) for a in _ancestors(elem))
                if not in_rev:
                    texts = [t.text or "" for t in elem.iter(qn("t"))]
                    normal_parts.append("".join(texts))
            else:
                for child in elem:
                    collect(child, depth+1)
        
        for child in para:
            collect(child)
        
        normal_text = "".join(normal_parts)
        del_text = " | ".join(del_parts) if del_parts else ""
        ins_text = " | ".join(ins_parts) if ins_parts else ""
        
        if del_text or ins_text:
            results.append({
                "context": normal_text[:120],
                "deleted": del_text,
                "inserted": ins_text,
            })
    
    return results


def main():
    if len(sys.argv) < 2:
        # 默认路径
        docx_path = r"D:\project\数检_程序-AI\output\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_20260605_161554.docx"
    else:
        docx_path = sys.argv[1]
    
    if not Path(docx_path).exists():
        print(f"文件不存在: {docx_path}")
        sys.exit(1)
    
    print(f"读取: {Path(docx_path).name}")
    print("=" * 80)
    
    revisions = extract_revisions_from_docx(docx_path)
    
    if not revisions:
        print("未找到任何修订标记。")
        return
    
    print(f"共找到 {len(revisions)} 处修订\n")
    
    for i, rev in enumerate(revisions, 1):
        print(f"[{i:>3}] 上下文: {rev['context']!r}")
        if rev['deleted']:
            print(f"      删除: {rev['deleted']!r}")
        if rev['inserted']:
            print(f"      插入: {rev['inserted']!r}")
        print()


if __name__ == "__main__":
    main()
