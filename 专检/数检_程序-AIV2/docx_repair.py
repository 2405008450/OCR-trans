"""
修复 DOCX 中指向缺失部件的悬空关系（dangling relationships）。

背景：
某些文档的 .rels 中存在损坏的内部关系，例如：
    <Relationship Id="rId110"
                  Type=".../image"
                  Target="NULL" />
python-docx 打开文档时会遍历所有 .rels 并按 Internal 关系去 zip 里读取对应部件，
若 Target 指向的部件（如 word/NULL）不存在，会抛出：
    KeyError: "There is no item named 'word/NULL' in the archive"
导致连 Document() 都打不开，无法通过 python-docx 的 save 修复。

本模块在 zip 层直接清理这些悬空关系（与 numbering_to_static 相同的重打包套路）：
1. 遍历所有 *.rels
2. 对每条 TargetMode 非 External 的关系，按 rels 所在目录解析 Target 相对路径
3. 若目标部件不在 zip namelist 中 → 删除该 <Relationship> 节点
4. 有改动的 .rels 重新序列化并整体重写 zip

该函数是幂等的：文档无损坏关系时几乎零开销，可安全地在每次 Document() 打开前调用。
"""
import os
import shutil
import posixpath
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_TAG = f"{{{REL_NS}}}Relationship"


def _rels_base_dir(rels_name: str) -> str:
    """
    由 .rels 文件路径推断其关系 Target 的解析基准目录。

    规则：<dir>/_rels/<part>.rels 中的 Target 相对于 <dir>。
      word/_rels/document.xml.rels        → word
      word/_rels/header1.xml.rels         → word
      _rels/.rels                         → ""（包根）
    """
    parent = posixpath.dirname(rels_name)          # 如 word/_rels 或 _rels
    if posixpath.basename(parent) == "_rels":
        return posixpath.dirname(parent)           # 去掉 _rels 这一层
    return parent


def _resolve_target(base_dir: str, target: str) -> str:
    """把关系 Target（相对路径）归一化为 zip 内部件路径。"""
    joined = posixpath.join(base_dir, target) if base_dir else target
    # 归一化 ../ 与 ./
    return posixpath.normpath(joined).replace("\\", "/")


def repair_dangling_rels(docx_path: str, verbose: bool = True) -> int:
    """
    清理 docx 中所有指向缺失内部部件的悬空关系。

    返回被删除的关系条数（0 表示无需修复）。
    仅当确有改动时才重写文件。
    """
    abs_path = os.path.abspath(docx_path)

    # ── 第一遍：只读扫描，找出需要修改的 .rels 及其新内容 ──────────────
    try:
        with ZipFile(abs_path, "r") as zf:
            all_names = set(zf.namelist())
            rels_names = [n for n in all_names if n.endswith(".rels")]

            patched: dict[str, bytes] = {}   # rels_name -> 新的 XML bytes
            removed_total = 0

            for rels_name in rels_names:
                try:
                    raw = zf.read(rels_name)
                    root = etree.fromstring(raw)
                except Exception:
                    continue  # 解析失败的 .rels 跳过，不影响其它

                base_dir = _rels_base_dir(rels_name)
                to_remove = []

                for rel in root.findall(_REL_TAG):
                    mode = (rel.get("TargetMode") or "Internal").strip()
                    if mode.lower() == "external":
                        continue  # 外部链接目标本就不在包内，保留

                    target = (rel.get("Target") or "").strip()
                    if not target:
                        to_remove.append(rel)
                        continue

                    resolved = _resolve_target(base_dir, target)
                    if resolved not in all_names:
                        to_remove.append(rel)

                if to_remove:
                    for rel in to_remove:
                        if verbose:
                            rid = rel.get("Id", "?")
                            tgt = rel.get("Target", "?")
                            print(f"    [关系修复] 删除悬空关系 {rid} → '{tgt}' （{rels_name}）")
                        root.remove(rel)
                    removed_total += len(to_remove)
                    patched[rels_name] = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    except Exception as e:
        if verbose:
            print(f"    [关系修复] 扫描失败（跳过）: {e}")
        return 0

    if not patched:
        return 0

    # ── 第二遍：重写 zip，仅替换被修补的 .rels ────────────────────────
    tmp_path = abs_path + ".repair.tmp"
    try:
        with ZipFile(abs_path, "r") as zf_in, \
             ZipFile(tmp_path, "w", ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                name = item.filename
                data = patched.get(name)
                if data is None:
                    data = zf_in.read(name)
                zf_out.writestr(item, data)
        shutil.move(tmp_path, abs_path)
    except Exception as e:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if verbose:
            print(f"    [关系修复] 重写失败（跳过）: {e}")
        return 0

    if verbose:
        print(f"    [关系修复] 已清理 {removed_total} 条悬空关系")
    return removed_total


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python docx_repair.py <docx路径>")
        sys.exit(1)
    n = repair_dangling_rels(sys.argv[1])
    print(f"清理关系数: {n}")
