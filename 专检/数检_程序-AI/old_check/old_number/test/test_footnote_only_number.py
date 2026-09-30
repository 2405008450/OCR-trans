"""
专门测试脚注检查和替换
"""
import sys
from pathlib import Path

from 数检.backup_copy.backup_manager import ensure_backup_copy

sys.path.insert(0, str(Path(__file__).parent))

from replace.word.footnote_replacer import check_text_in_footnotes, replace_in_footnotes_xml


def test_footnote(doc_path):
    """测试脚注功能"""
    print("=" * 80)
    print("🔍 脚注检查和替换测试")
    print("=" * 80)
    
    doc_path = str(Path(doc_path).resolve())
    print(f"\n📂 文档: {Path(doc_path).name}")
    print(f"📍 完整路径: {doc_path}\n")
    
    # 测试文本
    test_texts = [
        "严控标签",
        "严控标签：",
        "严控标签：一般为已经停止服务",
        "严控标签：一般为已经停止服务/更新、存在严重安全漏洞或者不再推荐使用的开源软件或某个版本，由管理方式添加严控标识。"
    ]
    
    print("🔍 检查文本是否在脚注中:")
    print("-" * 80)
    
    for i, text in enumerate(test_texts, 1):
        try:
            found = check_text_in_footnotes(doc_path, text)
            status = "✅ 找到" if found else "❌ 未找到"
            print(f"{i}. {status} '{text[:50]}...'")
        except Exception as e:
            print(f"{i}. ⚠️ 检查失败 '{text[:50]}...'")
            print(f"   错误: {e}")
    backup_path = ensure_backup_copy(doc_path, suffix="quicktest")
    print(f"✓ 备份: {backup_path}")
    # 尝试替换
    print("\n" + "=" * 80)
    print("🔄 尝试替换测试:")
    print("=" * 80)
    
    old_text = "严控标签：一般为已经停止服务/更新、存在严重安全漏洞或者不再推荐使用的开源软件或某个版本，由管理方式添加严控标识。"
    new_text = "Strict control label: Generally refers to open source software or a specific version that has ceased service/updates, has serious security vulnerabilities, or is no longer recommended for use, with a strict control identifier added by the manager."
    
    print(f"\n原文: {old_text[:60]}...")
    print(f"译文: {new_text[:60]}...")
    
    try:
        # 先检查
        found = check_text_in_footnotes(doc_path, old_text)
        print(f"\n检查结果: {'✅ 找到' if found else '❌ 未找到'}")
        
        if found:
            # 尝试替换
            print("\n执行替换...")
            success = replace_in_footnotes_xml(backup_path, old_text, new_text, "测试替换")
            print(f"替换结果: {'✅ 成功' if success else '❌ 失败'}")
        else:
            print("\n⚠️ 文本不在脚注中，无法替换")
            
    except Exception as e:
        print(f"\n⚠️ 操作失败: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 80)


if __name__ == "__main__":
    DOC_PATH = r"../测试文件/译文-B260328127-Y-中国银行开源软件管理指引.docx"
    
    if len(sys.argv) > 1:
        DOC_PATH = sys.argv[1]
    
    if not Path(DOC_PATH).exists():
        print(f"❌ 文档不存在: {DOC_PATH}")
    else:
        test_footnote(DOC_PATH)
