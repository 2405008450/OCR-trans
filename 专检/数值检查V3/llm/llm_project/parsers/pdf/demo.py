import argparse
import sys
from pathlib import Path

# 添加父目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from to_pdf import CompletePDFExtractor, ContentFormatter

def parse_pdf(pdf_path: str, mode: str = "clean"):
    """仅执行 PDF 解析"""
    pdf_file = Path(pdf_path)

    if not pdf_file.exists():
        print(f"❌ 错误: 文件 {pdf_file} 不存在")
        return

    print(f"🔍 正在解析 PDF: {pdf_file.name} ...")

    try:
        # 直接实例化提取器，跳过 Word 转换步骤
        extractor = CompletePDFExtractor(str(pdf_file))
        pages = extractor.extract_all(verbose=True)

        # 格式化输出
        text_output = ContentFormatter.format_to_text(pages, mode=mode)

        print(f"✅ 解析成功！共处理 {len(pages)} 页。")
        # print("-" * 30 + " 预览 " + "-" * 30)
        # print(text_output)  # 打印前500个字符预览
        # print("-" * 66)
        return text_output

    except Exception as e:
        print(f"❌ 解析过程发生错误: {e}")
        return None  # 发生异常时也返回 None，防止后续程序崩溃


if __name__ == "__main__":
    pdf_path=r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查1\测试文件\译文-数值检查测试文件1(1).pdf"
    text=parse_pdf(pdf_path)
    print(text)
    parser = argparse.ArgumentParser(description="纯 PDF 解析工具")
    parser.add_argument("pdf_path", help="PDF 文件路径")
    parser.add_argument("--mode", choices=["clean", "structured"], default="clean", help="提取模式")

    args = parser.parse_args()
    parse_pdf(args.pdf_path, args.mode)