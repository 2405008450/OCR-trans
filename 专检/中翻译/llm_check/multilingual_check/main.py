import os
import sys
import importlib
from pathlib import Path
from register import registry
# from demo1 import run_fixes_and_report
from router import ai_router
from zhongfanyi.llm.llm_project.parsers.excel.excel_parser import parse_excel_with_pandas
from zhongfanyi.llm.llm_project.parsers.pptx.pptx_parser import parse_pptx
from zhongfanyi.llm.llm_project.parsers.txt.txt_parser import parse_txt
from zhongfanyi.llm.llm_project.parsers.word.body_extractor import extract_body_text
from zhongfanyi.llm.llm_project.parsers.pdf.pdf_parser import parse_pdf
from zhongfanyi.llm.llm_project.parsers.json.json_parser import jsonParser

# 自动加载 tools 目录下的所有工具
def auto_load_tools():
    """自动扫描并导入 tools 目录下的所有 Python 模块"""
    tools_dir = Path(__file__).parent / "tools"
    
    if not tools_dir.exists():
        print(f"[警告] tools 目录不存在: {tools_dir}")
        return
    
    # 遍历 tools 目录下的所有 .py 文件
    for tool_file in tools_dir.glob("*.py"):
        if tool_file.name.startswith("_"):
            continue
        
        module_name = f"tools.{tool_file.stem}"
        try:
            importlib.import_module(module_name)
            print(f"[✓] 已加载工具模块: {module_name}")
        except Exception as e:
            print(f"[✗] 加载工具模块失败 {module_name}: {e}")

def parse_file(file_path: str) -> str:
    """根据文件类型解析文件内容"""
    file_path = Path(file_path)
    
    if not file_path.exists():
        return f"[错误] 文件不存在: {file_path}"
    
    suffix = file_path.suffix.lower()
    
    try:
        # TXT 文件
        if suffix == ".txt":
            print(f"[解析] 正在解析 TXT 文件...")
            return parse_txt(str(file_path), mode="clean")
        
        # PDF 文件
        elif suffix == ".pdf":
            print(f"[解析] 正在解析 PDF 文件...")
            return parse_pdf(str(file_path), mode="clean")
        
        # Word 文件
        elif suffix in [".docx", ".doc"]:
            print(f"[解析] 正在解析 Word 文件...")
            return extract_body_text(str(file_path))
        
        # Excel 文件
        elif suffix in [".xlsx", ".xls"]:
            print(f"[解析] 正在解析 Excel 文件...")
            df = parse_excel_with_pandas(str(file_path))
            return df.to_string() if df is not None else "[错误] Excel 解析失败"

        # ppt 文件
        elif suffix in [".pptx"]:
            print(f"[解析] 正在解析 ppt 文件...")
            df = parse_pptx(str(file_path))
            return df

        # # html 文件
        # elif suffix in [".html"]:
        #     print(f"[解析] 正在解析 html 文件...")
        #     return parse_html(str(file_path))
        
        # JSON 文件
        elif suffix == ".json":
            print(f"[解析] 正在解析 JSON 文件...")
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                content = f.read()
            parsed = jsonParser.load_error_json_text(content)
            return str(parsed)
        
        else:
            return f"[错误] 不支持的文件类型: {suffix}\n支持的格式: .pptx, .txt, .pdf, .docx, .doc, .xlsx, .xls, .json"
    
    except Exception as e:
        return f"[错误] 文件解析失败: {e}"

def main():
    """主函数：处理用户上传的文件并进行标点检查"""
    print("=" * 60)
    print("多语种标点检查系统")
    print("=" * 60)
    
    # 自动加载所有工具
    auto_load_tools()
    
    print(f"\n已注册 {len(registry.tools_metadata)} 个工具:")
    for tool in registry.tools_metadata:
        print(f"  - {tool['name']}: {tool['description']}")
    
    print("\n" + "=" * 60)
    print("使用说明:")
    print("  1. 输入文件路径进行检查（支持 .xlsx, .xls, .pptx, .txt, .pdf, .docx, .xlsx, .json）")
    print("  2. 直接输入文本进行检查（只支持多行分行检查，输入空行结束）")
    print("  3. 输入 'quit' 退出")
    print("=" * 60)
    
    while True:
        first_line = input("\n> ").strip()
        
        if first_line.lower() in ['quit', 'exit', 'q']:
            print("再见！")
            break
        
        if not first_line:
            continue
        
        # 收集多行输入：如果用户粘贴了多行文本，继续读取直到空行
        lines = [first_line]
        import msvcrt
        import time
        time.sleep(0.1)  # 等待粘贴缓冲区
        while msvcrt.kbhit():
            extra_line = input().strip()
            if extra_line:
                lines.append(extra_line)
            else:
                break
        
        user_input = "\n".join(lines)
        
        # 判断是文件路径还是文本
        file_path = Path(user_input)
        if file_path.exists() and file_path.is_file():
            # 是文件，先解析文件内容
            print(f"\n[文件检测] 检测到文件: {file_path.name}")
            file_content = parse_file(str(file_path))
            
            if file_content.startswith("[错误]"):
                print(file_content)
                continue
            
            print(f"[文件内容] 已提取 {len(file_content)} 字符")
            print(f"[内容预览] {file_content}...")
            
            # 将文件内容作为检查对象
            check_text = file_content
        else:
            # 是普通文本
            check_text = user_input
        
        # 使用 AI 路由决策调用哪个工具
        decision = ai_router(check_text, registry.tools_metadata)
        tool_name = decision.get("tool_name")
        arguments = decision.get("arguments", {})
        
        # 确保 prompt 参数包含要检查的文本
        if "prompt" not in arguments:
            arguments["prompt"] = check_text
        
        print(f"\n[AI 决策] 工具: {tool_name}")
        print(f"[AI 决策] 参数: {arguments.get('prompt', '')[:100]}...")
        
        # 执行对应的工具
        if tool_name in registry.handlers:
            try:
                result = registry.handlers[tool_name](arguments)
                print(result)
            except Exception as e:
                print(f"\n[错误] 工具执行失败: {e}")
        else:
            print(f"\n[错误] 未找到工具: {tool_name}")

if __name__ == "__main__":
    main()
