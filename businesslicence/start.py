"""图片翻译系统启动文件

等同于命令:
$env:DEEPSEEK_API_KEY = "sk-a681d5bc9d3f4278b084ad4606251298"
python -m src.cli.main picture/test_image.png -o translated/test_result.png --verbose

直接运行: python start.py
右键运行: 右键点击此文件 -> 运行Python文件
"""

import os
import sys
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# 获取脚本所在目录（项目根目录）
SCRIPT_DIR = Path(__file__).parent.absolute()

# 确保项目根目录在 Python 路径中（对于右键运行很重要）
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# 加载项目根目录的 .env 统一配置
_project_root = SCRIPT_DIR.parent  # businesslicence 的上级即项目根
load_dotenv(_project_root / ".env")

# ==================== 配置区域 ====================
# 在这里修改你的配置

# API 密钥（从项目根 .env 统一管理）
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")

# 输入图片路径（使用绝对路径，无论从哪里运行都能找到文件）
INPUT_IMAGE = str(SCRIPT_DIR / r"C:\Users\Administrator\Desktop\picture\zhuhai.jpg")

# 输出目录（翻译后的图片会保存到这里）
OUTPUT_DIR = r"D:\it\xsfy\translated"

# 交互式验证模式
USE_GUI_MODE = True         # 使用GUI模式进行交互式验证（推荐，避免键盘输入问题）

# 其他选项
VERBOSE = True              # 详细输出
# 配置文件路径：支持竖版和横版配置
# - config/vertical.yaml: 竖版文档配置（营业执照等）
# - config/horizontal.yaml: 横版文档配置（标准参数）
# - config/horizontal_optimized.yaml: 横版文档配置（优化参数，推荐）
# - config/default.yaml: 默认配置
# - config/debug_merge.yaml: 旧格式配置（向后兼容）
# - config/auto.yaml: 自动识别方向配置（推荐，已启用交互式验证）
CONFIG_FILE = "config/auto.yaml"  # 使用自动识别配置（已启用交互式验证）
SOURCE_LANG = "zh"          # 源语言（中文）
TARGET_LANG = "en"          # 目标语言（英文）

# ==================================================


def main():
    """主函数"""
    # 声明使用全局变量
    global USE_GUI_MODE
    
    # 如果启用GUI模式，设置GUI输入处理
    if USE_GUI_MODE:
        try:
            import tkinter as tk
            from tkinter import messagebox, simpledialog
            import re
            
            # 创建全局root窗口
            root = tk.Tk()
            root.withdraw()
            
            # 定义GUI对话框函数
            def show_verification_dialog(region_info):
                """显示验证对话框"""
                dialog = tk.Toplevel(root)
                dialog.title(f"印章文字验证 ({region_info['index']}/{region_info['total']})")
                dialog.geometry("500x300")
                dialog.resizable(False, False)
                dialog.attributes('-topmost', True)
                dialog.focus_force()
                
                result = {'action': None, 'text': None}
                
                # 标题
                title_label = tk.Label(dialog, text=f"检测到{region_info['type_name']}", 
                                      font=("Arial", 14, "bold"), fg="blue")
                title_label.pack(pady=10)
                
                # 信息框
                info_frame = tk.Frame(dialog, relief=tk.SUNKEN, borderwidth=2)
                info_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
                
                info_text = f"""
类型: {region_info['text_type']}
位置: {region_info['bbox']}
置信度: {region_info['confidence']:.2f}

识别内容:
{region_info['text']}
                """
                
                info_label = tk.Label(info_frame, text=info_text, font=("Courier New", 10),
                                     justify=tk.LEFT, anchor=tk.W)
                info_label.pack(padx=10, pady=10)
                
                # 按钮框
                button_frame = tk.Frame(dialog)
                button_frame.pack(pady=10)
                
                def on_confirm():
                    result['action'] = 'confirm'
                    dialog.destroy()
                
                def on_correct():
                    corrected = simpledialog.askstring("修正内容", "请输入正确的文字内容:",
                                                      parent=dialog, initialvalue=region_info['text'])
                    if corrected:
                        result['action'] = 'correct'
                        result['text'] = corrected
                        dialog.destroy()
                
                def on_skip():
                    result['action'] = 'skip'
                    dialog.destroy()
                
                # 按钮
                tk.Button(button_frame, text="确认正确 (Y)", command=on_confirm, width=15, height=2,
                         bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
                tk.Button(button_frame, text="修正内容 (N)", command=on_correct, width=15, height=2,
                         bg="#FF9800", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
                tk.Button(button_frame, text="跳过 (S)", command=on_skip, width=15, height=2,
                         bg="#9E9E9E", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
                
                # 键盘快捷键
                dialog.bind('y', lambda e: on_confirm())
                dialog.bind('Y', lambda e: on_confirm())
                dialog.bind('n', lambda e: on_correct())
                dialog.bind('N', lambda e: on_correct())
                dialog.bind('s', lambda e: on_skip())
                dialog.bind('S', lambda e: on_skip())
                dialog.bind('<Return>', lambda e: on_confirm())
                dialog.bind('<Escape>', lambda e: on_skip())
                
                dialog.wait_window()
                return result['action'], result['text']
            
            # 输入捕获类
            class InteractiveInputCapture:
                def __init__(self):
                    self.current_region = {}
                    self.buffer = []
                
                def parse_region_info(self, lines):
                    info = {}
                    for line in lines:
                        match = re.search(r'检测到(印章内文字|被覆盖的日期)\s*\(#(\d+)/(\d+)\)', line)
                        if match:
                            info['type_name'] = match.group(1)
                            info['text_type'] = 'seal_inner' if match.group(1) == '印章内文字' else 'seal_overlap'
                            info['index'] = int(match.group(2))
                            info['total'] = int(match.group(3))
                        
                        match = re.search(r'位置:\s*\(([^)]+)\)', line)
                        if match:
                            bbox_str = match.group(1)
                            bbox_parts = [int(float(x.strip())) for x in bbox_str.split(',')]
                            if len(bbox_parts) == 4:
                                info['bbox'] = tuple(bbox_parts)
                        
                        match = re.search(r'识别内容:\s*"([^"]*)"', line)
                        if match:
                            info['text'] = match.group(1)
                        
                        match = re.search(r'置信度:\s*([\d.]+)', line)
                        if match:
                            info['confidence'] = float(match.group(1))
                    return info
                
                def custom_input(self, prompt):
                    self.buffer.append(prompt)
                    
                    if "识别是否正确" in prompt:
                        region_info = self.parse_region_info(self.buffer)
                        self.buffer = []
                        
                        if region_info:
                            action, text = show_verification_dialog(region_info)
                            
                            if action == 'confirm':
                                print("✅ 确认识别正确，继续处理")
                                return 'y'
                            elif action == 'correct':
                                print(f"✏️ 用户选择修正内容: \"{text}\"")
                                self.current_region['corrected_text'] = text
                                return 'n'
                            elif action == 'skip':
                                print("⏭️ 跳过此区域")
                                return 's'
                        return 'y'
                    
                    elif "请输入正确的文字内容" in prompt:
                        corrected = self.current_region.get('corrected_text', '')
                        print(f"✅ 使用修正后的内容: \"{corrected}\"")
                        return corrected
                    
                    return ''
            
            # 替换input函数
            import builtins
            original_input = builtins.input
            input_capture = InteractiveInputCapture()
            builtins.input = input_capture.custom_input
            
            print("=" * 60)
            print("GUI模式已启用 - 交互式验证将使用图形界面")
            print("=" * 60)
            print()
            
        except ImportError:
            print("=" * 60)
            print("警告: 无法导入Tkinter，将使用终端模式")
            print("=" * 60)
            print()
            USE_GUI_MODE = False
    
    # 检查命令行参数
    if len(sys.argv) > 1:
        # 如果提供了命令行参数，使用第一个参数作为输入图片路径
        INPUT_IMAGE_OVERRIDE = sys.argv[1]
        # 检查是否是绝对路径
        if not Path(INPUT_IMAGE_OVERRIDE).is_absolute():
            # 如果是相对路径，转换为绝对路径
            INPUT_IMAGE_OVERRIDE = str(SCRIPT_DIR / INPUT_IMAGE_OVERRIDE)
    else:
        INPUT_IMAGE_OVERRIDE = INPUT_IMAGE
    
    # 生成输出文件路径：原文件名_translated_时间戳.扩展名
    input_path = Path(INPUT_IMAGE_OVERRIDE)    
    # 生成时间戳（格式：YYYYMMDD_HHMM）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    
    # 构建输出文件名：原文件名_translated_时间戳.扩展名
    output_filename = f"{input_path.stem}_translated_{timestamp}{input_path.suffix}"
    OUTPUT_IMAGE = str(Path(OUTPUT_DIR) / output_filename)
    
    # 确保输出目录存在
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    
    # 设置环境变量
    os.environ['DEEPSEEK_API_KEY'] = DEEPSEEK_API_KEY
    
    # 初始化 ConfigManager
    try:
        from src.config.config_manager import ConfigManager
        
        # 如果指定了配置文件，使用指定的配置文件；否则使用默认配置
        config_path = str(SCRIPT_DIR / CONFIG_FILE) if CONFIG_FILE else None
        config = ConfigManager(config_path=config_path)
        
        # 验证配置
        errors = config.validate()
        if errors:
            print("配置验证失败 | Configuration validation failed:")
            for error in errors:
                print(f"  - {error}")
            input("按回车键退出...")
            sys.exit(1)
        
    except Exception as e:
        print(f"错误: 无法初始化配置管理器: {e}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")
        sys.exit(1)
    
    # 打印启动信息
    print("=" * 60)
    print("图片翻译系统 - 双配置架构支持")
    print("Image Translation System - Dual Configuration Architecture")
    print("=" * 60)
    print(f"Python: {sys.executable}")
    print(f"工作目录 | Working Directory: {os.getcwd()}")
    print(f"脚本目录 | Script Directory: {SCRIPT_DIR}")
    print(f"输入图片 | Input Image: {INPUT_IMAGE_OVERRIDE}")
    print(f"输出图片 | Output Image: {OUTPUT_IMAGE}")
    print(f"配置文件 | Config File: {CONFIG_FILE if CONFIG_FILE else 'default'}")
    print(f"文档方向 | Document Orientation: {config.get_orientation()}")
    print(f"配置格式 | Config Format: {'Legacy' if config.is_legacy_format() else 'Dual-Config'}")
    print(f"源语言 | Source Language: {SOURCE_LANG}")
    print(f"目标语言 | Target Language: {TARGET_LANG}")
    print(f"详细输出 | Verbose: {'是 | Yes' if VERBOSE else '否 | No'}")
    print("=" * 60)
    print()
    
    # 检查输入文件
    if not Path(INPUT_IMAGE_OVERRIDE).exists():
        print(f"错误: 输入文件不存在: {INPUT_IMAGE_OVERRIDE}")
        input("按回车键退出...")
        sys.exit(1)
    
    # 导入所需模块
    try:
        from src.pipeline.translation_pipeline import TranslationPipeline
    except ImportError as e:
        print(f"错误: 无法导入模块: {e}")
        print(f"请确保在项目根目录运行此脚本")
        input("按回车键退出...")
        sys.exit(1)
    
    # 初始化 TranslationPipeline 并传递 ConfigManager
    try:
        pipeline = TranslationPipeline(config)
        
        # 如果启用了GUI模式，设置GUI回调函数
        if USE_GUI_MODE:
            try:
                # show_verification_dialog 应该在上面的 if USE_GUI_MODE 块中定义
                pipeline._gui_verification_callback = show_verification_dialog
                print("GUI回调函数已设置")
            except NameError:
                print("警告: GUI回调函数未定义，将使用终端模式")
        
        # 执行翻译
        print(f"开始翻译图片... | Starting translation...")
        result = pipeline.translate_image(
            INPUT_IMAGE_OVERRIDE,
            OUTPUT_IMAGE,
            source_lang=SOURCE_LANG,
            target_lang=TARGET_LANG
        )
        
        # 检查输出文件
        if Path(OUTPUT_IMAGE).exists():
            print()
            print("=" * 60)
            print(f"翻译完成！输出文件: {OUTPUT_IMAGE}")
            print("=" * 60)
            sys.exit(0)
        else:
            print()
            print("=" * 60)
            print(f"翻译可能失败，未找到输出文件")
            print("=" * 60)
            sys.exit(1)
        
    except Exception as e:
        print()
        print("=" * 60)
        print(f"错误: {e}")
        print("=" * 60)
        import traceback
        traceback.print_exc()
        input("按回车键退出...")
        sys.exit(1)
    
    finally:
        # 清理GUI资源
        if USE_GUI_MODE:
            try:
                import builtins
                if hasattr(builtins, 'input') and hasattr(builtins.input, '__self__'):
                    # 恢复原始input函数
                    builtins.input = original_input
                # 关闭root窗口
                if 'root' in locals():
                    root.destroy()
            except:
                pass


if __name__ == '__main__':
    main()
