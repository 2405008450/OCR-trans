#!/usr/bin/env python
"""配置迁移工具 (Configuration Migration Tool)

将旧版单配置格式迁移到新版双配置格式。

使用方法:
    python -m src.tools.migrate_config old_config.yaml new_config.yaml
    
或者:
    python src/tools/migrate_config.py old_config.yaml new_config.yaml
"""

import argparse
import sys
import yaml
from pathlib import Path
from typing import Dict, Any

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.exceptions import ConfigMigrationError


def load_yaml(path: str) -> Dict[str, Any]:
    """加载 YAML 配置文件
    
    Args:
        path: 配置文件路径
        
    Returns:
        配置字典
        
    Raises:
        ConfigMigrationError: 如果文件无法读取或解析
    """
    try:
        with open(path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            return config if config else {}
    except FileNotFoundError:
        raise ConfigMigrationError(
            "legacy", "dual-config",
            f"配置文件不存在: {path}"
        )
    except yaml.YAMLError as e:
        raise ConfigMigrationError(
            "legacy", "dual-config",
            f"YAML 解析失败: {e}"
        )
    except Exception as e:
        raise ConfigMigrationError(
            "legacy", "dual-config",
            f"读取文件失败: {e}"
        )


def save_yaml(config: Dict[str, Any], path: str) -> None:
    """保存配置到 YAML 文件
    
    Args:
        config: 配置字典
        path: 输出文件路径
        
    Raises:
        ConfigMigrationError: 如果文件无法写入
    """
    try:
        # 确保父目录存在
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        
        with open(path, 'w', encoding='utf-8') as f:
            yaml.dump(
                config,
                f,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,
                indent=2
            )
    except Exception as e:
        raise ConfigMigrationError(
            "legacy", "dual-config",
            f"写入文件失败: {e}"
        )


def is_legacy_format(config: Dict[str, Any]) -> bool:
    """检查是否是旧版配置格式
    
    Args:
        config: 配置字典
        
    Returns:
        True 如果是旧版格式
    """
    # 新格式包含 document_orientation 字段
    return 'document_orientation' not in config


def migrate_config(
    old_config: Dict[str, Any],
    target_orientation: str = "horizontal"
) -> Dict[str, Any]:
    """将旧版配置迁移到新版格式
    
    Args:
        old_config: 旧版配置字典
        target_orientation: 目标方向（"vertical" 或 "horizontal"）
        
    Returns:
        新版配置字典
        
    Raises:
        ConfigMigrationError: 如果迁移失败
    """
    if not is_legacy_format(old_config):
        raise ConfigMigrationError(
            "legacy", "dual-config",
            "配置已经是新版格式，无需迁移"
        )
    
    if target_orientation not in ["vertical", "horizontal"]:
        raise ConfigMigrationError(
            "legacy", "dual-config",
            f"无效的目标方向: {target_orientation}，必须是 'vertical' 或 'horizontal'"
        )
    
    # 创建新版配置结构
    new_config = {
        "document_orientation": target_orientation
    }
    
    # 分离全局配置和方向特定配置
    global_config = {}
    orientation_config = {}
    
    # 全局配置字段（两种方向共享）
    global_fields = ['logging', 'performance']
    
    # 方向特定配置字段
    orientation_fields = ['api', 'ocr', 'icon_detection', 'rendering', 'quality', 'image', 'translation']
    
    for key, value in old_config.items():
        if key in global_fields:
            global_config[key] = value
        elif key in orientation_fields:
            orientation_config[key] = value
        else:
            # 未知字段，放入全局配置
            global_config[key] = value
    
    # 添加全局配置
    new_config.update(global_config)
    
    # 创建两个配置区块（竖版和横版）
    # 默认情况下，两个区块使用相同的配置
    new_config['vertical_config'] = orientation_config.copy()
    new_config['horizontal_config'] = orientation_config.copy()
    
    # 添加注释说明
    new_config['_migration_note'] = (
        f"此配置由旧版格式自动迁移而来，默认方向为 {target_orientation}。"
        "建议根据实际需求调整 vertical_config 和 horizontal_config 的参数。"
    )
    
    return new_config


def add_comments_to_config(config_path: str) -> None:
    """为配置文件添加注释说明
    
    Args:
        config_path: 配置文件路径
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 在文件开头添加注释
        header = """# 双配置架构配置文件
# ==========================================
# 本配置文件由旧版格式自动迁移而来
#
# 文档方向说明:
#   - vertical: 竖版文档（如竖版营业执照）
#   - horizontal: 横版文档（如横版营业执照）
#
# 配置区块说明:
#   - vertical_config: 竖版文档专用配置
#   - horizontal_config: 横版文档专用配置
#   - 全局配置: logging, performance 等字段在两种方向下共享
#
# 使用方法:
#   1. 修改 document_orientation 字段选择文档方向
#   2. 根据需要调整对应配置区块的参数
#   3. 竖版和横版配置完全独立，互不影响
#
# ==========================================

"""
        
        # 写回文件
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(header + content)
            
    except Exception as e:
        # 添加注释失败不影响迁移结果
        print(f"警告: 无法添加注释: {e}", file=sys.stderr)


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器
    
    Returns:
        ArgumentParser 实例
    """
    parser = argparse.ArgumentParser(
        prog='migrate_config',
        description='将旧版单配置格式迁移到新版双配置格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  迁移配置（默认为横版）:
    python -m src.tools.migrate_config old.yaml new.yaml
    
  迁移为竖版配置:
    python -m src.tools.migrate_config old.yaml new.yaml --orientation vertical
    
  检查配置格式:
    python -m src.tools.migrate_config old.yaml --check-only

注意:
  - 迁移后的配置包含 vertical_config 和 horizontal_config 两个区块
  - 默认情况下，两个区块使用相同的配置（从旧配置复制）
  - 建议根据实际需求调整两个区块的参数
  - 旧配置文件不会被修改
"""
    )
    
    parser.add_argument(
        'input',
        type=str,
        help='输入配置文件路径（旧版格式）'
    )
    
    parser.add_argument(
        'output',
        type=str,
        nargs='?',
        default=None,
        help='输出配置文件路径（新版格式）'
    )
    
    parser.add_argument(
        '--orientation',
        type=str,
        choices=['vertical', 'horizontal'],
        default='horizontal',
        help='目标文档方向（默认: horizontal）'
    )
    
    parser.add_argument(
        '--check-only',
        action='store_true',
        help='仅检查配置格式，不执行迁移'
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='强制覆盖已存在的输出文件'
    )
    
    parser.add_argument(
        '--no-comments',
        action='store_true',
        help='不添加注释说明'
    )
    
    return parser


def main(argv=None):
    """主函数
    
    Args:
        argv: 命令行参数列表（用于测试）
        
    Returns:
        退出码（0 表示成功）
    """
    parser = create_parser()
    args = parser.parse_args(argv)
    
    try:
        # 加载旧配置
        print(f"加载配置文件: {args.input}")
        old_config = load_yaml(args.input)
        
        # 检查格式
        if is_legacy_format(old_config):
            print("✓ 检测到旧版配置格式")
        else:
            print("✓ 配置已经是新版格式")
            if args.check_only:
                return 0
            else:
                print("无需迁移")
                return 0
        
        # 如果只是检查，到此结束
        if args.check_only:
            print("\n配置格式检查完成")
            return 0
        
        # 检查输出路径
        if args.output is None:
            print("错误: 需要指定输出文件路径", file=sys.stderr)
            return 1
        
        # 检查输出文件是否已存在
        if Path(args.output).exists() and not args.force:
            print(f"错误: 输出文件已存在: {args.output}", file=sys.stderr)
            print("使用 --force 参数强制覆盖", file=sys.stderr)
            return 1
        
        # 执行迁移
        print(f"\n开始迁移配置...")
        print(f"  目标方向: {args.orientation}")
        new_config = migrate_config(old_config, args.orientation)
        
        # 保存新配置
        print(f"  保存到: {args.output}")
        save_yaml(new_config, args.output)
        
        # 添加注释
        if not args.no_comments:
            print(f"  添加注释说明...")
            add_comments_to_config(args.output)
        
        print("\n✓ 配置迁移成功!")
        print(f"\n新配置文件: {args.output}")
        print(f"文档方向: {args.orientation}")
        
        print("\n后续步骤:")
        print("1. 检查新配置文件，确认参数正确")
        print("2. 根据需要调整 vertical_config 和 horizontal_config 的参数")
        print("3. 使用 --config 参数指定新配置文件进行测试")
        
        return 0
        
    except ConfigMigrationError as e:
        print(f"迁移失败: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"未知错误: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
