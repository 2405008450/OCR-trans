"""工具模块 (Tools Module)

本模块包含各种实用工具，如配置迁移工具等。
"""

from .migrate_config import migrate_config, is_legacy_format

__all__ = ['migrate_config', 'is_legacy_format']
