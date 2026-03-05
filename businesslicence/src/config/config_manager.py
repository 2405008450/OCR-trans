"""配置管理器模块

本模块提供 ConfigManager 类，用于管理系统配置，包括从 YAML 文件和环境变量读取配置、
配置验证和模板生成。

ConfigManager 支持双配置架构，可以为竖版和横版文档方向提供不同的处理策略。

Configuration manager for image translation system.

This module provides the ConfigManager class for managing system configuration,
including reading from YAML files and environment variables, validation,
and template generation.

The ConfigManager supports dual-configuration architecture for vertical and horizontal
document orientations, allowing different processing strategies for different document types.
"""

import os
import re
import stat
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml
from dotenv import load_dotenv

from src.exceptions import ConfigError

# Set up logger
logger = logging.getLogger(__name__)


class ConfigManager:
    """Manages system configuration from YAML files and environment variables.
    
    The ConfigManager supports:
    - Reading configuration from YAML files
    - Loading environment variables from .env files
    - Environment variable substitution using ${VAR_NAME} syntax
    - Nested key access using dot notation (e.g., "api.timeout")
    - Configuration validation
    - Template file generation
    - Dual-configuration architecture (vertical/horizontal document orientations)
    - Backward compatibility with legacy single-configuration format
    
    Attributes:
        config_path: Path to the configuration file
        _config: Internal configuration dictionary
        _raw_config: Raw configuration loaded from file (before processing)
        _orientation: Current document orientation ("vertical" or "horizontal")
        _is_legacy: Whether the configuration uses legacy format
    """
    
    # Required configuration fields with their expected types
    REQUIRED_FIELDS = {
        'api.deepseek_key': str,
    }
    
    # Optional fields with default values
    DEFAULT_CONFIG = {
        'api': {
            'deepseek_key': '${DEEPSEEK_API_KEY}',
            'timeout': 10,
            'max_retries': 3,
            'retry_backoff': [1, 2, 4]
        },
        'ocr': {
            'engine': 'paddleocr',
            'confidence_threshold': 0.6,
            'min_text_area': 100,
            'merge_threshold': 10,
            'merge_threshold_horizontal': 10,  # 水平合并阈值（同一行）
            'merge_threshold_vertical': 3,     # 垂直合并阈值（不同行，设置较小避免多行合并）
            'font_size_diff_threshold': 0.2,   # 字体大小差异阈值（20%，只有字体相近才合并）
            'use_server_model': False          # 是否使用服务器版模型（更准确但更慢）
        },
        'icon_detection': {
            'aspect_ratio_threshold': 0.8,
            'complexity_threshold': 0.7,
            'text_density_threshold': 0.3,
            'whitelist': ['qrcode', 'seal', 'logo']
        },
        'rendering': {
            'font_family': 'Arial',
            'font_fallback': ['DejaVu Sans', 'Liberation Sans'],
            'enable_antialiasing': True,
            'enable_stroke': False,
            'stroke_width': 1,
            'background_sample_radius': 5,
            'use_inpainting': True
        },
        'quality': {
            'min_translation_coverage': 0.9,
            'check_artifacts': True,
            'generate_report': True
        },
        'performance': {
            'cache_ocr_results': True,
            'parallel_translation': True,
            'max_workers': 4
        }
    }
    
    # Validation rules: field -> (type, min_value, max_value) or (type, allowed_values)
    VALIDATION_RULES = {
        'api.timeout': (int, 1, 300),
        'api.max_retries': (int, 0, 10),
        'ocr.confidence_threshold': (float, 0.0, 1.0),
        'ocr.min_text_area': (int, 1, 10000),
        'ocr.merge_threshold': (int, 0, 100),
        'ocr.merge_threshold_horizontal': (int, 0, 100),
        'ocr.merge_threshold_vertical': (int, 0, 100),
        'ocr.font_size_diff_threshold': (float, 0.0, 1.0),
        'icon_detection.aspect_ratio_threshold': (float, 0.0, 1.0),
        'icon_detection.complexity_threshold': (float, 0.0, 1.0),
        'icon_detection.text_density_threshold': (float, 0.0, 1.0),
        'rendering.stroke_width': (int, 0, 20),
        'rendering.background_sample_radius': (int, 1, 50),
        'quality.min_translation_coverage': (float, 0.0, 1.0),
        'performance.max_workers': (int, 1, 32),
    }
    
    def __init__(self, config_path: Optional[str] = None, env_path: Optional[str] = None):
        """Initialize the configuration manager.
        
        Args:
            config_path: Path to YAML configuration file. If None, uses default config.
            env_path: Path to .env file. If None, looks for .env in current directory.
        """
        self.config_path = config_path
        self._config: Dict[str, Any] = {}
        self._raw_config: Dict[str, Any] = {}
        self._orientation: str = "horizontal"  # Default orientation
        self._is_legacy: bool = False
        
        # Load environment variables from .env file
        if env_path:
            load_dotenv(env_path)
        else:
            load_dotenv()
        
        # Start with default configuration
        self._config = self._deep_copy(self.DEFAULT_CONFIG)
        
        # Load and merge configuration from file if provided
        if config_path and Path(config_path).exists():
            self._raw_config = self._load_yaml(config_path)
            
            # Detect configuration format and set orientation
            self._detect_format_and_orientation()
            
            # Process configuration based on format
            if self._is_legacy:
                logger.warning(
                    "Legacy configuration format detected. "
                    "Consider upgrading to the new dual-configuration format. "
                    "The configuration will be mapped to horizontal_config."
                )
                # Map legacy config to horizontal_config
                file_config = self._raw_config
            else:
                # Extract orientation-specific configuration
                file_config = self._extract_orientation_config()
            
            self._merge_config(file_config)
        
        # Substitute environment variables
        self._substitute_env_vars(self._config)
    
    def _deep_copy(self, obj: Any) -> Any:
        """Create a deep copy of a nested dictionary/list structure."""
        if isinstance(obj, dict):
            return {k: self._deep_copy(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._deep_copy(item) for item in obj]
        else:
            return obj
    
    def _load_yaml(self, path: str) -> Dict[str, Any]:
        """Load configuration from a YAML file.
        
        Args:
            path: Path to the YAML file
            
        Returns:
            Dictionary containing the configuration
            
        Raises:
            ConfigError: If the file cannot be read or parsed
        """
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                return config if config else {}
        except yaml.YAMLError as e:
            raise ConfigError(f"Failed to parse YAML configuration: {e}")
        except IOError as e:
            raise ConfigError(f"Failed to read configuration file: {e}")
    
    def _merge_config(self, source: Dict[str, Any]) -> None:
        """Merge source configuration into the current configuration.
        
        Args:
            source: Configuration dictionary to merge
        """
        self._deep_merge(self._config, source)
    
    def _deep_merge(self, target: Dict[str, Any], source: Dict[str, Any]) -> None:
        """Recursively merge source dictionary into target dictionary.
        
        Args:
            target: Target dictionary to merge into
            source: Source dictionary to merge from
        """
        for key, value in source.items():
            if key in target and isinstance(target[key], dict) and isinstance(value, dict):
                self._deep_merge(target[key], value)
            else:
                target[key] = value
    
    def _substitute_env_vars(self, config: Any) -> None:
        """Substitute environment variables in configuration values.
        
        Supports ${VAR_NAME} syntax for environment variable substitution.
        
        Args:
            config: Configuration dictionary or value to process
        """
        if isinstance(config, dict):
            for key, value in config.items():
                if isinstance(value, str):
                    config[key] = self._substitute_string(value)
                elif isinstance(value, (dict, list)):
                    self._substitute_env_vars(value)
        elif isinstance(config, list):
            for i, item in enumerate(config):
                if isinstance(item, str):
                    config[i] = self._substitute_string(item)
                elif isinstance(item, (dict, list)):
                    self._substitute_env_vars(item)
    
    def _substitute_string(self, value: str) -> str:
        """Substitute environment variables in a string value.
        
        Args:
            value: String potentially containing ${VAR_NAME} patterns
            
        Returns:
            String with environment variables substituted
        """
        pattern = r'\$\{([^}]+)\}'
        
        def replace_var(match):
            var_name = match.group(1)
            env_value = os.environ.get(var_name, '')
            return env_value
        
        return re.sub(pattern, replace_var, value)
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value by key.
        
        Supports dot notation for nested keys (e.g., "api.timeout").
        
        Args:
            key: Configuration key, supports dot notation for nested access
            default: Default value if key is not found
            
        Returns:
            Configuration value or default
        """
        keys = key.split('.')
        value = self._config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any) -> None:
        """Set a configuration value by key.
        
        Supports dot notation for nested keys.
        
        Args:
            key: Configuration key, supports dot notation
            value: Value to set
        """
        keys = key.split('.')
        config = self._config
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value

    
    def validate(self) -> List[str]:
        """Validate the configuration for completeness and correctness.
        
        Checks:
        - Required fields are present and non-empty
        - Field values are of correct types
        - Numeric values are within valid ranges
        
        Returns:
            List of error messages. Empty list indicates valid configuration.
        """
        errors = []
        
        # Check required fields
        for field, expected_type in self.REQUIRED_FIELDS.items():
            value = self.get(field)
            if value is None:
                errors.append(f"Missing required field: {field}")
            elif not isinstance(value, expected_type):
                errors.append(f"Field '{field}' must be of type {expected_type.__name__}")
            elif isinstance(value, str) and not value.strip():
                errors.append(f"Required field '{field}' cannot be empty")
        
        # Validate field types and ranges
        for field, rule in self.VALIDATION_RULES.items():
            value = self.get(field)
            if value is None:
                continue  # Skip if not set (will use default)
            
            expected_type = rule[0]
            
            # Type check
            if not isinstance(value, expected_type):
                # Allow int for float fields
                if expected_type == float and isinstance(value, int):
                    pass
                else:
                    errors.append(
                        f"Field '{field}' must be of type {expected_type.__name__}, "
                        f"got {type(value).__name__}"
                    )
                    continue
            
            # Range check for numeric types
            if len(rule) == 3 and expected_type in (int, float):
                min_val, max_val = rule[1], rule[2]
                if not min_val <= value <= max_val:
                    errors.append(
                        f"Field '{field}' must be between {min_val} and {max_val}, "
                        f"got {value}"
                    )
        
        # Validate list fields
        list_fields = [
            'api.retry_backoff',
            'icon_detection.whitelist',
            'rendering.font_fallback'
        ]
        for field in list_fields:
            value = self.get(field)
            if value is not None and not isinstance(value, list):
                errors.append(f"Field '{field}' must be a list")
        
        # Validate boolean fields
        bool_fields = [
            'rendering.enable_antialiasing',
            'rendering.enable_stroke',
            'rendering.use_inpainting',
            'quality.check_artifacts',
            'quality.generate_report',
            'performance.cache_ocr_results',
            'performance.parallel_translation'
        ]
        for field in bool_fields:
            value = self.get(field)
            if value is not None and not isinstance(value, bool):
                errors.append(f"Field '{field}' must be a boolean")
        
        return errors
    
    def create_template(self, path: str) -> None:
        """Create a configuration template file with all options and comments.
        
        The template includes all available configuration options with
        explanatory comments and default values.
        
        Args:
            path: Path where the template file should be created
            
        Raises:
            ConfigError: If the file cannot be created
        """
        template_content = '''# Image Translation System Configuration
# =====================================
# This file contains all configuration options for the image translation system.
# Environment variables can be referenced using ${VAR_NAME} syntax.

# API Configuration
# -----------------
api:
  # DeepSeek API key for translation service
  # Can be set via environment variable: ${DEEPSEEK_API_KEY}
  deepseek_key: ${DEEPSEEK_API_KEY}
  
  # Request timeout in seconds (1-300)
  timeout: 10
  
  # Maximum number of retry attempts (0-10)
  max_retries: 3
  
  # Retry backoff intervals in seconds
  retry_backoff: [1, 2, 4]

# OCR Configuration
# -----------------
ocr:
  # OCR engine to use (currently only 'paddleocr' is supported)
  engine: paddleocr
  
  # Minimum confidence threshold for text detection (0.0-1.0)
  confidence_threshold: 0.6
  
  # Minimum text area in pixels to consider (1-10000)
  min_text_area: 100
  
  # Distance threshold for merging adjacent text regions (0-100)
  # 通用合并阈值（如果不设置水平/垂直阈值，使用此值）
  merge_threshold: 10
  
  # 水平合并阈值：用于合并同一行的文字 (0-100)
  # 设置较大的值可以合并同一行中间隔较远的文字
  merge_threshold_horizontal: 10
  
  # 垂直合并阈值：用于合并不同行的文字 (0-100)
  # 设置较小的值可以避免将多行文字合并成一行
  # 推荐值：3-5 像素
  merge_threshold_vertical: 3
  
  # 字体大小差异阈值 (0.0-1.0)
  # 只有字体大小相近的文字才会合并在一起翻译
  # 0.2 表示允许20%的字体大小差异
  # 设置更小的值（如0.1）会更严格地区分不同大小的文字
  font_size_diff_threshold: 0.2

# Icon Detection Configuration
# ----------------------------
icon_detection:
  # Aspect ratio threshold for icon detection (0.0-1.0)
  # Values closer to 1.0 indicate more square-like regions
  aspect_ratio_threshold: 0.8
  
  # Content complexity threshold (0.0-1.0)
  complexity_threshold: 0.7
  
  # Text density threshold (0.0-1.0)
  text_density_threshold: 0.3
  
  # List of icon types to always preserve
  whitelist:
    - qrcode
    - seal
    - logo

# Text Rendering Configuration
# ----------------------------
rendering:
  # Primary font family for rendered text
  font_family: Arial
  
  # Fallback fonts if primary is not available
  font_fallback:
    - DejaVu Sans
    - Liberation Sans
  
  # Enable anti-aliasing for smoother text
  enable_antialiasing: true
  
  # Enable text stroke/outline effect
  enable_stroke: false
  
  # Stroke width in pixels (0-20)
  stroke_width: 1
  
  # Radius for background color sampling (1-50)
  background_sample_radius: 5
  
  # Use inpainting for background restoration
  use_inpainting: true

# Quality Validation Configuration
# --------------------------------
quality:
  # Minimum acceptable translation coverage (0.0-1.0)
  min_translation_coverage: 0.9
  
  # Check for visual artifacts in output
  check_artifacts: true
  
  # Generate quality report after processing
  generate_report: true

# Performance Configuration
# -------------------------
performance:
  # Cache OCR results to avoid redundant processing
  cache_ocr_results: true
  
  # Enable parallel translation of text regions
  parallel_translation: true
  
  # Maximum number of parallel workers (1-32)
  max_workers: 4
'''
        
        try:
            # Ensure parent directory exists
            Path(path).parent.mkdir(parents=True, exist_ok=True)
            
            # Write template file
            with open(path, 'w', encoding='utf-8') as f:
                f.write(template_content)
            
            # Set file permissions to owner read/write only (600)
            # This is important for security when the file contains API keys
            try:
                os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)
            except OSError:
                # On Windows, chmod may not work as expected
                pass
                
        except IOError as e:
            raise ConfigError(f"Failed to create template file: {e}")
    
    def get_all(self) -> Dict[str, Any]:
        """Get the entire configuration dictionary.
        
        Returns:
            Complete configuration dictionary
        """
        return self._deep_copy(self._config)
    
    def reload(self) -> None:
        """Reload configuration from the original file.
        
        Useful when the configuration file has been modified externally.
        """
        # Reset to defaults
        self._config = self._deep_copy(self.DEFAULT_CONFIG)
        
        # Reload from file if path was provided
        if self.config_path and Path(self.config_path).exists():
            file_config = self._load_yaml(self.config_path)
            self._merge_config(file_config)
        
        # Re-substitute environment variables
        self._substitute_env_vars(self._config)
    
    def _detect_format_and_orientation(self) -> None:
        """检测配置格式（旧版 vs 新版）并设置文档方向。
        
        新格式包含 'document_orientation' 字段和独立的配置区块。
        旧格式的配置直接位于根级别。
        
        Detect configuration format (legacy vs new) and set orientation.
        
        New format has 'document_orientation' field and separate config blocks.
        Legacy format has configuration directly at the root level.
        """
        if 'document_orientation' in self._raw_config:
            # 新的双配置格式
            # New dual-configuration format
            self._is_legacy = False
            orientation = self._raw_config['document_orientation']
            
            # 验证方向值
            # Validate orientation value
            if orientation not in ['vertical', 'horizontal', 'auto']:
                raise ConfigError(
                    f"Invalid document_orientation value: '{orientation}'. "
                    f"Must be 'vertical', 'horizontal', or 'auto'."
                )
            
            # 如果是 auto 模式，默认先加载 horizontal 配置
            # 实际方向会在运行时根据图片检测结果动态切换
            if orientation == 'auto':
                self._orientation = 'horizontal'  # 默认先用横版
                logger.info(f"加载自动检测配置格式（默认: horizontal）| Loaded auto-detection configuration format (default: horizontal)")
            else:
                self._orientation = orientation
                logger.info(f"加载双配置格式，文档方向: {orientation} | Loaded dual-configuration format with orientation: {orientation}")
        else:
            # 旧的单配置格式
            # Legacy single-configuration format
            self._is_legacy = True
            self._orientation = 'horizontal'  # 默认为横版以保持向后兼容 | Default to horizontal for backward compatibility
            logger.debug("检测到旧版配置格式 | Detected legacy configuration format")
    
    def _extract_orientation_config(self) -> Dict[str, Any]:
        """Extract the configuration block for the current orientation.
        
        Returns:
            Configuration dictionary for the current orientation
            
        Raises:
            ConfigError: If the required configuration block is missing
        """
        # 注意：_orientation 在 auto 模式下已经被设置为默认值（horizontal）
        # 所以这里直接使用 _orientation 即可
        config_key = f"{self._orientation}_config"
        
        if config_key not in self._raw_config:
            raise ConfigError(
                f"Missing required configuration block: '{config_key}' "
                f"for orientation '{self._orientation}'"
            )
        
        orientation_config = self._raw_config[config_key]
        
        # Also merge global configuration if present
        global_config = {}
        for key, value in self._raw_config.items():
            if key not in ['document_orientation', 'vertical_config', 'horizontal_config']:
                global_config[key] = value
        
        # Merge global config first, then orientation-specific config
        # (orientation-specific config takes precedence)
        result = self._deep_copy(global_config)
        self._deep_merge(result, orientation_config)
        
        return result
    
    def get_orientation(self) -> str:
        """获取当前文档方向。
        
        Returns:
            当前方向："vertical"（竖版）或 "horizontal"（横版）
            
        Get the current document orientation.
        
        Returns:
            Current orientation: "vertical" or "horizontal"
        """
        return self._orientation
    
    def is_auto_mode(self) -> bool:
        """检查配置是否使用自动检测模式。
        
        Returns:
            如果配置为 auto 模式则返回 True
            
        Check if the configuration uses auto-detection mode.
        
        Returns:
            True if using auto-detection mode
        """
        return self._raw_config.get('document_orientation') == 'auto'
    
    def switch_orientation(self, orientation: str) -> None:
        """动态切换文档方向配置（仅在 auto 模式下有效）。
        
        Args:
            orientation: 目标方向 "vertical" 或 "horizontal"
            
        Dynamically switch document orientation configuration (only works in auto mode).
        
        Args:
            orientation: Target orientation "vertical" or "horizontal"
            
        Raises:
            ConfigError: 如果不是 auto 模式或方向值无效
        """
        if not self.is_auto_mode():
            logger.warning(
                f"Cannot switch orientation: configuration is not in auto mode "
                f"(current: {self._raw_config.get('document_orientation')})"
            )
            return
        
        if orientation not in ['vertical', 'horizontal']:
            raise ConfigError(
                f"Invalid orientation value: '{orientation}'. "
                f"Must be 'vertical' or 'horizontal'."
            )
        
        if orientation == self._orientation:
            logger.debug(f"Orientation already set to {orientation}, no switch needed")
            return
        
        logger.info(f"🔄 切换配置方向: {self._orientation} -> {orientation} | Switching orientation: {self._orientation} -> {orientation}")
        
        # 更新方向
        self._orientation = orientation
        
        # 重新加载配置
        self._config = self._deep_copy(self.DEFAULT_CONFIG)
        file_config = self._extract_orientation_config()
        self._merge_config(file_config)
        self._substitute_env_vars(self._config)
        
        logger.info(f"✅ 配置已切换到 {orientation} 模式 | Configuration switched to {orientation} mode")
    
    def is_legacy_format(self) -> bool:
        """检查配置是否使用旧版格式。
        
        Returns:
            如果使用旧版单配置格式则返回 True，否则返回 False
            
        Check if the configuration uses legacy format.
        
        Returns:
            True if using legacy single-configuration format, False otherwise
        """
        return self._is_legacy
