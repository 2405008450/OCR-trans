"""Command Line Interface main module for image translation system.

This module provides the main CLI functionality including argument parsing,
execution logic, error handling, and help information.

_Requirements: 13.1, 13.2, 13.3, 13.4, 13.5, 10.2_
"""

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from src.config import ConfigManager
from src.pipeline import TranslationPipeline
from src.models import QualityReport
from src.exceptions import (
    ImageTranslationError,
    ConfigError,
    PipelineError,
    ImageLoadError,
    ImageSaveError
)


# Configure module logger
logger = logging.getLogger(__name__)


# Supported image extensions
SUPPORTED_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp'}


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser.
    
    Creates an ArgumentParser with all supported command-line arguments
    for the image translation system.
    
    Returns:
        Configured ArgumentParser instance
        
    _Requirements: 13.1, 13.2, 13.3, 13.4_
    """
    parser = argparse.ArgumentParser(
        prog='image-translate',
        description='Image Translation System - Translate text in images from Chinese to English',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Translate a single image:
    %(prog)s input.jpg -o output.jpg
    
  Translate with custom output path:
    %(prog)s input.png --output translated/result.png
    
  Batch translate all images in a folder:
    %(prog)s input_folder/ -o output_folder/ --batch
    
  Use custom configuration file:
    %(prog)s input.jpg -o output.jpg --config my_config.yaml
    
  Enable verbose output:
    %(prog)s input.jpg -o output.jpg --verbose
    
  Specify source and target languages:
    %(prog)s input.jpg -o output.jpg --source-lang zh --target-lang en
    
  Generate configuration template:
    %(prog)s --generate-config config.yaml

Configuration:
  The system looks for configuration in the following order:
  1. Command-line specified config file (--config)
  2. config.yaml in current directory
  3. Environment variables (DEEPSEEK_API_KEY)
  4. Default values

For more information, visit: https://github.com/your-repo/image-translation
"""
    )
    
    # Positional argument: input path (optional when using --generate-config)
    parser.add_argument(
        'input',
        type=str,
        nargs='?',
        default=None,
        help='Input image path or directory for batch processing'
    )
    
    # Output path (optional when using --generate-config)
    parser.add_argument(
        '-o', '--output',
        type=str,
        default=None,
        help='Output image path or directory for batch processing'
    )
    
    # Configuration file
    parser.add_argument(
        '-c', '--config',
        type=str,
        default=None,
        help='Path to configuration file (YAML format)'
    )
    
    # Batch processing mode
    parser.add_argument(
        '-b', '--batch',
        action='store_true',
        help='Enable batch processing mode for directories'
    )
    
    # Verbose mode
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output with detailed progress information'
    )
    
    # Quiet mode
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress all output except errors'
    )
    
    # Source language
    parser.add_argument(
        '--source-lang',
        type=str,
        default='zh',
        help='Source language code (default: zh for Chinese)'
    )
    
    # Target language
    parser.add_argument(
        '--target-lang',
        type=str,
        default='en',
        help='Target language code (default: en for English)'
    )
    
    # Parallel processing
    parser.add_argument(
        '--parallel',
        action='store_true',
        default=None,
        help='Enable parallel processing for batch mode'
    )
    
    parser.add_argument(
        '--no-parallel',
        action='store_true',
        help='Disable parallel processing for batch mode'
    )
    
    # Generate config template
    parser.add_argument(
        '--generate-config',
        type=str,
        metavar='PATH',
        help='Generate a configuration template file at the specified path and exit'
    )
    
    # Version
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 1.0.0'
    )
    
    return parser


def validate_args(args: argparse.Namespace) -> List[str]:
    """Validate command-line arguments.
    
    Checks that the provided arguments are valid and consistent.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        List of error messages. Empty list indicates valid arguments.
        
    _Requirements: 13.5_
    """
    errors = []
    
    # Skip validation if generating config
    if args.generate_config:
        return errors
    
    # Check required arguments for translation mode
    if args.input is None:
        errors.append("Input path is required")
        return errors
    
    if args.output is None:
        errors.append("Output path is required (use -o/--output)")
        return errors
    
    # Check input path exists
    if not os.path.exists(args.input):
        errors.append(f"Input path does not exist: {args.input}")
        return errors
    
    # Check batch mode consistency
    is_input_dir = os.path.isdir(args.input)
    
    if args.batch and not is_input_dir:
        errors.append("Batch mode requires input to be a directory")
    
    if is_input_dir and not args.batch:
        errors.append(
            "Input is a directory. Use --batch flag for batch processing"
        )
    
    # Check output path for single file mode
    if not args.batch and not is_input_dir:
        # Single file mode - output should be a file path
        output_dir = os.path.dirname(args.output)
        if output_dir and not os.path.exists(output_dir):
            # Will be created, but warn if parent doesn't exist
            parent_dir = os.path.dirname(output_dir)
            if parent_dir and not os.path.exists(parent_dir):
                errors.append(
                    f"Output directory parent does not exist: {parent_dir}"
                )
        
        # Check input file extension
        input_ext = Path(args.input).suffix.lower()
        if input_ext not in SUPPORTED_EXTENSIONS:
            errors.append(
                f"Unsupported input file format: {input_ext}. "
                f"Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
            )
    
    # Check conflicting options
    if args.verbose and args.quiet:
        errors.append("Cannot use both --verbose and --quiet options")
    
    if args.parallel and args.no_parallel:
        errors.append("Cannot use both --parallel and --no-parallel options")
    
    return errors


def setup_logging(verbose: bool = False, quiet: bool = False) -> None:
    """Configure logging based on verbosity settings.
    
    Args:
        verbose: Enable verbose (DEBUG) logging
        quiet: Suppress all output except errors
        
    _Requirements: 13.3_
    """
    if quiet:
        level = logging.ERROR
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO
    
    # Configure root logger
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Also set level for our modules
    for module in ['src.cli', 'src.pipeline', 'src.ocr', 'src.translation']:
        logging.getLogger(module).setLevel(level)


def find_images_in_directory(directory: str) -> List[str]:
    """Find all supported image files in a directory.
    
    Args:
        directory: Path to directory to search
        
    Returns:
        List of image file paths
    """
    images = []
    
    for root, _, files in os.walk(directory):
        for file in files:
            if Path(file).suffix.lower() in SUPPORTED_EXTENSIONS:
                images.append(os.path.join(root, file))
    
    return sorted(images)


def print_quality_report(report: QualityReport, verbose: bool = False) -> None:
    """Print a quality report to stdout.
    
    Args:
        report: Quality report to print
        verbose: Whether to include detailed information
        
    _Requirements: 13.2_
    """
    print(f"\n{'='*50}")
    print("Translation Quality Report")
    print(f"{'='*50}")
    print(f"Overall Quality: {report.overall_quality.value.upper()}")
    print(f"Translation Coverage: {report.translation_coverage:.1%}")
    print(f"Total Regions: {report.total_regions}")
    print(f"Translated Regions: {report.translated_regions}")
    print(f"Failed Regions: {len(report.failed_regions)}")
    print(f"Has Artifacts: {'Yes' if report.has_artifacts else 'No'}")
    
    if verbose and report.failed_regions:
        print(f"\nFailed Regions Details:")
        for i, region in enumerate(report.failed_regions, 1):
            print(f"  {i}. Text: '{region.text[:30]}...' at {region.bbox}")
    
    if verbose and report.artifact_locations:
        print(f"\nArtifact Locations:")
        for i, loc in enumerate(report.artifact_locations, 1):
            print(f"  {i}. {loc}")
    
    print(f"{'='*50}\n")


def print_batch_summary(
    results: List[Tuple[str, Optional[QualityReport], Optional[str]]]
) -> None:
    """Print a summary of batch processing results.
    
    Args:
        results: List of (input_path, report, error) tuples
        
    _Requirements: 13.2_
    """
    total = len(results)
    successful = sum(1 for _, report, _ in results if report is not None)
    failed = total - successful
    
    print(f"\n{'='*50}")
    print("Batch Processing Summary")
    print(f"{'='*50}")
    print(f"Total Images: {total}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    
    if successful > 0:
        coverages = [
            report.translation_coverage 
            for _, report, _ in results 
            if report is not None
        ]
        avg_coverage = sum(coverages) / len(coverages)
        print(f"Average Coverage: {avg_coverage:.1%}")
    
    if failed > 0:
        print(f"\nFailed Images:")
        for path, _, error in results:
            if error:
                print(f"  - {path}: {error}")
    
    print(f"{'='*50}\n")


def run_single_translation(
    pipeline: TranslationPipeline,
    input_path: str,
    output_path: str,
    source_lang: str,
    target_lang: str,
    verbose: bool
) -> Tuple[bool, Optional[QualityReport], Optional[str]]:
    """Run translation for a single image.
    
    Args:
        pipeline: Translation pipeline instance
        input_path: Path to input image
        output_path: Path for output image
        source_lang: Source language code
        target_lang: Target language code
        verbose: Whether to print verbose output
        
    Returns:
        Tuple of (success, report, error_message)
        
    _Requirements: 13.1, 13.2_
    """
    print(f"Translating: {input_path}")
    
    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    report, error = pipeline.translate_image_safe(
        input_path, output_path, source_lang, target_lang
    )
    
    if report:
        print(f"Output saved to: {output_path}")
        print_quality_report(report, verbose)
        return True, report, None
    else:
        return False, None, error


def run_batch_translation(
    pipeline: TranslationPipeline,
    input_dir: str,
    output_dir: str,
    source_lang: str,
    target_lang: str,
    parallel: Optional[bool],
    verbose: bool
) -> Tuple[bool, List[Tuple[str, Optional[QualityReport], Optional[str]]]]:
    """Run batch translation for a directory of images.
    
    Args:
        pipeline: Translation pipeline instance
        input_dir: Path to input directory
        output_dir: Path to output directory
        source_lang: Source language code
        target_lang: Target language code
        parallel: Whether to use parallel processing
        verbose: Whether to print verbose output
        
    Returns:
        Tuple of (all_successful, results_list)
        
    _Requirements: 13.2, 13.4_
    """
    # Find all images
    images = find_images_in_directory(input_dir)
    
    if not images:
        print(f"No supported images found in: {input_dir}")
        print(f"Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")
        return False, []
    
    print(f"Found {len(images)} images to process")
    print(f"Output directory: {output_dir}")
    
    # Run batch translation
    results = pipeline.translate_batch(
        images, output_dir, source_lang, target_lang, parallel
    )
    
    # Print summary
    print_batch_summary(results)
    
    # Print individual reports if verbose
    if verbose:
        for path, report, error in results:
            if report:
                print(f"\n--- {os.path.basename(path)} ---")
                print_quality_report(report, verbose=False)
    
    # Check if all succeeded
    all_successful = all(report is not None for _, report, _ in results)
    
    return all_successful, results


def run_cli(args: argparse.Namespace) -> int:
    """Execute the CLI with parsed arguments.
    
    Main execution logic that coordinates configuration loading,
    pipeline initialization, and translation execution.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        Exit code (0 for success, non-zero for failure)
        
    _Requirements: 13.1, 13.2, 10.2_
    """
    # Handle config template generation
    if args.generate_config:
        try:
            config = ConfigManager()
            config.create_template(args.generate_config)
            print(f"Configuration template created: {args.generate_config}")
            return 0
        except ConfigError as e:
            print(f"Error creating config template: {e}", file=sys.stderr)
            return 1
    
    # Setup logging
    setup_logging(args.verbose, args.quiet)
    
    # Validate arguments
    errors = validate_args(args)
    if errors:
        for error in errors:
            print(f"Error: {error}", file=sys.stderr)
        print("\nUse --help for usage information.", file=sys.stderr)
        return 1
    
    try:
        # Load configuration
        # 加载配置
        if not args.quiet:
            print("Loading configuration...")
            print("加载配置...")
        
        # 确定配置文件路径
        # Determine configuration file path
        # Priority: 1. Command-line --config, 2. config.yaml in current directory, 3. Default config
        config_path = args.config
        if not config_path and os.path.exists('config.yaml'):
            config_path = 'config.yaml'
            if not args.quiet:
                print(f"Using configuration file: config.yaml")
                print(f"使用配置文件: config.yaml")
        elif config_path:
            if not args.quiet:
                print(f"Using configuration file: {config_path}")
                print(f"使用配置文件: {config_path}")
        else:
            if not args.quiet:
                print("Using default configuration")
                print("使用默认配置")
        
        # 初始化 ConfigManager，支持双配置架构（竖版/横版）
        # Initialize ConfigManager with dual-configuration architecture support (vertical/horizontal)
        # ConfigManager will automatically detect format and load the appropriate configuration
        config = ConfigManager(config_path=config_path)
        
        # 显示配置信息
        # Display configuration information
        if not args.quiet:
            orientation = config.get_orientation()
            is_legacy = config.is_legacy_format()
            if is_legacy:
                print(f"Configuration format: Legacy (mapped to horizontal)")
                print(f"配置格式: 旧版（映射到横版）")
                logger.info("Legacy configuration format detected and mapped to horizontal_config")
            else:
                print(f"Configuration format: Dual-configuration architecture")
                print(f"配置格式: 双配置架构")
                print(f"Document orientation: {orientation}")
                print(f"文档方向: {orientation}")
                logger.info(f"Dual-configuration format loaded with orientation: {orientation}")
        
        # Validate configuration
        # 验证配置
        config_errors = config.validate()
        if config_errors:
            print("Configuration errors:", file=sys.stderr)
            print("配置错误:", file=sys.stderr)
            for error in config_errors:
                print(f"  - {error}", file=sys.stderr)
            print("\nUse --generate-config to create a template.", file=sys.stderr)
            print("使用 --generate-config 创建配置模板。", file=sys.stderr)
            return 1
        
        # Initialize pipeline
        # 初始化翻译管道，ConfigManager 会被传递给所有核心模块
        if not args.quiet:
            print("Initializing translation pipeline...")
            print("初始化翻译管道...")
        
        # TranslationPipeline 会将 ConfigManager 传递给所有核心模块：
        # OCREngine, IconDetector, TranslationService, BackgroundSampler,
        # TextRenderer, ImageProcessor, QualityValidator
        # 所有核心模块将根据 ConfigManager 提供的配置（竖版或横版）动态调整行为
        # All core modules will dynamically adjust their behavior based on the configuration
        # (vertical or horizontal) provided by ConfigManager
        pipeline = TranslationPipeline(config)
        
        logger.info("Translation pipeline initialized with ConfigManager")
        
        # Determine parallel setting
        parallel = None
        if args.parallel:
            parallel = True
        elif args.no_parallel:
            parallel = False
        
        # Execute translation
        if args.batch or os.path.isdir(args.input):
            success, _ = run_batch_translation(
                pipeline,
                args.input,
                args.output,
                args.source_lang,
                args.target_lang,
                parallel,
                args.verbose
            )
        else:
            success, _, error = run_single_translation(
                pipeline,
                args.input,
                args.output,
                args.source_lang,
                args.target_lang,
                args.verbose
            )
            if error and not args.quiet:
                print(f"Translation failed: {error}", file=sys.stderr)
        
        return 0 if success else 1
        
    except ConfigError as e:
        print(f"Configuration error: {e}", file=sys.stderr)
        logger.debug("Configuration error details:", exc_info=True)
        return 1
    except ImageLoadError as e:
        print(f"Failed to load image: {e}", file=sys.stderr)
        logger.debug("Image load error details:", exc_info=True)
        return 1
    except ImageSaveError as e:
        print(f"Failed to save image: {e}", file=sys.stderr)
        logger.debug("Image save error details:", exc_info=True)
        return 1
    except PipelineError as e:
        print(f"Translation pipeline error: {e}", file=sys.stderr)
        logger.debug("Pipeline error details:", exc_info=True)
        return 1
    except ImageTranslationError as e:
        print(f"Translation error: {e}", file=sys.stderr)
        logger.debug("Translation error details:", exc_info=True)
        return 1
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.", file=sys.stderr)
        return 130
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        logger.debug("Unexpected error details:", exc_info=True)
        return 1


def main(argv: Optional[List[str]] = None) -> int:
    """Main entry point for the CLI.
    
    Parses command-line arguments and executes the appropriate action.
    
    Args:
        argv: Command-line arguments (defaults to sys.argv[1:])
        
    Returns:
        Exit code (0 for success, non-zero for failure)
        
    _Requirements: 13.1, 13.5_
    """
    parser = create_parser()
    
    try:
        args = parser.parse_args(argv)
    except SystemExit as e:
        # argparse calls sys.exit on error or --help/--version
        return e.code if e.code is not None else 0
    
    return run_cli(args)


if __name__ == '__main__':
    sys.exit(main())
