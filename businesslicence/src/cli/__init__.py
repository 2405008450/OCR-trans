"""Command Line Interface for image translation system.

This module provides the CLI for the image translation system,
supporting single image translation, batch processing, and
various configuration options.

_Requirements: 13.1, 13.2, 13.3, 13.4, 13.5_
"""

from src.cli.main import main, create_parser, run_cli

__all__ = ['main', 'create_parser', 'run_cli']
