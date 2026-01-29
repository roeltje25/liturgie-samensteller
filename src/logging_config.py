"""Logging configuration for Liturgy Builder."""

import getpass
import logging
import os
import platform
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler

# Create logger
logger = logging.getLogger("liturgy_builder")
logger.setLevel(logging.DEBUG)

# Determine log file location
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    log_dir = os.path.dirname(sys.executable)
else:
    # Running as script
    log_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

log_file = os.path.join(log_dir, "liturgy_builder.log")

# File handler with rotation (max 5MB, keep 3 backups)
try:
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=5*1024*1024,
        backupCount=3,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s - %(message)s'
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
except Exception as e:
    # If we can't create the log file, just use console
    print(f"Warning: Could not create log file: {e}")

# Console handler for errors and above
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.WARNING)
console_formatter = logging.Formatter('%(levelname)s: %(message)s')
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)


def get_logger(name: str = None) -> logging.Logger:
    """Get a logger instance.

    Args:
        name: Optional name for the logger (will be prefixed with 'liturgy_builder.')

    Returns:
        Logger instance
    """
    if name:
        return logging.getLogger(f"liturgy_builder.{name}")
    return logger


def log_startup_info() -> None:
    """Log startup banner with version and user information."""
    from . import __version__, __revision__, __build_date__

    # ASCII banner
    banner = r"""
    ╔═══════════════════════════════════════════════════════════════════╗
    ║                                                                   ║
    ║   ██╗     ██╗████████╗██╗   ██╗██████╗  ██████╗ ██╗███████╗       ║
    ║   ██║     ██║╚══██╔══╝██║   ██║██╔══██╗██╔════╝ ██║██╔════╝       ║
    ║   ██║     ██║   ██║   ██║   ██║██████╔╝██║  ███╗██║█████╗         ║
    ║   ██║     ██║   ██║   ██║   ██║██╔══██╗██║   ██║██║██╔══╝         ║
    ║   ███████╗██║   ██║   ╚██████╔╝██║  ██║╚██████╔╝██║███████╗       ║
    ║   ╚══════╝╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚═╝╚══════╝       ║
    ║                                                                   ║
    ║              ███████╗ █████╗ ███╗   ███╗███████╗███╗   ██╗        ║
    ║              ██╔════╝██╔══██╗████╗ ████║██╔════╝████╗  ██║        ║
    ║              ███████╗███████║██╔████╔██║█████╗  ██╔██╗ ██║        ║
    ║              ╚════██║██╔══██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║        ║
    ║              ███████║██║  ██║██║ ╚═╝ ██║███████╗██║ ╚████║        ║
    ║              ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝        ║
    ║                                                                   ║
    ║                    Liturgie Samensteller                          ║
    ║                     PowerPoint Mixer                              ║
    ║                                                                   ║
    ╚═══════════════════════════════════════════════════════════════════╝
    """

    # Get user info
    try:
        username = getpass.getuser()
    except Exception:
        username = "unknown"

    # Get system info
    try:
        hostname = platform.node()
    except Exception:
        hostname = "unknown"

    # Log the banner
    for line in banner.strip().split('\n'):
        logger.info(line)

    # Log version and system info
    logger.info("=" * 71)
    logger.info(f"  Version: {__version__} (revision: {__revision__})")
    logger.info(f"  Build date: {__build_date__}")
    logger.info(f"  User: {username}@{hostname}")
    logger.info(f"  Python: {platform.python_version()}")
    logger.info(f"  Platform: {platform.system()} {platform.release()}")
    logger.info(f"  Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 71)
