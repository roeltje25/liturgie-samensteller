"""Logging configuration for Liturgy Builder."""

import logging
import os
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
