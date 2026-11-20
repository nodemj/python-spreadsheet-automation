"""Structured logging configuration for sheetclean."""

from __future__ import annotations

import logging
from pathlib import Path


def setup_logging(
    log_file: Path | None = None,
    verbose: bool = False,
) -> logging.Logger:
    """Configure console and optional file logging for sheetclean.

    Args:
        log_file: Optional path to write detailed log records.
        verbose: When True, sets console level to DEBUG; otherwise INFO.

    Returns:
        Configured ``sheetclean`` logger instance.
    """
    logger = logging.getLogger("sheetclean")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if log_file is not None:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
