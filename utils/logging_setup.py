"""Structured logging setup for BBBG application."""

import os
import time
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

LOG_DIR = 'logs'
_log_file = None


def setup_logging():
    """Initialize structured file logging. Returns log file path."""
    global _log_file
    if _log_file is not None:
        return _log_file

    os.makedirs(LOG_DIR, exist_ok=True)

    try:
        cutoff = time.time() - (7 * 24 * 60 * 60)
        for f in os.listdir(LOG_DIR):
            path = os.path.join(LOG_DIR, f)
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                try:
                    os.remove(path)
                except OSError:
                    pass
    except OSError:
        pass

    timestamp = datetime.now().strftime('%Y-%m-%d-%H%M%S')
    _log_file = os.path.join(LOG_DIR, f'fix-{timestamp}.log')

    logger = logging.getLogger('bbbg')
    logger.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(
        logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    )

    file_handler = RotatingFileHandler(
        _log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter('[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s')
    )

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return _log_file


def get_logger(name: str = 'bbbg') -> logging.Logger:
    """Get or create logger with bbbg prefix."""
    setup_logging()
    return logging.getLogger(f'bbbg.{name}')


setup_logging()
