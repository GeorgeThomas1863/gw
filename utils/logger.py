from __future__ import annotations

import logging
import sys
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

import config

_LOG_FORMAT: str = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"
_INITIALIZED: set[str] = set()


def _resolve_log_dir() -> Path:
    """Return a writable log directory, falling back to a local logs/ folder."""
    candidate: Path = config.LOG_DIR
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        probe = candidate / ".write_probe"
        probe.touch()
        probe.unlink()
        return candidate
    except OSError:
        fallback: Path = Path(__file__).parent.parent / "logs"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def get_logger(name: str) -> logging.Logger:
    """Return a named logger, attaching handlers on first call.

    Subsequent calls with the same name return the existing logger without
    adding duplicate handlers.
    """
    logger = logging.getLogger(name)

    if name in _INITIALIZED:
        return logger

    _INITIALIZED.add(name)
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)

    # File handler — daily rotation, 30-day retention
    log_dir = _resolve_log_dir()
    file_handler = TimedRotatingFileHandler(
        filename=str(log_dir / "graywolfe.log"),
        when="midnight",
        interval=1,
        backupCount=30,
        encoding="utf-8",
        utc=False,
        delay=True,
    )
    file_handler.suffix = "%Y-%m-%d"
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # Console handler for dev convenience
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(logging.DEBUG)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.propagate = False

    return logger
