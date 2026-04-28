"""GrayWolfe — main entry point.

Initializes the local SQLite database and launches the Tkinter application.

Usage:
    python main.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path when run directly
sys.path.insert(0, str(Path(__file__).parent))

import config
from data.database import get_connection, initialize_schema, get_current_user
from util.logger import get_logger

log = get_logger(__name__)


def _resolve_local_db_path() -> Path:
    """Return the path to this user's local SQLite database file.

    Path: {USER_DB_DIR}/{username}.db
    Creates the parent directory if it doesn't exist (or falls back to a
    local ``db/`` directory if the shared drive is unreachable).
    """
    username = get_current_user()
    candidate: Path = config.USER_DB_DIR / f"{username}.db"
    try:
        candidate.parent.mkdir(parents=True, exist_ok=True)
        return candidate
    except OSError:
        log.warning(
            "Shared drive unreachable (%s). Falling back to local db/ directory.",
            config.USER_DB_DIR,
        )
        fallback_dir: Path = Path(__file__).parent / "db"
        fallback_dir.mkdir(exist_ok=True)
        return fallback_dir / f"{username}.db"


def main() -> None:
    """Initialize DB and launch the GUI."""
    # Defer tkinter import so import errors surface clearly
    try:
        import tkinter as tk
    except ImportError as exc:
        print(f"ERROR: tkinter is required but not available: {exc}", file=sys.stderr)
        sys.exit(1)

    from display.main import GrayWolfeApp

    db_path = _resolve_local_db_path()
    log.info("Starting %s v%s — DB: %s", config.APP_NAME, config.APP_VERSION, db_path)

    conn = get_connection(db_path)
    initialize_schema(conn)

    app = GrayWolfeApp(conn)
    app.mainloop()

    conn.close()
    log.info("GrayWolfe shut down cleanly.")


if __name__ == "__main__":
    main()
