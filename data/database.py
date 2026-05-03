"""
data/database.py — SQLite connection manager, schema initializer, and CRUD helpers.
"""
from __future__ import annotations

import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

import dataclasses

from models import Selector, SelectorType, Target
from config import ID_STRFTIME
from util.errors import (
    ERR_DB_UPDATE,
    ERR_SELECTOR_DUPLICATE,
    ERR_TARGET_NOT_FOUND,
    GWError,
    raise_gw,
)

# ---------------------------------------------------------------------------
# Sequence for generate_id — guarantees uniqueness even under rapid calls.
# Tracks the last emitted full 15-char ID and increments the ms field if the
# natural value would produce a collision.
# ---------------------------------------------------------------------------
_last_id: str = ""
_id_lock = threading.Lock()

# ---------------------------------------------------------------------------
# DDL
# ---------------------------------------------------------------------------

_DDL = """
CREATE TABLE IF NOT EXISTS selectors (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    selector_id     TEXT UNIQUE NOT NULL,
    selector        TEXT NOT NULL,
    selector_clean  TEXT NOT NULL,
    selector_type   TEXT NOT NULL,
    target_id       TEXT REFERENCES targets(target_id),
    nork_id         TEXT REFERENCES norks(nork_id),
    date_created    TEXT NOT NULL,
    created_by      TEXT NOT NULL,
    last_updated    TEXT NOT NULL,
    last_updated_by TEXT NOT NULL,
    data_source     TEXT
);
CREATE INDEX IF NOT EXISTS idx_sel_clean  ON selectors(selector_clean);
CREATE INDEX IF NOT EXISTS idx_sel_target ON selectors(target_id);

CREATE TABLE IF NOT EXISTS targets (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    target_id       TEXT UNIQUE NOT NULL,
    target_name     TEXT,
    case_number     TEXT,
    laptop_count    INTEGER DEFAULT 0,
    date_created    TEXT NOT NULL,
    created_by      TEXT NOT NULL,
    last_updated    TEXT NOT NULL,
    last_updated_by TEXT NOT NULL,
    data_source     TEXT
);

CREATE TABLE IF NOT EXISTS norks (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nork_id         TEXT UNIQUE NOT NULL,
    nork_name       TEXT,
    date_created    TEXT NOT NULL,
    created_by      TEXT NOT NULL,
    last_updated    TEXT NOT NULL,
    last_updated_by TEXT NOT NULL
);
"""


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

def get_connection(db_path: str | Path) -> sqlite3.Connection:
    """Open a SQLite connection with WAL mode and row_factory = sqlite3.Row."""
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

def initialize_schema(conn: sqlite3.Connection) -> None:
    """Create all tables and indexes if they don't exist.

    Also sets row_factory = sqlite3.Row on the connection so all subsequent
    queries return Row objects regardless of how the connection was opened.
    """
    conn.row_factory = sqlite3.Row
    conn.executescript(_DDL)
    conn.commit()


# ---------------------------------------------------------------------------
# ID / timestamp helpers
# ---------------------------------------------------------------------------

def generate_id() -> str:
    """Return a unique 15-char ID: YYMMDDHHMMSS + zero-padded 3-digit ms.

    The YYMMDDHHMMSS prefix is derived from wall-clock time. The 3-digit ms
    suffix starts from the actual millisecond of the current time and is
    incremented (within [0, 999]) if the candidate would duplicate the previous
    ID.  If ms wraps past 999 without finding a free slot, the function yields a
    new natural ID once wall-clock time has advanced — ensuring uniqueness
    without busy-waiting.
    """
    global _last_id
    with _id_lock:
        now = datetime.now()
        ts = now.strftime(ID_STRFTIME)
        ms = int(now.microsecond / 1000)

        candidate = f"{ts}{ms:03d}"
        if candidate <= _last_id and _last_id.startswith(ts):
            # Bump ms past the last-used value for this second.
            last_ms = int(_last_id[12:15])
            ms = (last_ms + 1) % 1000
            candidate = f"{ts}{ms:03d}"

        _last_id = candidate
        return candidate


def get_current_user() -> str:
    """Return the Windows username (lowercase) from the environment."""
    username = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or os.environ.get("LOGNAME")
        or "unknown"
    )
    return username.lower()


def now_iso() -> str:
    """Return current UTC datetime as ISO 8601 string."""
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


# ---------------------------------------------------------------------------
# Selectors
# ---------------------------------------------------------------------------

def insert_selector(conn: sqlite3.Connection, selector: Selector) -> str:
    """Insert a selector row. Returns selector_id.

    Raises GWError(ERR_SELECTOR_DUPLICATE) if selector_clean already exists.
    The duplicate check is import-wide: any matching clean value blocks the
    insert, regardless of target_id.
    """
    existing = conn.execute(
        "SELECT selector_id FROM selectors WHERE selector_clean = ?",
        (selector.selector_clean,),
    ).fetchone()
    if existing is not None:
        raise GWError(
            ERR_SELECTOR_DUPLICATE,
            f"Selector '{selector.selector_clean}' already exists "
            f"(id={existing['selector_id']})",
        )

    conn.execute(
        """
        INSERT INTO selectors (
            selector_id, selector, selector_clean, selector_type,
            target_id, nork_id,
            date_created, created_by,
            last_updated, last_updated_by,
            data_source
        ) VALUES (
            :selector_id, :selector, :selector_clean, :selector_type,
            :target_id, :nork_id,
            :date_created, :created_by,
            :last_updated, :last_updated_by,
            :data_source
        )
        """,
        dataclasses.asdict(selector),
    )
    conn.commit()
    return selector.selector_id


def get_selector(conn: sqlite3.Connection, selector_clean: str) -> Selector | None:
    """Return selector row as Selector by selector_clean, or None."""
    row = conn.execute(
        "SELECT * FROM selectors WHERE selector_clean = ?",
        (selector_clean,),
    ).fetchone()
    if row is None:
        return None
    data = {k: row[k] for k in row.keys() if k != "id"}
    data["selector_type"] = SelectorType(data["selector_type"])
    return Selector(**data)


# ---------------------------------------------------------------------------
# Targets
# ---------------------------------------------------------------------------

def insert_target(conn: sqlite3.Connection, target: Target) -> str:
    """Insert a target row. Returns target_id."""
    conn.execute(
        """
        INSERT INTO targets (
            target_id, target_name, case_number, laptop_count,
            date_created, created_by,
            last_updated, last_updated_by,
            data_source
        ) VALUES (
            :target_id, :target_name, :case_number, :laptop_count,
            :date_created, :created_by,
            :last_updated, :last_updated_by,
            :data_source
        )
        """,
        dataclasses.asdict(target),
    )
    conn.commit()
    return target.target_id


def get_target(conn: sqlite3.Connection, target_id: str) -> Target | None:
    """Return target row as Target by target_id, or None."""
    row = conn.execute(
        "SELECT * FROM targets WHERE target_id = ?",
        (target_id,),
    ).fetchone()
    if row is None:
        return None
    return Target(**{k: row[k] for k in row.keys() if k != "id"})


def update_target(
    conn: sqlite3.Connection,
    target_id: str,
    fields: dict,
    updated_by: str,
) -> None:
    """Update the given fields on a target.

    Always refreshes last_updated and last_updated_by regardless of what
    is passed in fields.
    """
    _VALID_TARGET_FIELDS = {
        "target_name", "case_number", "laptop_count", "last_updated", "last_updated_by"
    }
    invalid = set(fields.keys()) - _VALID_TARGET_FIELDS
    if invalid:
        raise_gw(
            ERR_DB_UPDATE,
            f"update_target() received invalid field(s): {', '.join(sorted(invalid))}",
        )

    # Merge timestamp refresh into the field set so a single UPDATE suffices.
    updates = {**fields, "last_updated": now_iso(), "last_updated_by": updated_by}

    set_clause = ", ".join(f"{col} = :{col}" for col in updates)
    updates["_target_id"] = target_id

    conn.execute(
        f"UPDATE targets SET {set_clause} WHERE target_id = :_target_id",
        updates,
    )
    conn.commit()


def update_target_id(
    conn: sqlite3.Connection,
    old_id: str,
    new_id: str,
    updated_by: str,
) -> None:
    """Rename a target's ID across targets and selectors tables atomically.

    Raises GWError(ERR_TARGET_NOT_FOUND) if old_id doesn't exist.
    Raises GWError(ERR_DB_UPDATE, "Target ID already exists") if new_id is already taken.
    Uses PRAGMA defer_foreign_keys to defer FK constraint checks until COMMIT,
    since selectors.target_id references targets(target_id).
    """
    # Validate old_id exists
    if get_target(conn, old_id) is None:
        raise_gw(ERR_TARGET_NOT_FOUND, f"Target '{old_id}' not found.")

    # Validate new_id not taken
    if get_target(conn, new_id) is not None:
        raise_gw(ERR_DB_UPDATE, f"Target ID '{new_id}' already exists.")

    # Defer FK checks until COMMIT so both updates can complete safely
    conn.execute("PRAGMA defer_foreign_keys = ON")
    conn.execute(
        "UPDATE targets SET target_id = ?, last_updated = ?, last_updated_by = ? WHERE target_id = ?",
        (new_id, now_iso(), updated_by, old_id),
    )
    conn.execute(
        "UPDATE selectors SET target_id = ? WHERE target_id = ?",
        (new_id, old_id),
    )
    conn.commit()


def delete_target(conn: sqlite3.Connection, target_id: str) -> None:
    """Delete a target by target_id. Silently succeeds if not found."""
    conn.execute("DELETE FROM targets WHERE target_id = ?", (target_id,))
    conn.commit()
