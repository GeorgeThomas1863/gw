"""
data/sync.py — Master/local DB sync logic.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

from util.errors import ERR_SYNC_FAILED, GWError
from util.logger import get_logger

logger = get_logger("gw.sync")

# Tables to sync in dependency order (targets before selectors so FK is satisfied).
_SYNC_TABLES: tuple[str, ...] = ("norks", "targets", "selectors")

# Column names for each table (must match DDL in database.py, minus the rowid 'id').
# Exported (no leading underscore) so admin/merge_user_db.py can import the same list.
TABLE_COLUMNS: dict[str, tuple[str, ...]] = {
    "selectors": (
        "selector_id", "selector", "selector_clean", "selector_type",
        "target_id", "nork_id",
        "date_created", "created_by",
        "last_updated", "last_updated_by",
        "data_source",
    ),
    "targets": (
        "target_id", "target_name", "case_number", "laptop_count",
        "date_created", "created_by",
        "last_updated", "last_updated_by",
        "data_source",
    ),
    "norks": (
        "nork_id", "nork_name",
        "date_created", "created_by",
        "last_updated", "last_updated_by",
    ),
}

# Primary key column per table.
_PK: dict[str, str] = {
    "selectors": "selector_id",
    "targets": "target_id",
    "norks": "nork_id",
}

# Stat key per table.
_STAT_KEY: dict[str, str] = {
    "selectors": "selectors_added",
    "targets": "targets_added",
    "norks": "norks_added",
}


def pull_from_master(
    local_conn: sqlite3.Connection,
    master_db_path: "str | Path | sqlite3.Connection",
) -> dict:
    """Pull all records from master DB into local DB.

    Strategy: INSERT OR REPLACE for all tables (master wins on conflict).

    Parameters
    ----------
    local_conn:
        Open SQLite connection to the local (user) database.
    master_db_path:
        Path to the master DB file, OR an already-open sqlite3.Connection
        (the latter is used by tests to pass an in-memory DB directly).

    Returns
    -------
    dict
        ``{"selectors_added": N, "targets_added": M, "norks_added": K}``

    Raises
    ------
    GWError(ERR_SYNC_FAILED)
        If the master DB path is unreachable or cannot be opened.
    """
    master_conn, _we_opened = _open_master(master_db_path)

    stats: dict[str, int] = {v: 0 for v in _STAT_KEY.values()}

    try:
        for table in _SYNC_TABLES:
            count = _sync_table(local_conn, master_conn, table)
            stats[_STAT_KEY[table]] = count

        local_conn.commit()
        logger.debug("pull_from_master complete: %s", stats)
    finally:
        if _we_opened:
            master_conn.close()

    return stats


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

def _open_master(
    source: "str | Path | sqlite3.Connection",
) -> tuple[sqlite3.Connection, bool]:
    """Return (connection, we_opened_it).

    If *source* is already a Connection, return it as-is (caller owns it).
    If *source* is a path, open read-only; raise GWError on failure.
    """
    if isinstance(source, sqlite3.Connection):
        return source, False

    path = Path(source)
    uri = f"file:{path}?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        # Trigger a lightweight probe to confirm the file is actually readable.
        conn.execute("SELECT 1")
        return conn, True
    except (sqlite3.OperationalError, sqlite3.DatabaseError) as exc:
        raise GWError(
            ERR_SYNC_FAILED,
            f"Cannot open master DB at '{source}': {exc}",
        ) from exc


def _sync_table(
    local_conn: sqlite3.Connection,
    master_conn: sqlite3.Connection,
    table: str,
) -> int:
    """Copy all rows from master *table* into local using INSERT OR REPLACE.

    Returns the number of rows written (includes rows that replaced existing ones).
    """
    cols = TABLE_COLUMNS[table]
    pk = _PK[table]
    col_list = ", ".join(cols)
    placeholder_list = ", ".join("?" * len(cols))

    master_rows = master_conn.execute(
        f"SELECT {col_list} FROM {table}"  # noqa: S608
    ).fetchall()

    if not master_rows:
        return 0

    # Count rows that are genuinely new (not already present in local).
    existing_pks: set[str] = {
        row[0]
        for row in local_conn.execute(f"SELECT {pk} FROM {table}").fetchall()  # noqa: S608
    }

    rows_to_insert = [tuple(row[c] for c in cols) for row in master_rows]

    local_conn.executemany(
        f"INSERT OR REPLACE INTO {table} ({col_list}) VALUES ({placeholder_list})",  # noqa: S608
        rows_to_insert,
    )

    # "added" = rows that were not in local before this pull.
    new_count = sum(
        1 for row in master_rows if row[pk] not in existing_pks
    )
    logger.debug("sync_table '%s': %d master rows, %d new", table, len(master_rows), new_count)
    return new_count
