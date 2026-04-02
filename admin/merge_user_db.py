"""
admin/merge_user_db.py — CLI tool to merge a user's local DB into the master DB.

Usage:
    python admin/merge_user_db.py path/to/user.db [--dry-run]

Conflicts (same primary key, different last_updated) are flagged to a log file
and stdout but are NEVER written to master — the admin reviews them manually.
"""
from __future__ import annotations

import argparse
import datetime
import logging
import sqlite3
import sys
from pathlib import Path

# Ensure project root is importable when run as a script.
sys.path.insert(0, str(Path(__file__).parent.parent))

import config
from data.database import get_connection, initialize_schema
from utils.logger import get_logger

logger = get_logger("gw.admin.merge")

# Tables and their primary key columns (FK order: norks → targets → selectors).
_MERGE_TABLES: dict[str, str] = {
    "norks": "nork_id",
    "targets": "target_id",
    "selectors": "selector_id",
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_db(path: str | Path) -> sqlite3.Connection:
    """Open (or create) a file-based SQLite DB with the GW schema applied.

    Caller is responsible for closing the connection.
    """
    conn = get_connection(path)
    initialize_schema(conn)
    return conn


def merge_user_db(
    user_source: "str | Path | sqlite3.Connection",
    master_source: "str | Path | sqlite3.Connection",
    dry_run: bool = False,
) -> dict:
    """Merge user DB records into the master DB.

    For each table (targets, selectors):
      - Records in user but NOT in master → inserted into master (unless dry_run).
      - Records in BOTH with different last_updated → flagged as conflict, not written.
      - Records in BOTH with identical last_updated → skipped (already in sync).

    Parameters
    ----------
    user_source:
        Path to the user's DB file, or an open sqlite3.Connection.
    master_source:
        Path to the master DB file, or an open sqlite3.Connection.
    dry_run:
        If True, report what would happen but make no changes to master.

    Returns
    -------
    dict with keys:
        "inserted":  {"selectors": int, "targets": int}
        "conflicts": list of conflict dicts
        "dry_run":   bool
    """
    user_conn, user_opened = _open_source(user_source)
    master_conn, master_opened = _open_source(master_source)

    inserted: dict[str, int] = {"norks": 0, "targets": 0, "selectors": 0}
    conflicts: list[dict] = []

    try:
        for table, pk_col in _MERGE_TABLES.items():
            table_inserted, table_conflicts = _merge_table(
                user_conn, master_conn, table, pk_col, dry_run
            )
            inserted[table] += table_inserted
            conflicts.extend(table_conflicts)

        if not dry_run:
            master_conn.commit()
    finally:
        if user_opened:
            user_conn.close()
        if master_opened:
            master_conn.close()

    result = {"inserted": inserted, "conflicts": conflicts, "dry_run": dry_run}
    _log_conflicts(conflicts, dry_run)
    return result


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

def _open_source(source: "str | Path | sqlite3.Connection") -> tuple[sqlite3.Connection, bool]:
    """Return (conn, we_opened_it)."""
    if isinstance(source, sqlite3.Connection):
        return source, False
    conn = get_connection(source)
    return conn, True


def _merge_table(
    user_conn: sqlite3.Connection,
    master_conn: sqlite3.Connection,
    table: str,
    pk_col: str,
    dry_run: bool,
) -> tuple[int, list[dict]]:
    """Merge one table. Returns (inserted_count, conflict_list)."""
    user_rows = {
        row[pk_col]: {k: row[k] for k in row.keys()}
        for row in user_conn.execute(f"SELECT * FROM {table}").fetchall()  # noqa: S608
    }

    if not user_rows:
        return 0, []

    master_rows = {
        row[pk_col]: {k: row[k] for k in row.keys()}
        for row in master_conn.execute(f"SELECT * FROM {table}").fetchall()  # noqa: S608
    }

    inserted = 0
    conflicts: list[dict] = []

    for pk_val, user_row in user_rows.items():
        if pk_val not in master_rows:
            # New record — insert into master.
            if not dry_run:
                cols = list(user_row.keys())
                placeholders = ", ".join("?" * len(cols))
                col_list = ", ".join(cols)
                master_conn.execute(
                    f"INSERT OR IGNORE INTO {table} ({col_list}) VALUES ({placeholders})",  # noqa: S608
                    list(user_row.values()),
                )
            inserted += 1
        else:
            master_row = master_rows[pk_val]
            if user_row.get("last_updated") != master_row.get("last_updated"):
                conflicts.append({
                    "table": table,
                    "id": pk_val,
                    "user_updated": user_row.get("last_updated"),
                    "master_updated": master_row.get("last_updated"),
                })
            # Identical last_updated → already in sync, skip silently.

    return inserted, conflicts


def _log_conflicts(conflicts: list[dict], dry_run: bool) -> None:
    if not conflicts:
        return

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        log_dir = config.LOG_DIR
        log_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        log_dir = Path(__file__).parent.parent / "logs"
        log_dir.mkdir(exist_ok=True)

    conflict_log = log_dir / f"merge_conflicts_{ts}.log"
    prefix = "[DRY RUN] " if dry_run else ""

    with conflict_log.open("w", encoding="utf-8") as fh:
        fh.write(f"{prefix}Merge conflict report — {ts}\n")
        fh.write("=" * 60 + "\n\n")
        for c in conflicts:
            line = (
                f"Table: {c['table']}  ID: {c['id']}\n"
                f"  User last_updated:   {c['user_updated']}\n"
                f"  Master last_updated: {c['master_updated']}\n\n"
            )
            fh.write(line)
            logger.warning("%sConflict — %s", prefix, line.replace("\n", " "))

    print(f"{prefix}{len(conflicts)} conflict(s) written to: {conflict_log}")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge a user's GrayWolfe DB into the master DB."
    )
    parser.add_argument("user_db", help="Path to the user's local .db file")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report what would happen without writing to master",
    )
    parser.add_argument(
        "--master-db",
        default=str(config.MASTER_DB_PATH),
        help=f"Path to master DB (default: {config.MASTER_DB_PATH})",
    )
    args = parser.parse_args()

    result = merge_user_db(
        user_source=args.user_db,
        master_source=args.master_db,
        dry_run=args.dry_run,
    )

    prefix = "[DRY RUN] " if result["dry_run"] else ""
    print(f"\n{prefix}Merge complete:")
    print(f"  Inserted: {result['inserted']}")
    print(f"  Conflicts: {len(result['conflicts'])}")


if __name__ == "__main__":
    main()
