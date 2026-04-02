"""
core/import_data.py — Import workflows.

Translates VBA mod1b_RunAddData.bas to Python.
"""
from __future__ import annotations

import logging
import sqlite3

from config import INTERNAL_DELIM, ROW_DELIM, SELECTOR_TYPES
from core.targets import collect_target_ids_for_row, create_target, merge_targets
from core.type_engine import build_selector_clean, detect_selector_type
from data.database import generate_id, get_current_user, insert_selector, now_iso
from utils.errors import ERR_SELECTOR_DUPLICATE, GWError

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# parse_import_input
# ---------------------------------------------------------------------------

def parse_import_input(raw: str, delimiter: str | None = None) -> list[list[str]]:
    """Parse multi-column import input into list of rows.

    Row delimiter: ROW_DELIM ("+++") or newline
    Column delimiter: INTERNAL_DELIM ("!!") or tab
    Strips and filters empty values.
    Returns list of rows, each row is a list of column values.
    """
    if not raw or not raw.strip():
        return []

    # Split into rows
    if ROW_DELIM in raw:
        raw_rows = raw.split(ROW_DELIM)
    else:
        raw_rows = raw.split("\n")

    # Determine column delimiter
    if delimiter is not None:
        col_delim = delimiter
    elif INTERNAL_DELIM in raw:
        col_delim = INTERNAL_DELIM
    elif "\t" in raw:
        col_delim = "\t"
    else:
        col_delim = None

    rows: list[list[str]] = []
    for raw_row in raw_rows:
        raw_row = raw_row.strip()
        if not raw_row:
            continue
        if col_delim is not None:
            cols = [c.strip() for c in raw_row.split(col_delim) if c.strip()]
        else:
            cols = [raw_row]
        if cols:
            rows.append(cols)

    return rows


# ---------------------------------------------------------------------------
# detect_column_types
# ---------------------------------------------------------------------------

def detect_column_types(rows: list[list[str]]) -> list[dict]:
    """For each column position across all rows, count how many values match
    each selector type.

    Returns list of:
    {"type": best_type, "confidence": float (0.0–1.0), "col_index": int}

    One dict per column. Columns with all-null/empty values get type="other",
    confidence=0.0.
    """
    if not rows:
        return []

    col_count = max(len(row) for row in rows)
    results: list[dict] = []

    for col_idx in range(col_count):
        values = [row[col_idx] for row in rows if col_idx < len(row) and row[col_idx].strip()]

        if not values:
            results.append({"type": "other", "confidence": 0.0, "col_index": col_idx})
            continue

        # Count matches per type
        counts: dict[str, int] = {t: 0 for t in SELECTOR_TYPES}
        for val in values:
            detected = detect_selector_type(val)
            counts[detected] = counts.get(detected, 0) + 1

        best_type = max(counts, key=lambda t: counts[t])
        confidence = counts[best_type] / len(values) if values else 0.0

        results.append({"type": best_type, "confidence": confidence, "col_index": col_idx})

    return results


# ---------------------------------------------------------------------------
# _make_selector_dict
# ---------------------------------------------------------------------------

def _make_selector_dict(
    value: str,
    sel_type: str,
    username: str,
) -> dict | None:
    """Build a selector dict ready for insert_selector.

    Returns None if the cleaned value is empty (unusable).
    """
    clean = build_selector_clean(value, sel_type)
    if not clean:
        return None

    ts = now_iso()
    return {
        "selector_id": generate_id(),
        "selector": value,
        "selector_clean": clean,
        "selector_type": sel_type,
        "target_id": None,
        "nork_id": None,
        "date_created": ts,
        "created_by": username,
        "last_updated": ts,
        "last_updated_by": username,
        "data_source": None,
    }


# ---------------------------------------------------------------------------
# run_unrelated_import
# ---------------------------------------------------------------------------

def run_unrelated_import(
    raw_input: str,
    sel_type: str,
    conn: sqlite3.Connection,
    username: str | None = None,
    delimiter: str | None = None,
) -> int:
    """Import each selector independently (no target linking).

    sel_type is either a specific type or "auto" (auto-detect each value).
    Skips duplicates (logs warning, does not raise).
    Returns count of successfully inserted selectors.
    """
    if username is None:
        username = get_current_user()

    from core.search import parse_raw_input
    values = parse_raw_input(raw_input, delimiter)

    inserted = 0
    for value in values:
        effective_type = detect_selector_type(value) if sel_type == "auto" else sel_type
        sel_dict = _make_selector_dict(value, effective_type, username)
        if sel_dict is None:
            logger.warning("run_unrelated_import: skipping value %r — could not clean", value)
            continue

        try:
            insert_selector(conn, sel_dict)
            inserted += 1
        except GWError as exc:
            if exc.code == ERR_SELECTOR_DUPLICATE:
                logger.warning("run_unrelated_import: duplicate skipped — %s", exc)
            else:
                raise

    return inserted


# ---------------------------------------------------------------------------
# run_default_import
# ---------------------------------------------------------------------------

def run_default_import(
    rows: list[list[str]],
    confirmed_types: list[str],
    conn: sqlite3.Connection,
    username: str | None = None,
) -> int:
    """Process each row: insert selectors, resolve targets via 3-case logic.

    For each row:
    1. For each cell (value, type pair where type != "null"):
       - detect_selector_type if type is "auto"; else use confirmed type
       - fix/clean the value
       - insert_selector (skip if duplicate — log warning)
    2. collect_target_ids_for_row → 0/1/2+ targets
       - 0: create_target, link all row selectors to new target
       - 1: link all row selectors to existing target
       - 2+: merge_targets (fold all into first), link all row selectors to surviving target

    A row with only 1 column does NOT do target linking (mirrors VBA colMax >= 2).
    Returns total count of inserted selectors.
    """
    if username is None:
        username = get_current_user()

    total_inserted = 0

    for row in rows:
        # Track the selector_clean values inserted this round (for linking)
        inserted_cleans: list[str] = []
        # Track all selector values in this row (for target collection)
        row_values: list[str] = []

        for col_idx, value in enumerate(row):
            if col_idx >= len(confirmed_types):
                break

            col_type = confirmed_types[col_idx]
            if col_type == "null":
                continue

            value = value.strip()
            if not value:
                continue

            effective_type = detect_selector_type(value) if col_type == "auto" else col_type
            sel_dict = _make_selector_dict(value, effective_type, username)
            if sel_dict is None:
                logger.warning("run_default_import: skipping %r — empty after clean", value)
                continue

            row_values.append(value)

            try:
                insert_selector(conn, sel_dict)
                inserted_cleans.append(sel_dict["selector_clean"])
                total_inserted += 1
            except GWError as exc:
                if exc.code == ERR_SELECTOR_DUPLICATE:
                    logger.warning("run_default_import: duplicate skipped — %s", exc)
                    # row_values already has this value (appended above); do not append again
                else:
                    raise

        # Single-column rows skip target linking (VBA: colMax >= 2)
        if len(row) < 2:
            continue

        if not row_values:
            continue

        # Resolve targets using the full row values (includes duplicates already in DB)
        target_ids = collect_target_ids_for_row(row_values, conn)

        if len(target_ids) == 0:
            # Case 0: no existing target → create new one
            new_tid = create_target(conn, username=username)
            _link_selectors_to_target(conn, inserted_cleans, new_tid)
        elif len(target_ids) == 1:
            # Case 1: exactly one existing target → link to it
            _link_selectors_to_target(conn, inserted_cleans, target_ids[0])
        else:
            # Case 2+: multiple targets → merge all into the first, link to survivor
            survivor_id = target_ids[0]
            for absorb_id in target_ids[1:]:
                merge_targets(survivor_id, absorb_id, conn, username=username)
            _link_selectors_to_target(conn, inserted_cleans, survivor_id)

    return total_inserted


def _link_selectors_to_target(
    conn: sqlite3.Connection,
    selector_cleans: list[str],
    target_id: str,
) -> None:
    """UPDATE selectors SET target_id = ? WHERE selector_clean IN (...)
    AND target_id IS NULL — only assigns selectors that have no existing target.

    Processes in chunks of 500 to stay under SQLite's 999-variable limit.
    """
    if not selector_cleans:
        return

    chunk_size = 500
    for start in range(0, len(selector_cleans), chunk_size):
        chunk = selector_cleans[start : start + chunk_size]
        placeholders = ",".join("?" * len(chunk))
        conn.execute(
            f"UPDATE selectors SET target_id = ? WHERE selector_clean IN ({placeholders}) AND target_id IS NULL",
            [target_id, *chunk],
        )
    conn.commit()
