"""
core/targets.py — Target lifecycle management.

Translates VBA MergeTargets() and CollectTargetIdsForRow() from
mod2a_Sharepoint_Tables.bas into Python.
"""
from __future__ import annotations

import sqlite3

from data.database import (
    generate_id,
    get_current_user,
    now_iso,
    insert_target,
    get_target,
    update_target as _db_update_target,
    delete_target,
)
from src.type_engine import build_selector_clean, detect_selector_type
from util.errors import (
    GWError,
    raise_gw,
    ERR_MERGE_FAILED,
    ERR_TARGET_NOT_FOUND,
    ERR_INVALID_TARGET_ID,
)


# ---------------------------------------------------------------------------
# collect_target_ids_for_row
# ---------------------------------------------------------------------------

def collect_target_ids_for_row(selectors: list[str], conn: sqlite3.Connection) -> list[str]:
    """For each selector value in the list, compute selector_clean and query
    localSelectors for any non-null target_ids.

    Returns a deduplicated, insertion-stable list of distinct target_ids found.
    """
    seen: dict[str, None] = {}  # ordered set via dict keys

    for value in selectors:
        sel_type = detect_selector_type(value)
        clean = build_selector_clean(value, sel_type)
        if not clean:
            continue

        rows = conn.execute(
            """
            SELECT DISTINCT target_id
            FROM selectors
            WHERE selector_clean = ?
              AND target_id IS NOT NULL
              AND target_id != ''
            """,
            (clean,),
        ).fetchall()

        for row in rows:
            tid = row[0] if not hasattr(row, "keys") else row["target_id"]
            seen[tid] = None

    return list(seen.keys())


# ---------------------------------------------------------------------------
# create_target
# ---------------------------------------------------------------------------

def create_target(conn: sqlite3.Connection, username: str | None = None) -> str:
    """Create a new target with a generated ID.

    Default target_name: "[New Target]"
    Returns target_id.
    """
    if username is None:
        username = get_current_user()

    tid = generate_id()
    ts = now_iso()

    target = {
        "target_id": tid,
        "target_name": "[New Target]",
        "case_number": None,
        "laptop_count": 0,
        "date_created": ts,
        "created_by": username,
        "last_updated": ts,
        "last_updated_by": username,
        "data_source": None,
    }
    insert_target(conn, target)
    return tid


# ---------------------------------------------------------------------------
# merge_targets
# ---------------------------------------------------------------------------

def merge_targets(
    keep_id: str,
    absorb_id: str,
    conn: sqlite3.Connection,
    username: str | None = None,
) -> bool:
    """Merge absorb_id into keep_id.

    Steps:
    1. Validate both IDs are non-empty and different.
    2. Fetch both targets; raise if either not found.
    3. UPDATE selectors SET target_id=keep WHERE target_id=absorb.
    4. DELETE duplicate selector_cleans within keep (keep MIN(id)).
    5. If keep.target_name is blank/None and absorb has one → copy absorb's name.
    6. If keep.case_number is blank/None and absorb has one → copy absorb's case_number.
    7. Sum laptop_count: keep += absorb (only if absorb > 0).
    8. DELETE absorbed target from targets table.
    9. Update keep's last_updated + last_updated_by.

    Returns True on success.
    Raises GWError on validation failure or missing records.
    """
    if username is None:
        username = get_current_user()

    # --- Step 1: Validate IDs ---
    if not keep_id or not keep_id.strip():
        raise GWError(ERR_INVALID_TARGET_ID, f"keep_id is empty or blank")
    if not absorb_id or not absorb_id.strip():
        raise GWError(ERR_INVALID_TARGET_ID, f"absorb_id is empty or blank")
    if keep_id.strip() == absorb_id.strip():
        raise GWError(ERR_MERGE_FAILED, "keep_id and absorb_id must be different")

    # --- Step 2: Fetch both targets ---
    keep = get_target(conn, keep_id)
    if keep is None:
        raise GWError(ERR_TARGET_NOT_FOUND, f"Keep target not found: {keep_id!r}")

    absorb = get_target(conn, absorb_id)
    if absorb is None:
        raise GWError(ERR_TARGET_NOT_FOUND, f"Absorb target not found: {absorb_id!r}")

    # --- Step 3: Move absorb's selectors to keep ---
    conn.execute(
        "UPDATE selectors SET target_id = ? WHERE target_id = ?",
        (keep_id, absorb_id),
    )

    # --- Step 4: Delete duplicate selector_cleans within keep (keep MIN(id)) ---
    conn.execute(
        """
        DELETE FROM selectors
        WHERE target_id = ?
          AND id NOT IN (
              SELECT MIN(id)
              FROM selectors
              WHERE target_id = ?
              GROUP BY selector_clean
          )
        """,
        (keep_id, keep_id),
    )
    conn.commit()

    # --- Steps 5-7: Compute metadata updates for keep ---
    fields_to_update: dict = {}

    keep_name = keep["target_name"] or ""
    absorb_name = absorb["target_name"] or ""
    if not keep_name.strip() and absorb_name.strip():
        fields_to_update["target_name"] = absorb_name

    keep_case = keep["case_number"] or ""
    absorb_case = absorb["case_number"] or ""
    if not keep_case.strip() and absorb_case.strip():
        fields_to_update["case_number"] = absorb_case

    absorb_laptops = absorb["laptop_count"] or 0
    if absorb_laptops > 0:
        keep_laptops = keep["laptop_count"] or 0
        fields_to_update["laptop_count"] = keep_laptops + absorb_laptops

    # --- Step 9: Always refresh last_updated on keep ---
    # update_target always stamps last_updated/last_updated_by, so pass empty
    # fields dict if nothing else changed — it will still update the timestamp.
    update_target(keep_id, fields_to_update, conn, username)

    # --- Step 8: Delete the absorbed target ---
    delete_target(conn, absorb_id)

    return True


# ---------------------------------------------------------------------------
# update_target (thin wrapper that keeps callers in core.targets)
# ---------------------------------------------------------------------------

def update_target(
    target_id: str,
    fields: dict,
    conn: sqlite3.Connection,
    username: str | None = None,
) -> None:
    """Update fields on a target.  Thin wrapper around data.database.update_target
    with argument order matching the rest of this module (target_id first, conn last).
    """
    if username is None:
        username = get_current_user()

    _db_update_target(conn, target_id, fields, username)
