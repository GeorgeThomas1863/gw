"""
core/search.py — Search orchestration.

Translates VBA mod1a_RunSearch.bas to Python.
"""
from __future__ import annotations

import logging
import sqlite3

from config import INTERNAL_DELIM
from core.type_engine import build_selector_clean, detect_selector_type
from utils.errors import ERR_EMPTY_SEARCH, GWError

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# parse_raw_input
# ---------------------------------------------------------------------------

def parse_raw_input(raw: str, delimiter: str | None = None) -> list[str]:
    """Split raw input into individual selector values.

    Delimiter detection order (if delimiter not specified):
    1. INTERNAL_DELIM ("!!")
    2. Tab
    3. Newline
    4. Comma (only if no other delimiter found)
    Strips whitespace from each value; filters empty strings.
    Returns list of clean selector strings.
    """
    if not raw or not raw.strip():
        return []

    if delimiter is not None:
        parts = raw.split(delimiter)
    else:
        if INTERNAL_DELIM in raw:
            parts = raw.split(INTERNAL_DELIM)
        elif "\t" in raw:
            parts = raw.split("\t")
        elif "\n" in raw:
            parts = raw.split("\n")
        elif "," in raw:
            parts = raw.split(",")
        else:
            parts = [raw]

    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# search_graywolfe
# ---------------------------------------------------------------------------

def search_graywolfe(selectors: list[str], conn: sqlite3.Connection) -> list[dict]:
    """Search local DB for each selector.

    For each value in selectors:
    1. Compute selector_clean via build_selector_clean(value, detect_selector_type(value))
    2. Query: SELECT s.*, t.target_name FROM selectors s
               LEFT JOIN targets t ON s.target_id = t.target_id
               WHERE s.selector_clean = ?
    3. If found: result has in_gray_wolfe=True, all selector + target fields
    4. If not found: result has selector=value, in_gray_wolfe=False, all other fields None

    Returns list of result dicts. One row per (selector, target) pair found.
    The original query term is preserved as 'query_value' in each result.
    """
    if not selectors:
        return []

    results: list[dict] = []

    for value in selectors:
        sel_type = detect_selector_type(value)
        clean = build_selector_clean(value, sel_type)

        if clean:
            rows = conn.execute(
                """
                SELECT s.selector_id, s.selector, s.selector_clean, s.selector_type,
                       s.target_id, s.nork_id,
                       s.date_created, s.created_by,
                       s.last_updated, s.last_updated_by,
                       s.data_source,
                       t.target_name
                FROM selectors s
                LEFT JOIN targets t ON s.target_id = t.target_id
                WHERE s.selector_clean = ?
                """,
                (clean,),
            ).fetchall()
        else:
            rows = []

        if rows:
            for row in rows:
                result = dict(row) if hasattr(row, "keys") else {
                    "selector_id": row[0],
                    "selector": row[1],
                    "selector_clean": row[2],
                    "selector_type": row[3],
                    "target_id": row[4],
                    "nork_id": row[5],
                    "date_created": row[6],
                    "created_by": row[7],
                    "last_updated": row[8],
                    "last_updated_by": row[9],
                    "data_source": row[10],
                    "target_name": row[11],
                }
                result["in_gray_wolfe"] = True
                result["query_value"] = value
                results.append(result)
        else:
            results.append({
                "query_value": value,
                "selector": value,
                "selector_id": None,
                "selector_clean": None,
                "selector_type": None,
                "target_id": None,
                "nork_id": None,
                "date_created": None,
                "created_by": None,
                "last_updated": None,
                "last_updated_by": None,
                "data_source": None,
                "target_name": None,
                "in_gray_wolfe": False,
            })

    return results


# ---------------------------------------------------------------------------
# search_s
# ---------------------------------------------------------------------------

def search_s(selectors: list[str], s_client) -> list[dict]:
    """Call SApiClient.search for each selector, aggregate results.

    Validates the token against the live S API before issuing any queries.
    Each result dict has a 'selector' field with the search term.
    Returns flat list. On per-selector error, logs warning and continues.
    """
    if not selectors:
        return []

    s_client.validate_token()

    results: list[dict] = []
    for value in selectors:
        try:
            batch = s_client.search(value)
            for item in batch:
                item["selector"] = value
                results.append(item)
        except Exception as exc:  # noqa: BLE001
            logger.warning("search_s: error searching %r: %s", value, exc)

    return results


# ---------------------------------------------------------------------------
# run_search
# ---------------------------------------------------------------------------

def run_search(
    raw_input: str,
    delimiter: str | None,
    conn: sqlite3.Connection,
    s_client=None,
    search_gw: bool = True,
    search_s_flag: bool = True,
) -> tuple[list[dict], list[dict]]:
    """Main search entry point.

    Returns (gw_results, s_results). Either list may be empty if flag is False
    or s_client is None.
    Raises GWError(ERR_EMPTY_SEARCH) if raw_input is blank after parsing.
    """
    selectors = parse_raw_input(raw_input, delimiter)

    if not selectors:
        raise GWError(ERR_EMPTY_SEARCH, "Search input is empty after parsing.")

    gw_results: list[dict] = []
    s_results: list[dict] = []

    if search_gw:
        gw_results = search_graywolfe(selectors, conn)

    if search_s_flag and s_client is not None:
        s_results = search_s(selectors, s_client)

    return gw_results, s_results
