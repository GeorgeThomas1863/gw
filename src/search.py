"""
core/search.py — Search orchestration.

Translates VBA mod1a_RunSearch.bas to Python.
"""
from __future__ import annotations

import logging
import sqlite3
from typing import Callable

from models import GWResult, SApiResult, SelectorType
from src.type_engine import build_selector_clean, detect_selector_type
from util.errors import ERR_EMPTY_SEARCH, GWError

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# parse_raw_input
# ---------------------------------------------------------------------------

def parse_raw_input(raw: str, delimiter: str | None = None) -> list[str]:
    """Split raw input into individual selector values.

    Delimiter detection order (if delimiter not specified):
    1. Tab
    2. Newline
    3. Comma (only if no other delimiter found)
    Strips whitespace from each value; filters empty strings.
    Returns list of clean selector strings.
    """
    if not raw or not raw.strip():
        return []

    if delimiter is not None:
        parts = raw.split(delimiter)
    else:
        if "\t" in raw:
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

def search_graywolfe(selectors: list[str], conn: sqlite3.Connection) -> list[GWResult]:
    """Search local DB for each selector.

    For each value in selectors:
    1. Compute selector_clean via build_selector_clean(value, detect_selector_type(value))
    2. Query: SELECT s.*, t.target_name FROM selectors s
               LEFT JOIN targets t ON s.target_id = t.target_id
               WHERE s.selector_clean = ?
    3. If found: result has in_gray_wolfe=True, all selector + target fields
    4. If not found: result has selector=value, in_gray_wolfe=False, all other fields None

    Returns list of GWResult. One row per (selector, target) pair found.
    The original query term is preserved as 'query_value' in each result.
    """
    if not selectors:
        return []

    results: list[GWResult] = []

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
                _raw_type = row["selector_type"]
                _sel_type: SelectorType | None
                if _raw_type:
                    try:
                        _sel_type = SelectorType(_raw_type)
                    except ValueError:
                        _sel_type = None
                else:
                    _sel_type = None

                results.append(GWResult(
                    query_value=value,
                    selector=row["selector"],
                    selector_id=row["selector_id"],
                    selector_clean=row["selector_clean"],
                    selector_type=_sel_type,
                    target_id=row["target_id"],
                    nork_id=row["nork_id"],
                    date_created=row["date_created"],
                    created_by=row["created_by"],
                    last_updated=row["last_updated"],
                    last_updated_by=row["last_updated_by"],
                    data_source=row["data_source"],
                    target_name=row["target_name"],
                    in_gray_wolfe=True,
                ))
        else:
            results.append(GWResult(
                query_value=value,
                selector=value,
                selector_id=None,
                selector_clean=None,
                selector_type=None,
                target_id=None,
                nork_id=None,
                date_created=None,
                created_by=None,
                last_updated=None,
                last_updated_by=None,
                data_source=None,
                target_name=None,
                in_gray_wolfe=False,
            ))

    return results


# ---------------------------------------------------------------------------
# search_s
# ---------------------------------------------------------------------------

def search_s(
    selectors: list[str],
    s_client,
    progress_cb: Callable[[str, int, int], None] | None = None,
    ask_cb: Callable[[str, int], bool] | None = None,
) -> list[SApiResult]:
    """Call SApiClient.search for each selector, aggregate results.

    Validates the token against the live S API before issuing any queries.
    Returns flat list. On per-selector error, logs warning and continues.

    Args:
        selectors: List of selector values to search.
        s_client: SApiClient instance.
        progress_cb: Optional callback(selector, idx, total) for rate limit progress.
        ask_cb: Optional callback(selector, num_found) for continuing on rate limit.
    """
    if not selectors:
        return []

    s_client.validate_token()

    results: list[SApiResult] = []
    for i, value in enumerate(selectors):
        def _rate_cb(v=value, idx=i, tot=len(selectors)):
            if progress_cb is not None:
                progress_cb(v, idx, tot)

        def _ask(q, n):
            return ask_cb(q, n) if ask_cb is not None else True

        try:
            batch = s_client.search(value, on_rate_limit=_rate_cb, ask_continue=_ask)
            results.extend(batch)
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
    progress_cb: Callable[[str, int, int], None] | None = None,
    ask_cb: Callable[[str, int], bool] | None = None,
) -> tuple[list[GWResult], list[SApiResult]]:
    """Main search entry point.

    Returns (gw_results, s_results). Either list may be empty if flag is False
    or s_client is None.
    Raises GWError(ERR_EMPTY_SEARCH) if raw_input is blank after parsing.

    Args:
        raw_input: Raw search input string.
        delimiter: Optional delimiter; auto-detected if None.
        conn: SQLite connection for GrayWolfe search.
        s_client: Optional SApiClient for S API search.
        search_gw: Whether to search GrayWolfe local DB.
        search_s_flag: Whether to search S API.
        progress_cb: Optional callback(selector, idx, total) for rate limit progress.
        ask_cb: Optional callback(selector, num_found) for continuing on rate limit.
    """
    selectors = parse_raw_input(raw_input, delimiter)

    if not selectors:
        raise GWError(ERR_EMPTY_SEARCH, "Search input is empty after parsing.")

    gw_results: list[GWResult] = []
    s_results: list[SApiResult] = []

    if search_gw:
        gw_results = search_graywolfe(selectors, conn)

    if search_s_flag and s_client is not None:
        s_results = search_s(selectors, s_client, progress_cb, ask_cb)

    return gw_results, s_results
