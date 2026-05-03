"""
data/s_api.py — S API client.

Translates VBA mod2b_S.bas + mod4a_APICalls.bas to Python using requests.
"""
from __future__ import annotations

import time
import webbrowser
from collections.abc import Callable

import requests

from util.logger import get_logger
from models import SApiResult
from config import (
    S_BATCH_SIZE,
    S_LINK_TEMPLATE,
    S_RATE_LIMIT_SLEEP,
    S_SEARCH_URL,
    S_TOKEN_MIN_LEN,
    S_VALIDATE_URLS,
)
from util.errors import (
    ERR_MISSING_S_TOKEN,
    ERR_S_AUTH_FAILED,
    ERR_S_PARSE_FAILED,
    ERR_S_REQUEST_FAILED,
    ERR_S_TOKEN_TOO_SHORT,
    ERR_S_TOKEN_WRONG_FORMAT,
    GWError,
    raise_gw,
)

_log = get_logger(__name__)


class SApiClient:
    """Client for the S search API."""

    def __init__(self, token: str) -> None:
        """Validate token format, create Session with Bearer auth header."""
        if not token:
            raise_gw(ERR_MISSING_S_TOKEN, "S API token is missing.")
        if len(token) < S_TOKEN_MIN_LEN:
            raise_gw(
                ERR_S_TOKEN_TOO_SHORT,
                f"S API token is too short (min {S_TOKEN_MIN_LEN} chars).",
            )
        if " " in token:
            raise_gw(ERR_S_TOKEN_WRONG_FORMAT, "S API token contains spaces.")

        self.token: str = token
        self.session: requests.Session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {token}"})

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def validate_token(self) -> bool:
        """Test token against all S_VALIDATE_URLS.

        Returns True if any endpoint returns 2xx.
        Returns False (via GWError) only after all URLs have been tried.
        Logs network errors per-URL and continues to the next URL.
        """
        print("S token validation: testing against live API...")
        any_success = False
        for entry in S_VALIDATE_URLS:
            method = entry["method"].upper()
            url = entry["url"]
            try:
                if method == "GET":
                    response = self.session.get(url)
                else:
                    response = self.session.post(url)
                if response.status_code < 300:
                    print(f"  PASS  {method} {url} -> HTTP {response.status_code}")
                    any_success = True
                else:
                    print(f"  FAIL  {method} {url} -> HTTP {response.status_code}")
            except requests.exceptions.RequestException as exc:
                print(f"  ERROR {method} {url} -> {exc}")
                _log.warning("validate_token: network error for %s %s: %s", method, url, exc)
                continue

        if not any_success:
            print("S token validation: FAILED — all endpoints rejected the token.")
            raise_gw(ERR_S_AUTH_FAILED, "S API authentication failed — all validation endpoints rejected the token.")

        print("S token validation: OK")
        return True

    def search(
        self,
        query: str,
        on_rate_limit: Callable[[], None] | None = None,
        ask_continue: Callable[[str, int], bool] | None = None,
    ) -> list[SApiResult]:
        """Search S API for query. Handles pagination (500/batch).

        Args:
            query: Search query string.
            on_rate_limit: Optional callback invoked before each rate-limit sleep.
            ask_continue: Optional callback invoked after first batch if more results exist.
                         Receives (query, num_found) and returns bool. If False, stops pagination.

        Behavior:
            - Fetches first batch and checks num_found.
            - If ask_continue is provided and num_found > S_BATCH_SIZE:
              calls ask_continue(query, num_found). If it returns False, returns first batch only.
            - For each subsequent batch: calls on_rate_limit() if provided, then sleeps.
            - Returns flat list of parsed result dicts.
            - Raises GWError on error.
        """
        results: list[SApiResult] = []
        start = 0

        first_batch = self._search_batch(query, start)
        num_found: int = first_batch.get("numFound", 0)

        for item in first_batch.get("items", []):
            results.append(self._parse_item(item, query))

        start += S_BATCH_SIZE

        # Check if user wants to continue pagination
        if num_found > S_BATCH_SIZE and ask_continue is not None:
            if not ask_continue(query, num_found):
                return results

        while start < num_found:
            if on_rate_limit is not None:
                on_rate_limit()
            time.sleep(S_RATE_LIMIT_SLEEP)
            batch = self._search_batch(query, start)
            for item in batch.get("items", []):
                results.append(self._parse_item(item, query))
            start += S_BATCH_SIZE

        return results

    def _search_batch(self, query: str, start: int) -> dict:
        """POST one search batch. Returns parsed JSON dict.

        Raises GWError(ERR_S_REQUEST_FAILED) on network error.
        Raises GWError(ERR_S_PARSE_FAILED) if response is not valid JSON.
        """
        payload = {
            "q": f'"{query}"',
            "limit": S_BATCH_SIZE,
            "start": start,
        }
        response = None
        try:
            response = self.session.post(
                S_SEARCH_URL,
                json=payload,
                headers={"Content-Type": "application/json"},
            )
        except requests.exceptions.RequestException as exc:
            raise_gw(ERR_S_REQUEST_FAILED, f"Network error during S search: {exc}")

        try:
            response.raise_for_status()
        except requests.exceptions.HTTPError as exc:
            if response.status_code in (401, 403):
                raise_gw(ERR_S_AUTH_FAILED,
                         f"S API authentication failed (HTTP {response.status_code}) — token may be expired.")
            raise_gw(ERR_S_REQUEST_FAILED,
                     f"S API returned HTTP {response.status_code}: {exc}")

        try:
            return response.json()
        except ValueError as exc:
            raise_gw(ERR_S_PARSE_FAILED, f"Failed to parse S API response as JSON: {exc}")

    def _parse_item(self, item: dict, query: str) -> SApiResult:
        """Extract fields from one S API result item.

        Returns SApiResult with fields: s_id, selector, doc_id, doc_type, doc_sub_type,
        case, serial, case_serial_full, office, doc_title, author, created_date, link
        """
        unique_id = item.get("uniqueID", "")
        case = item.get("UCFN", "")
        serial = item.get("itemNumber", "")
        return SApiResult(
            s_id=unique_id,
            selector=query,
            doc_id=unique_id,
            doc_type=item.get("recordType", ""),
            doc_sub_type=item.get("recordSubType", ""),
            case=case,
            serial=serial,
            case_serial_full=f"{case}/{serial}",
            office=item.get("caseOfficeCode", ""),
            doc_title=item.get("title", ""),
            author=item.get("primaryAuthor", ""),
            created_date=item.get("createdDate", ""),
            link=self.get_link(unique_id),
        )

    def get_link(self, unique_id: str) -> str:
        """Return the S document URL for a given unique_id."""
        return S_LINK_TEMPLATE.format(unique_id=unique_id)

    @staticmethod
    def open_link(unique_id: str) -> None:
        """Open S document link in default browser via webbrowser.open."""
        url = S_LINK_TEMPLATE.format(unique_id=unique_id)
        webbrowser.open(url)
