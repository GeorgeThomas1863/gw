"""
data/s_api.py — S API client.

Translates VBA mod2b_S.bas + mod4a_APICalls.bas to Python using requests.
"""
from __future__ import annotations

import time
import webbrowser

import requests

from config import (
    S_BATCH_SIZE,
    S_LINK_TEMPLATE,
    S_RATE_LIMIT_SLEEP,
    S_SEARCH_URL,
    S_TOKEN_MIN_LEN,
    S_VALIDATE_URLS,
)
from utils.errors import (
    ERR_MISSING_S_TOKEN,
    ERR_S_AUTH_FAILED,
    ERR_S_PARSE_FAILED,
    ERR_S_REQUEST_FAILED,
    ERR_S_TOKEN_TOO_SHORT,
    ERR_S_TOKEN_WRONG_FORMAT,
    GWError,
    raise_gw,
)


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
        Raises GWError(ERR_S_AUTH_FAILED) if all return 403/401.
        Raises GWError(ERR_S_REQUEST_FAILED) on network error.
        """
        any_success = False
        try:
            for entry in S_VALIDATE_URLS:
                method = entry["method"].upper()
                url = entry["url"]
                if method == "GET":
                    response = self.session.get(url)
                else:
                    response = self.session.post(url)
                if response.status_code < 300:
                    any_success = True
        except requests.exceptions.RequestException as exc:
            raise_gw(ERR_S_REQUEST_FAILED, f"Network error during token validation: {exc}")

        if not any_success:
            raise_gw(ERR_S_AUTH_FAILED, "S API authentication failed — all validation endpoints rejected the token.")

        return True

    def search(self, query: str) -> list[dict]:
        """Search S API for query. Handles pagination (500/batch).

        Sleeps S_RATE_LIMIT_SLEEP seconds between batches (not after last batch).
        Returns flat list of parsed result dicts.
        Raises GWError on error.
        """
        results: list[dict] = []
        start = 0

        first_batch = self._search_batch(query, start)
        num_found: int = first_batch.get("numFound", 0)

        for item in first_batch.get("items", []):
            results.append(self._parse_item(item, query))

        start += S_BATCH_SIZE

        while start < num_found:
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
        try:
            response = self.session.post(
                S_SEARCH_URL,
                json=payload,
                headers={"Content-Type": "application/json"},
            )
        except requests.exceptions.RequestException as exc:
            raise_gw(ERR_S_REQUEST_FAILED, f"Network error during S search: {exc}")

        try:
            return response.json()
        except ValueError as exc:
            raise_gw(ERR_S_PARSE_FAILED, f"Failed to parse S API response as JSON: {exc}")

    def _parse_item(self, item: dict, query: str) -> dict:
        """Extract fields from one S API result item.

        Returns dict with keys: s_id, selector, doc_id, doc_type, doc_sub_type,
        case, serial, case_serial_full, office, doc_title, author, created_date, link
        """
        unique_id = item.get("uniqueID", "")
        case = item.get("UCFN", "")
        serial = item.get("itemNumber", "")
        return {
            "s_id": unique_id,
            "selector": query,
            "doc_id": unique_id,
            "doc_type": item.get("recordType", ""),
            "doc_sub_type": item.get("recordSubType", ""),
            "case": case,
            "serial": serial,
            "case_serial_full": f"{case}/{serial}",
            "office": item.get("caseOfficeCode", ""),
            "doc_title": item.get("title", ""),
            "author": item.get("primaryAuthor", ""),
            "created_date": item.get("createdDate", ""),
            "link": self.get_link(unique_id),
        }

    def get_link(self, unique_id: str) -> str:
        """Return the S document URL for a given unique_id."""
        return S_LINK_TEMPLATE.format(unique_id=unique_id)

    @staticmethod
    def open_link(unique_id: str) -> None:
        """Open S document link in default browser via webbrowser.open."""
        url = S_LINK_TEMPLATE.format(unique_id=unique_id)
        webbrowser.open(url)
