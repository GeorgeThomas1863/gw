from __future__ import annotations

# ---------------------------------------------------------------------------
# Error codes — mirror VBA mod4c_ErrorHandler.bas codes exactly
# ---------------------------------------------------------------------------
ERR_EMPTY_INPUT: int = 1950
ERR_EMPTY_SEARCH: int = 1951
ERR_DB_CONNECTION: int = 1952
ERR_DB_INSERT: int = 1953
ERR_DB_UPDATE: int = 1954
ERR_DB_DELETE: int = 1955
ERR_INVALID_SELECTOR: int = 1956
ERR_INVALID_TARGET_ID: int = 1957
ERR_TARGET_NOT_FOUND: int = 1958
ERR_SELECTOR_DUPLICATE: int = 1959
ERR_SCHEMA_EMPTY: int = 1960
ERR_SCHEMA_COL_MISMATCH: int = 1961
ERR_SPLIT_ARRAY_EMPTY: int = 1962
ERR_IMPORT_FAILED: int = 1963
ERR_SEARCH_FAILED: int = 1964
ERR_S_TOKEN_TOO_SHORT: int = 1965   # token present but below minimum length
ERR_MISSING_S_TOKEN: int = 1966
ERR_S_TOKEN_WRONG_FORMAT: int = 1967
ERR_S_AUTH_FAILED: int = 1968
ERR_S_REQUEST_FAILED: int = 1969
ERR_S_PARSE_FAILED: int = 1970
ERR_SYNC_FAILED: int = 1971
ERR_MERGE_FAILED: int = 1972


# ---------------------------------------------------------------------------
# Exception class
# ---------------------------------------------------------------------------
class GWError(Exception):
    """Application-level exception carrying a GrayWolfe error code."""

    def __init__(self, code: int, message: str) -> None:
        super().__init__(message)
        self.code: int = code
        self.message: str = message

    def __str__(self) -> str:
        return f"[GW{self.code}] {self.message}"


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------
def raise_gw(code: int, message: str) -> None:
    """Raise a GWError with the given code and message."""
    raise GWError(code, message)
