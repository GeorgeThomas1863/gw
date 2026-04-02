from pathlib import Path

# ---------------------------------------------------------------------------
# Shared paths (UNC — placeholder; replace with real server share)
# ---------------------------------------------------------------------------
GW_SHARED_PATH: Path = Path(r"\\server\graywolfe")
MASTER_DB_PATH: Path = GW_SHARED_PATH / "master.db"
USER_DB_DIR: Path = GW_SHARED_PATH / "users"
LOG_DIR: Path = GW_SHARED_PATH / "logs"

# ---------------------------------------------------------------------------
# S API
# ---------------------------------------------------------------------------
S_SEARCH_URL: str = "https://S-api.Fnet.F/services/search/api/search/v1"

# Validation endpoints. Most are GET; the last is a POST — method is explicit.
S_VALIDATE_URLS: list[dict[str, str]] = [
    {"method": "GET",  "url": "https://S-api.Fnet.F/services/externalservice/api/Lookups/Countries/v1"},
    {"method": "GET",  "url": "https://S-api.Fnet.F/services/externalservice/api/Lookups/CountriesDetails/v1"},
    {"method": "GET",  "url": "https://S-api.Fnet.F/services/externalservice/api/Lookups/Divisions/v1"},
    {"method": "POST", "url": "https://S-api.Fnet.F/services/externalservice/api/CaseClassifications/Divisions/v1"},
]

S_LINK_TEMPLATE: str = "https://S.Fnet.F/apps/desktop/#/main/serial/{unique_id}"
S_BATCH_SIZE: int = 500
S_RATE_LIMIT_SLEEP: int = 10      # seconds between batches
S_TOKEN_MIN_LEN: int = 20

# ---------------------------------------------------------------------------
# Selector types
# Order matters — type detection evaluates in this sequence.
# "row" is an internal marker for multi-selector schema rows; never stored in DB.
# ---------------------------------------------------------------------------
SELECTOR_TYPES: tuple[str, ...] = (
    "email",
    "phone",
    "ip",
    "address",
    "linkedin",
    "github",
    "telegram",
    "discord",
    "name",
    "other",
)

# ---------------------------------------------------------------------------
# Delimiters
# ---------------------------------------------------------------------------
INTERNAL_DELIM: str = "!!"   # joins selectors within a single row
ROW_DELIM: str = "+++"       # separates rows

# ---------------------------------------------------------------------------
# ID format
# Produces YYMMDDHHMMSS; caller appends zero-padded 3-digit milliseconds.
# ---------------------------------------------------------------------------
ID_STRFTIME: str = "%y%m%d%H%M%S"

# ---------------------------------------------------------------------------
# App metadata
# ---------------------------------------------------------------------------
APP_NAME: str = "GrayWolfe"
APP_VERSION: str = "2.0.0"
