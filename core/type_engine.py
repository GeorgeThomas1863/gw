"""
core/type_engine.py — Selector type detection, validation, and cleaning.

Detection priority order matches config.SELECTOR_TYPES:
  email, phone, ip, address, linkedin, github, telegram, discord, name, other

Public API
----------
check_<type>(value)           -> bool
fix_<type>(value)             -> str   (cleaned form, "" if invalid)
detect_selector_type(value)   -> str   (type name)
build_selector_clean(value, sel_type) -> str
normalize_state(value)        -> str   (2-letter abbrev or "FAIL")
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import SELECTOR_TYPES

# ---------------------------------------------------------------------------
# Compiled regex patterns
# ---------------------------------------------------------------------------

_EMAIL_RE = re.compile(
    r"^[a-zA-Z0-9]([a-zA-Z0-9._+-]*[a-zA-Z0-9])?@"
    r"[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?"
    r"(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$",
    re.IGNORECASE,
)

_PHONE_US_RE = re.compile(
    r"^[\s\(]*(\+?1[\s\-\.]?)?[\s\(]*([2-9]\d{2})[\s\)\-\.]*([2-9]\d{2})[\s\-\.]*(\d{4})[\s\)]*$"
)

_PHONE_INTL_RE = re.compile(r"^\+(\d{1,3})[\s\-\.]?(\d[\d\s\-\.]{6,14})$")

_IPV4_RE = re.compile(
    r"^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}"
    r"(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
)

_IPV6_RE = re.compile(
    r"^(([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}|"
    r"([0-9a-fA-F]{1,4}:){1,7}:|"
    r"([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|"
    r"([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|"
    r"([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|"
    r"([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|"
    r"([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|"
    r"[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|"
    r":((:[0-9a-fA-F]{1,4}){1,7}|:)|::)$",
    re.IGNORECASE,
)

_NAME_RE = re.compile(
    r"^[^\d#_/\\?\[\]@]{1,50}(\s[^\d#_/\\?\[\]@]{1,50}){0,3}$"
)

_TELEGRAM_USERNAME_RE = re.compile(r"^[a-z0-9_]{5,32}$", re.IGNORECASE)

# ---------------------------------------------------------------------------
# State data
# ---------------------------------------------------------------------------

_STATE_PAIRS: list[tuple[str, str]] = [
    # US states and territories
    ("AL", "Alabama"),
    ("AK", "Alaska"),
    ("AZ", "Arizona"),
    ("AR", "Arkansas"),
    ("AS", "American Samoa"),
    ("CA", "California"),
    ("CO", "Colorado"),
    ("CT", "Connecticut"),
    ("DE", "Delaware"),
    ("DC", "District of Columbia"),
    ("FL", "Florida"),
    ("GA", "Georgia"),
    ("GU", "Guam"),
    ("HI", "Hawaii"),
    ("ID", "Idaho"),
    ("IL", "Illinois"),
    ("IN", "Indiana"),
    ("IA", "Iowa"),
    ("KS", "Kansas"),
    ("KY", "Kentucky"),
    ("LA", "Louisiana"),
    ("ME", "Maine"),
    ("MD", "Maryland"),
    ("MA", "Massachusetts"),
    ("MI", "Michigan"),
    ("MN", "Minnesota"),
    ("MS", "Mississippi"),
    ("MO", "Missouri"),
    ("MT", "Montana"),
    ("NE", "Nebraska"),
    ("NV", "Nevada"),
    ("NH", "New Hampshire"),
    ("NJ", "New Jersey"),
    ("NM", "New Mexico"),
    ("NY", "New York"),
    ("NC", "North Carolina"),
    ("ND", "North Dakota"),
    ("MP", "Northern Marianas"),
    ("OH", "Ohio"),
    ("OK", "Oklahoma"),
    ("OR", "Oregon"),
    ("PA", "Pennsylvania"),
    ("PR", "Puerto Rico"),
    ("RI", "Rhode Island"),
    ("SC", "South Carolina"),
    ("SD", "South Dakota"),
    ("TN", "Tennessee"),
    ("TX", "Texas"),
    ("UT", "Utah"),
    ("VT", "Vermont"),
    ("VA", "Virginia"),
    ("VI", "Virgin Islands"),
    ("WA", "Washington"),
    ("WV", "West Virginia"),
    ("WI", "Wisconsin"),
    ("WY", "Wyoming"),
    # Canadian provinces and territories
    ("AB", "Alberta"),
    ("BC", "British Columbia"),
    ("MB", "Manitoba"),
    ("NB", "New Brunswick"),
    ("NL", "Newfoundland"),
    ("NT", "Northwest Territories"),
    ("NS", "Nova Scotia"),
    ("NU", "Nunavut"),
    ("ON", "Ontario"),
    ("PE", "Prince Edward Island"),
    ("QC", "Quebec"),
    ("SK", "Saskatchewan"),
    ("YT", "Yukon"),
]

# Maps lowercased full name -> abbreviation
_STATE_ABBREV_MAP: dict[str, str] = {
    full.lower(): abbrev for abbrev, full in _STATE_PAIRS
}
# Maps lowercased abbreviation -> abbreviation (canonical form)
_STATE_ABBREV_UPPER_MAP: dict[str, str] = {
    abbrev.lower(): abbrev for abbrev, _ in _STATE_PAIRS
}
# Maps abbreviation -> full name
_STATE_NAME_MAP: dict[str, str] = {abbrev: full for abbrev, full in _STATE_PAIRS}

# Frozenset of all recognized values (lowercased for lookup)
_STATES: frozenset[str] = frozenset(
    list(_STATE_ABBREV_MAP.keys()) + list(_STATE_ABBREV_UPPER_MAP.keys())
)

# Build the state alternation used in _ADDRESS_RE
_STATE_PATTERN_PART = "|".join(
    re.escape(abbrev) for abbrev, _ in _STATE_PAIRS
)

_ADDRESS_RE = re.compile(
    r"^\d+\s+.+,\s*.+,\s*(" + _STATE_PATTERN_PART + r")\s+\d{4,10}\s*$",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Street terms
# ---------------------------------------------------------------------------

_STREETS: frozenset[str] = frozenset({
    # Common abbreviations
    "st", "ave", "blvd", "rd", "dr", "ln", "pl", "ct", "cir",
    "hwy", "pkwy", "rt", "trl", "ter", "expy",
    # Full words
    "street", "avenue", "boulevard", "road", "drive", "lane",
    "place", "court", "circle", "highway", "parkway", "route",
    "trail", "terrace", "way", "loop", "run", "pass", "crossing",
    "expressway", "freeway", "alley", "path", "row", "ridge",
    "point", "park", "grove", "meadow", "creek", "hollow",
    "branch", "bend", "crest", "hill", "heights", "manor",
    "estates", "commons", "landing", "overlook", "trace",
    "walk", "square", "plaza", "close", "mews",
    # Additional from VBA DefineStreetArr
    "pike", "turnpike", "bypass", "spur", "cutoff", "extension",
    "connector", "frontage",
})

# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def check_email(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    return bool(_EMAIL_RE.match(value))


def fix_email(value: str) -> str:
    """Return cleaned email. Applies Gmail normalization (dot removal, plus strip).
    Returns '' if invalid."""
    value = value.strip().lower()
    if not _EMAIL_RE.match(value):
        return ""
    local, domain = value.rsplit("@", 1)
    if domain in ("gmail.com", "googlemail.com"):
        # Strip plus addressing
        local = local.split("+")[0]
        # Remove dots from local part
        local = local.replace(".", "")
    return f"{local}@{domain}"


# ---------------------------------------------------------------------------
# Phone
# ---------------------------------------------------------------------------

def check_phone(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    return bool(_PHONE_US_RE.match(value)) or bool(_PHONE_INTL_RE.match(value))


def fix_phone(value: str) -> str:
    """Normalize US phones to NXX-NXX-XXXX, international to +CC NNNNNNNN.
    Returns '' if invalid."""
    value = value.strip()
    if not value:
        return ""

    m = _PHONE_US_RE.match(value)
    if m:
        area = m.group(2)
        exchange = m.group(3)
        number = m.group(4)
        return f"{area}-{exchange}-{number}"

    m = _PHONE_INTL_RE.match(value)
    if m:
        country_code = m.group(1)
        # Strip all whitespace/dashes from the subscriber number
        subscriber = re.sub(r"[\s\-\.]", "", m.group(2))
        return f"+{country_code} {subscriber}"

    return ""


# ---------------------------------------------------------------------------
# IP
# ---------------------------------------------------------------------------

def check_ip(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    return bool(_IPV4_RE.match(value)) or bool(_IPV6_RE.match(value))


def fix_ip(value: str) -> str:
    """Return value as-is if valid IP, '' otherwise."""
    value = value.strip()
    if not check_ip(value):
        return ""
    return value


# ---------------------------------------------------------------------------
# Address
# ---------------------------------------------------------------------------

def check_address(value: str) -> bool:
    value = value.strip()
    if not value:
        return False

    # Must contain both numeric and alphabetic characters
    if not re.search(r"\d", value) or not re.search(r"[a-zA-Z]", value):
        return False

    # Must be 1-15 words
    word_count = len(value.split())
    if word_count < 1 or word_count > 15:
        return False

    # Must contain a recognized street term
    words_lower = [w.strip(".,").lower() for w in value.split()]
    if not any(w in _STREETS for w in words_lower):
        return False

    # Must match state+zip pattern
    if not _ADDRESS_RE.match(value):
        return False

    return True


def fix_address(value: str) -> str:
    """Normalize address to 'street city, ST zip'. Returns '' if invalid."""
    value = value.strip()
    if not check_address(value):
        return ""
    # Normalize internal whitespace; state normalization via normalize_state
    # Split at commas: [street], [city], [state zip]
    parts = [p.strip() for p in value.split(",")]
    if len(parts) >= 3:
        street = parts[0]
        city = parts[1]
        state_zip = parts[2].strip()
        state_zip_parts = state_zip.split()
        if len(state_zip_parts) >= 2:
            state = normalize_state(state_zip_parts[0])
            zip_code = state_zip_parts[1]
            if state != "FAIL":
                return f"{street}, {city}, {state} {zip_code}"
    return value


# ---------------------------------------------------------------------------
# LinkedIn
# ---------------------------------------------------------------------------

_LINKEDIN_DOMAIN_RE = re.compile(
    r"^(https?://)?linkedin\.com(/|$)", re.IGNORECASE
)


def check_linkedin(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    if not _LINKEDIN_DOMAIN_RE.match(value):
        return False
    if len(value) < 5 or len(value) > 200:
        return False
    if re.search(r"[ !,?\\+]", value):
        return False
    return True


def fix_linkedin(value: str) -> str:
    """Normalize to 'linkedin.com/in/handle'. Returns '' if invalid."""
    value = value.strip()
    if not check_linkedin(value):
        return ""
    # Strip scheme
    value = re.sub(r"^https?://", "", value, flags=re.IGNORECASE)
    return value


# ---------------------------------------------------------------------------
# GitHub
# ---------------------------------------------------------------------------

_GITHUB_DOMAIN_RE = re.compile(
    r"^(https?://)?github\.com(/|$)", re.IGNORECASE
)


def check_github(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    if not _GITHUB_DOMAIN_RE.match(value):
        return False
    if len(value) < 6 or len(value) > 200:
        return False
    if re.search(r"\s", value):
        return False
    return True


def fix_github(value: str) -> str:
    """Normalize to 'github.com/handle'. Returns '' if invalid."""
    value = value.strip()
    if not check_github(value):
        return ""
    value = re.sub(r"^https?://", "", value, flags=re.IGNORECASE)
    return value


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------

def check_telegram(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    # Only @handle form: starts with @, body is 5-32 alphanumeric/underscore.
    # Bare numeric IDs are indistinguishable from zip codes / case numbers and
    # are intentionally not supported.
    if value.startswith("@"):
        username = value[1:]
        return bool(_TELEGRAM_USERNAME_RE.match(username))
    return False


def fix_telegram(value: str) -> str:
    """Return '@handle' or numeric ID. Returns '' if invalid."""
    value = value.strip()
    if not check_telegram(value):
        return ""
    return value


# ---------------------------------------------------------------------------
# Discord
# ---------------------------------------------------------------------------

_DISCORD_FORBIDDEN_RE = re.compile(r"[ .,!%$/\\+\"]")


def check_discord(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    # Must start with @ OR contain # somewhere that isn't position 0
    starts_with_at = value.startswith("@")
    hash_pos = value.find("#")
    has_hash_not_start = hash_pos > 0

    if not starts_with_at and not has_hash_not_start:
        return False

    if len(value) < 5 or len(value) > 40:
        return False

    if _DISCORD_FORBIDDEN_RE.search(value):
        return False

    return True


def fix_discord(value: str) -> str:
    """Strip forbidden chars, preserve @/# prefix. Returns '' if invalid."""
    value = value.strip()
    if not check_discord(value):
        return ""
    # Remove forbidden characters (except the @ prefix and # separator we keep)
    cleaned = re.sub(r"[.,!%$/\\+\"]", "", value)
    return cleaned


# ---------------------------------------------------------------------------
# Name
# ---------------------------------------------------------------------------

def check_name(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    # Lowercase before applying the regex (consistent with VBA behavior)
    return bool(_NAME_RE.match(value.lower()))


def fix_name(value: str) -> str:
    """Lowercase and strip apostrophes. Returns '' if invalid."""
    value = value.strip()
    if not check_name(value):
        return ""
    cleaned = value.replace("'", "").lower()
    return cleaned


# ---------------------------------------------------------------------------
# Other
# ---------------------------------------------------------------------------

def fix_other(value: str) -> str:
    """Strip +!@# chars and lowercase. Returns '' if empty."""
    value = value.strip()
    if not value:
        return ""
    cleaned = re.sub(r"[+!@#]", "", value)
    return cleaned.lower()


# ---------------------------------------------------------------------------
# normalize_state
# ---------------------------------------------------------------------------

def normalize_state(value: str) -> str:
    """Return 2-letter abbreviation for a state name or abbreviation.
    Returns 'FAIL' if not recognized (mirrors VBA StateMap behavior)."""
    value = value.strip()
    if not value:
        return "FAIL"
    lower = value.lower()
    # Try full-name lookup first
    if lower in _STATE_ABBREV_MAP:
        return _STATE_ABBREV_MAP[lower]
    # Try abbreviation lookup
    if lower in _STATE_ABBREV_UPPER_MAP:
        return _STATE_ABBREV_UPPER_MAP[lower]
    return "FAIL"


# ---------------------------------------------------------------------------
# detect_selector_type
# ---------------------------------------------------------------------------

_CHECK_MAP: dict[str, object] = {
    "email":    check_email,
    "phone":    check_phone,
    "ip":       check_ip,
    "address":  check_address,
    "linkedin": check_linkedin,
    "github":   check_github,
    "telegram": check_telegram,
    "discord":  check_discord,
    "name":     check_name,
}


def detect_selector_type(value: str) -> str:
    """Try each type in SELECTOR_TYPES priority order.
    Returns matching type name, or 'other' if nothing matches."""
    for sel_type in SELECTOR_TYPES:
        if sel_type == "other":
            return "other"
        checker = _CHECK_MAP.get(sel_type)
        if checker and checker(value):  # type: ignore[operator]
            return sel_type
    return "other"


# ---------------------------------------------------------------------------
# build_selector_clean
# ---------------------------------------------------------------------------

_FIX_MAP: dict[str, object] = {
    "email":    fix_email,
    "phone":    fix_phone,
    "ip":       fix_ip,
    "address":  fix_address,
    "linkedin": fix_linkedin,
    "github":   fix_github,
    "telegram": fix_telegram,
    "discord":  fix_discord,
    "name":     fix_name,
    "other":    fix_other,
}


def build_selector_clean(value: str, sel_type: str) -> str:
    """Call the appropriate fix_{type} function for sel_type.
    Returns cleaned value, or '' if sel_type is unrecognized."""
    fixer = _FIX_MAP.get(sel_type)
    if fixer is None:
        return ""
    return fixer(value)  # type: ignore[operator]
