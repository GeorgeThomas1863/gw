# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

GrayWolfe is a Python/Tkinter desktop application for managing selectors (names, emails, phones, IPs, addresses, etc.). It searches across two data sources — the local GrayWolfe SQLite database and the "S" external API — and provides data import with automatic type detection and target management.

**Core problem:** Users collect thousands of selectors from disparate sources. GrayWolfe detects selector types, deduplicates them, and links related selectors into "targets" — connected groups representing a single entity. "Targets" means selector groups, not military targets.

**Storage:** Each user has a personal SQLite DB at `\\server\graywolfe\users\{username}.db`. A master DB lives at `\\server\graywolfe\master.db`. An admin tool merges user DBs into master. No SharePoint.

**`old_vba/`** contains the archived VBA source from before the Python rewrite. It is reference-only — do not modify it.

## Running

```bash
python app.py
```

Tests (gitignored locally):
```bash
pytest                          # all tests
pytest tests/test_type_engine.py   # single file
pytest -k test_detect_email     # single test
```

Admin — merge a user's DB into master:
```bash
python admin/merge_user_db.py path/to/user.db [--dry-run]
```

## Architecture

```
app.py           Entry point — resolves DB path, initializes schema, launches GrayWolfeApp
config.py        All constants: paths, S API config, selector types, delimiters, ID format
src/             Business logic
  search.py        Search orchestration (GW + S)
  import_data.py   Import workflows (default and unrelated)
  targets.py       Target lifecycle: create, merge, collect
  type_engine.py   Selector type detection, validation, cleaning
data/            Data access
  database.py      SQLite connection, DDL, CRUD helpers, ID generation
  s_api.py         S API client (search, auth, link generation)
  sync.py          pull_from_master() — merges master DB into local DB
util/
  errors.py        GWError exception + error codes (1950–1972)
  logger.py        get_logger() wrapper
display/         Tkinter windows
  main.py          GrayWolfeApp (root window, Search + Add Data tabs)
  results_window.py  Results display (GW and S tabs, dynamic filters)
  schema_detection.py  Column type confirmation during default import
  target_details.py    Target editor
  merge_modal.py       Merge two targets
  strings.py           UI string constants
admin/
  merge_user_db.py  CLI: merges user DB into master, flags conflicts to log
tests/           pytest test suite (gitignored — not committed)
old_vba/         Archived VBA source (reference only)
```

### Key Workflows

**Search:** `app.py` → `src/search.py:run_search()` → queries local SQLite + `data/s_api.py:SApiClient` → opens `display/results_window.py`

**Default Import:** `app.py` → `src/import_data.py:detect_column_types()` → opens `display/schema_detection.py` for user confirmation → `run_default_import()` → creates selectors + targets, handles bridging merges

**Unrelated Import:** `run_unrelated_import()` — direct type detection, skips relationship/target creation

**Target Edit:** `display/target_details.py` → `src/targets.py` → local SQLite write

**Merge Targets:** `display/merge_modal.py` → `src/targets.py:merge_targets()` — absorbs one target into another; reassigns all selectors

## Database Schema

Three tables (defined in `data/database.py`):

| Table | Purpose |
|-------|---------|
| `selectors` | Individual selector values with type, target_id, nork_id |
| `targets` | Selector groups with metadata (name, case_number, laptop_count) |
| `norks` | Entity metadata |

`selector_clean` is the normalized form used for deduplication. During import, any matching `selector_clean` blocks the insert (global uniqueness). During manual add via Target Details, `selector_clean + target_id` is checked instead.

## Selector Types

10 types in detection priority order (from `config.SELECTOR_TYPES`):
`email`, `phone`, `ip`, `address`, `linkedin`, `github`, `telegram`, `discord`, `name`, `other`

`src/type_engine.py` contains `check_<type>()` (bool), `fix_<type>()` (cleaned string), `detect_selector_type()`, and `build_selector_clean()`.

## Error Handling

`util/errors.py` defines `GWError(code, message)` and `raise_gw(code, msg)`. Codes 1950–1972 mirror the old VBA system. Raise `GWError` for all application-level errors; let SQLite/requests exceptions propagate or wrap them with `raise_gw()`.

## Key Technical Details

- **ID generation:** `data/database.py:generate_id()` — `YYMMDDHHMMSS` + 3-digit ms, thread-safe, collision-proof
- **Delimiters:** `INTERNAL_DELIM = "!!"` (columns within a row), `ROW_DELIM = "+++"` (row separator)
- **S API rate limiting:** 500-item batches with 10-second pauses (`config.S_BATCH_SIZE`, `S_RATE_LIMIT_SLEEP`)
- **Threading:** Search and import run in background threads; results returned to main thread via `queue.Queue` polled in `display/main.py:_poll_queue()`
- **DB connection:** WAL mode, `foreign_keys=ON`, `row_factory=sqlite3.Row`

## Target Bridging

When an imported row's selectors match 2+ existing targets, `collect_target_ids_for_row()` gathers all matching `target_id`s, then `merge_targets()` combines them. This preserves the invariant that selectors on the same row belong to one entity.

## Known Architectural Gaps

- **Search doesn't expand to target group:** `run_search()` returns exact selector matches only — it does not fetch sibling selectors from the same target.
- **No import transaction safety:** Import operations are not wrapped in transactions. A mid-import failure can leave partial data in local DB.
