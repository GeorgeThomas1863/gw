# GrayWolfe — Logic Reference: Search and Add Data Flows

Full call-chain documentation for what happens when the user hits Submit in either tab.
Function signatures are simplified; see the actual source for full type annotations.

---

## Table of Contents

- [Search Flow](#search-flow)
- [Add Data Flow — Default Import](#add-data-flow--default-import)
- [Add Data Flow — Unrelated Import](#add-data-flow--unrelated-import)
- [Shared Subroutines](#shared-subroutines)

---

## Search Flow

### Entry point

**User action:** clicks "Submit" in the Search tab.

**Bound to:** `display/main.py › GrayWolfeApp._do_search()`

---

### Step 1 — Input validation and setup (`display/main.py › _do_search`)

1. Reads raw text from the search text widget via `_get_text_widget_value(widget)`.
   Returns empty string if the placeholder is still active. Warns and aborts if empty.

2. Reads the search mode combobox (`"GW + S"` / `"GW Only"` / `"S Only"`) and derives two booleans:
   `search_gw` and `search_s`.

3. If `search_s` is True:
   - Reads the S token entry field.
   - Constructs `data/s_api.py › SApiClient(token)` — validates token format only (length, no spaces).
     Raises `GWError` if malformed; shows error dialog and aborts.

4. Resolves the delimiter: looks up `_DELIMITER_OPTIONS[self._search_delim_var.get()]`
   which maps the combobox label to a character or `None` (auto).

5. Calls `src/search.py › parse_raw_input(raw, delim)` to build `query_terms: list[str]`.
   These are used only for the results window display count — the actual search re-parses internally.

6. Sets up two callbacks:
   - `progress_cb(selector, idx, total)` — puts a `("progress", message, None)` tuple on `_result_queue`
     so the status bar updates while S is paginating.
   - `ask_cb(selector, num_found) -> bool` — puts `(selector, num_found)` on `_ask_queue` and
     blocks for up to 30 seconds waiting for a boolean answer on `_answer_queue`.
     The main thread's `_poll_queue()` loop drains `_ask_queue` and shows a `messagebox.askyesno`.

7. Sets status to `"Searching…"` and calls `_run_in_thread(self._search_worker, ...)`.

---

### Step 2 — Threading (`display/main.py › _run_in_thread`)

1. Calls `_set_busy(True)` — disables all action buttons.
2. Spawns a daemon `threading.Thread` that runs:
   ```
   result = func(*args)
   _result_queue.put(("ok", result, on_complete))
   ```
   On exception:
   ```
   _result_queue.put(("err", exc, on_error))
   ```
3. Returns immediately. The main thread continues its `after()`-based polling loop.

---

### Step 3 — Search worker (background thread) (`display/main.py › _search_worker`)

Calls:
```
src/search.py › run_search(raw, delim, conn, s_client, search_gw, search_s_flag, progress_cb, ask_cb)
```
Returns `(gw_results, s_results, query_terms, s_client)` to the queue.

---

### Step 4 — `run_search` (`src/search.py`)

1. Calls `parse_raw_input(raw, delim)` → `list[str]`.
   - If delimiter is specified: `raw.split(delimiter)`.
   - Auto-detect order: Tab → Newline → Comma.
   - Each value is stripped; empty strings are discarded.
   - Raises `GWError(ERR_EMPTY_SEARCH)` if result is empty.

2. If `search_gw`: calls `search_graywolfe(selectors, conn)` → `list[GWResult]`.

3. If `search_s_flag and s_client is not None`: calls `search_s(selectors, s_client, progress_cb, ask_cb)` → `list[SApiResult]`.

4. Returns `(gw_results, s_results)`.

---

### Step 5a — GrayWolfe local search (`src/search.py › search_graywolfe`)

For each selector value:

1. `src/type_engine.py › detect_selector_type(value)` → `SelectorType`
   Runs check functions in priority order: `check_email`, `check_phone`, `check_ip`,
   `check_address`, `check_linkedin`, `check_github`, `check_telegram`, `check_discord`,
   `check_name`. Returns first match, or `SelectorType.OTHER`.

2. `src/type_engine.py › build_selector_clean(value, sel_type)` → `str | None`
   Runs the corresponding `fix_<type>(value)` function to produce the normalized form
   used for deduplication (e.g., strips formatting from phone numbers, lowercases emails).
   Returns `None` if normalization fails or produces an empty string.

3. If `clean` is non-empty: executes:
   ```sql
   SELECT s.*, t.target_name
   FROM selectors s
   LEFT JOIN targets t ON s.target_id = t.target_id
   WHERE s.selector_clean = ?
   ```

4. If rows found: builds one `GWResult(in_gray_wolfe=True, ...)` per row.
   Safely coerces `selector_type` string from DB back to `SelectorType` enum (falls back to `None` on unknown value).

5. If no rows found: builds one `GWResult(in_gray_wolfe=False, selector=value, everything_else=None)`.

Returns `list[GWResult]`.

---

### Step 5b — S API search (`src/search.py › search_s`)

1. Calls `data/s_api.py › SApiClient.validate_token()`:
   - POSTs/GETs to each URL in `config.S_VALIDATE_URLS`.
   - Returns `True` if any returns 2xx. Raises `GWError(ERR_S_AUTH_FAILED)` if all fail.

2. For each selector value, calls `data/s_api.py › SApiClient.search(value, on_rate_limit, ask_continue)`:

   a. Calls `_search_batch(query, start=0)`:
      - POSTs to `config.S_SEARCH_URL` with `{"q": '"value"', "limit": 500, "start": 0}`.
      - Raises `GWError(ERR_S_REQUEST_FAILED)` on network error.
      - Raises `GWError(ERR_S_AUTH_FAILED)` on HTTP 401/403.
      - Returns parsed JSON dict.

   b. Reads `numFound` from response. Calls `_parse_item(item, query)` → `SApiResult` for each item.

   c. If `numFound > 500` and `ask_continue` is provided: calls `ask_continue(query, numFound)`.
      If user says no, returns first batch only.

   d. For each subsequent batch (start += 500):
      - Calls `on_rate_limit()` (triggers a status bar update).
      - Sleeps `config.S_RATE_LIMIT_SLEEP` seconds.
      - Fetches next batch, parses items.

   `_parse_item(item, query)` maps API JSON fields to `SApiResult` fields:
   `uniqueID` → `s_id`, `recordType` → `doc_type`, `UCFN` → `case`, `itemNumber` → `serial`, etc.
   `link` is constructed via `SApiClient.get_link(unique_id)` using `config.S_LINK_TEMPLATE`.

3. On per-selector error: logs warning and continues to next selector.

Returns `list[SApiResult]`.

---

### Step 6 — Queue drain and result display (main thread)

`display/main.py › _poll_queue()` runs every ~100 ms via `self.after()`.

On `("ok", result, on_complete)`:
1. Calls `_set_busy(False)` — re-enables buttons.
2. Calls `_on_search_complete(result)`:
   - Sets status to `"Ready"`.
   - Opens `display/results_window.py › ResultsWindow(parent, gw_results, s_results, query_terms, conn, s_client)`.

`ResultsWindow.__init__` calls:
- `_build_gw_tab()` / `_build_s_tab()` — builds the notebook UI.
- `_populate_gw()`:
  - Builds s_hit_count lookup from `s_results`.
  - For each `GWResult`: assembles a display dict (selector, type, target, in_gw, s_hits).
  - Populates target filter combobox.
  - Calls `_apply_gw_filter()` to render rows.
- `_populate_s()`:
  - Stores all `SApiResult` rows in `self._s_all_rows`.
  - Populates selector filter combobox.
  - Calls `_apply_s_filter()` to render rows.

Double-click on a GW row opens `display/target_details.py › TargetDetailsWindow`.
Double-click on an S row opens the document link in the browser via `webbrowser.open(row.link)`.

---

## Add Data Flow — Default Import

### Entry point

**User action:** clicks "Submit" in the Add Selectors tab with mode set to "Default Import".

**Bound to:** `display/main.py › GrayWolfeApp._do_add()`

---

### Step 1 — Input validation and setup (`display/main.py › _do_add`)

1. Reads raw text. Warns and aborts if empty.
2. Reads import mode, delimiter, and (optionally) S token.
3. For Default Import:
   - Calls `src/import_data.py › parse_import_input(raw, delim)` → `list[list[str]]`.
   - Calls `src/import_data.py › detect_column_types(rows)` → `list[ColumnTypeInfo]`.
   - Opens `display/schema_detection.py › SchemaDetectionDialog(parent, rows, detected, on_confirm)`.
   - The dialog is modal; execution continues when the user hits Submit or Cancel in it.

---

### Step 2 — Parse input (`src/import_data.py › parse_import_input`)

- Splits on `"\n"` to get rows.
- Within each row, splits on tab or explicit delimiter for columns.
- Strips and filters empty values.
- Returns `list[list[str]]` — each inner list is one row's column values.

---

### Step 3 — Detect column types (`src/import_data.py › detect_column_types`)

For each column index across all rows:

1. Collects non-empty values for that column position.
2. For each value: calls `src/type_engine.py › detect_selector_type(value)` → `SelectorType`.
3. Tallies counts per type; picks the type with the most matches.
4. `confidence = count_of_best_type / total_values_in_column`.
5. Empty columns get `selector_type=SelectorType.OTHER, confidence=0.0`.

Returns `list[ColumnTypeInfo(col_index, selector_type, confidence)]`.

---

### Step 4 — Schema Detection Dialog (`display/schema_detection.py › SchemaDetectionDialog`)

1. Shows a preview table of the first 5 rows.
2. Shows one combobox per column (capped at 6) pre-filled with the detected type.
   Available choices: all 10 selector types + `"null"` (skip column).

3. On Submit:
   - **typeNull check**: for any column with `confidence == 0.0` that the user left non-null —
     prompts: "GW couldn't detect a type for [sample value]. Mark as null?" If user says No, dialog stays open.
   - **typeWrong check**: for any column where the user changed a detection with `confidence >= 0.5` —
     prompts: "GW detected [type] for [sample value]. Override to [user_choice]?" If user says No, reverts the combobox.

4. Calls `on_confirm(confirmed_types: list[str])` → `display/main.py › _run_default_import(rows, types, raw, delim, s_client)`.

---

### Step 5 — Background import (`display/main.py › _run_default_import`)

Sets status to `"Importing…"`, calls `_run_in_thread(run_default_import, rows, confirmed_types, conn, username)`.

---

### Step 6 — Default import worker (`src/import_data.py › run_default_import`)

Runs in background thread. For each row:

#### 6a — Insert selectors

For each `(col_idx, value)` in the row where `confirmed_types[col_idx] != "null"`:

1. If type is `"auto"`: calls `detect_selector_type(value)` → `SelectorType`.
   Otherwise uses the confirmed type directly.

2. Calls `_make_selector(value, effective_type, username)` → `Selector | None`:
   - `src/type_engine.py › build_selector_clean(value, sel_type)` → normalized string.
   - Returns `None` if clean string is empty (value is unparseable for that type).
   - Otherwise constructs `Selector(selector_id=generate_id(), selector_clean=clean, ...)`.
   - `data/database.py › generate_id()` → 15-char timestamp-based ID (`YYMMDDHHMMSS` + 3-digit ms), thread-safe.
   - `data/database.py › now_iso()` → UTC ISO 8601 timestamp for `date_created` / `last_updated`.

3. Calls `data/database.py › insert_selector(conn, sel)`:
   - Checks `SELECT selector_id FROM selectors WHERE selector_clean = ?`.
   - If found: raises `GWError(ERR_SELECTOR_DUPLICATE)`.
   - If not found: `INSERT INTO selectors (...)` using `dataclasses.asdict(selector)`, then `conn.commit()`.

4. On `ERR_SELECTOR_DUPLICATE`: logs warning, increments `skipped`. Continues to next column.

#### 6b — Target resolution

Skipped entirely if the row has fewer than 2 columns.

Calls `src/targets.py › collect_target_ids_for_row(row_values, conn)`:
- For each value in `row_values` (all values submitted for that row, including duplicates):
  - `detect_selector_type(value)` + `build_selector_clean(value, sel_type)`.
  - Queries: `SELECT DISTINCT target_id FROM selectors WHERE selector_clean = ? AND target_id IS NOT NULL`.
- Returns deduplicated list of `target_id` strings in insertion-stable order.

**Case 0 — No existing targets:**
- Calls `src/targets.py › create_target(conn, username)`:
  - Constructs `Target(target_id=generate_id(), target_name="[New Target]", ...)`.
  - Calls `data/database.py › insert_target(conn, target)` → `INSERT INTO targets (...)`.
  - Returns `target_id`.
- Links all newly inserted selectors to the new target.

**Case 1 — Exactly one existing target:**
- Links all newly inserted selectors to that target.

**Case 2+ — Multiple existing targets (bridging merge):**
- Keeps the first target as survivor. For each additional target:
  - Calls `src/targets.py › merge_targets(keep_id, absorb_id, conn, username)`:
    1. Validates both IDs are non-empty and distinct.
    2. Fetches both `Target` objects; raises `GWError(ERR_TARGET_NOT_FOUND)` if either missing.
    3. `UPDATE selectors SET target_id = keep WHERE target_id = absorb`.
    4. `DELETE FROM selectors WHERE target_id = keep AND id NOT IN (SELECT MIN(id) ... GROUP BY selector_clean)` — removes duplicates, keeps oldest row per `selector_clean`.
    5. `conn.commit()`.
    6. Copies `target_name` from absorb → keep if keep's is blank.
    7. Copies `case_number` from absorb → keep if keep's is blank.
    8. Adds absorb's `laptop_count` to keep's (if absorb > 0).
    9. Updates keep's `last_updated` / `last_updated_by`.
    10. `data/database.py › delete_target(conn, absorb_id)`.
- Links all newly inserted selectors to survivor.

**Linking** (`_link_selectors_to_target(conn, selector_cleans, target_id)`):
```sql
UPDATE selectors SET target_id = ?
WHERE selector_clean IN (...)
AND target_id IS NULL
```
Processed in chunks of 500 to stay under SQLite's 999-variable limit.
Only assigns selectors that have no existing target (preserves pre-existing links).

Returns `(total_inserted, total_skipped)`.

---

### Step 7 — Import complete (main thread) (`display/main.py › _on_import_complete`)

1. Sets status to `"Ready"`.
2. Shows `messagebox.showinfo("Import Complete", "N submitted, N inserted, N skipped")`.
3. If `s_client is not None` (user checked "Also Search S?"):
   - Calls `_do_s_search_after_import(raw, delim, s_client)`.
   - Parses `query_terms = parse_raw_input(raw, delim)`.
   - Calls `_run_in_thread(_s_search_worker)` which calls:
     `src/search.py › run_search(raw, delim, conn, s_client, search_gw=False, search_s_flag=True, ...)`.
   - On complete: opens `display/results_window.py › ResultsWindow`.

---

## Add Data Flow — Unrelated Import

### Entry point

Same button/handler as Default Import: `display/main.py › _do_add()`.

Difference: no schema detection step, no target linking. Each selector is inserted as a standalone row.

---

### Step 1 — Setup (`display/main.py › _do_add`)

1. Reads `sel_type_display` from the Selector Type combobox.
   Maps `"Auto-Detect"` → `"auto"`, otherwise uses the type string directly (e.g. `"email"`).
2. Sets status to `"Importing…"`.
3. Calls `_run_in_thread(run_unrelated_import, raw, sel_type, conn, username, delim, ...)`.

---

### Step 2 — Unrelated import worker (`src/import_data.py › run_unrelated_import`)

Runs in background thread.

1. Calls `src/search.py › parse_raw_input(raw_input, delimiter)` → `list[str]`.
   This is a flat parse (no columns) — each line or tab-separated token is one value.

2. For each value:
   - If `sel_type == "auto"`: calls `detect_selector_type(value)` → `SelectorType`.
     Otherwise uses the given type.
   - Calls `_make_selector(value, effective_type, username)` → `Selector | None` (same as Default Import).
   - Calls `data/database.py › insert_selector(conn, sel)` (same duplicate check).
   - On duplicate: logs warning, increments `skipped`. Continues.

3. **No target resolution.** Selectors are inserted with `target_id = NULL`.

Returns `(inserted, skipped)`.

---

### Step 3 — Import complete

Same as Default Import Step 7: status reset, summary messagebox, optional S search.

---

## Shared Subroutines

### `src/type_engine.py › detect_selector_type(value: str) -> SelectorType`

Runs check functions in this priority order:
`check_email` → `check_phone` → `check_ip` → `check_address` → `check_linkedin` →
`check_github` → `check_telegram` → `check_discord` → `check_name`.
Returns the first matching `SelectorType`, or `SelectorType.OTHER` if none match.

### `src/type_engine.py › build_selector_clean(value: str, sel_type: SelectorType | str) -> str`

Calls the corresponding `fix_<type>(value)` function to normalize the value.
Examples: `fix_email` lowercases + strips; `fix_phone` strips all non-digits; `fix_ip` normalizes octets.
Returns empty string if normalization fails or the result is empty.

### `data/database.py › generate_id() -> str`

Returns a 15-character unique ID: `YYMMDDHHMMSS` + zero-padded 3-digit millisecond.
Thread-safe: uses a lock and monotonically increments the ms field if the wall-clock
would produce a duplicate. Defined in `config.py: ID_STRFTIME`.

### `data/database.py › now_iso() -> str`

Returns current UTC time as ISO 8601: `"YYYY-MM-DDTHH:MM:SSZ"`.

### `data/database.py › get_current_user() -> str`

Returns `os.environ["USERNAME"]` (Windows) or `USER` / `LOGNAME`, lowercased. Falls back to `"unknown"`.
