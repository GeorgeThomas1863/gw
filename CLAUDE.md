# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

GrayWolfe is a Microsoft Access application for managing selectors (names, emails, phones, IPs, addresses, etc.) stored in SharePoint. It provides search across two data sources (GrayWolfe local database and the "S" system), data import with automatic type detection, and target management.

**Core problem:** Users collect thousands of selectors from disparate sources. GrayWolfe detects selector types, deduplicates them, and links related selectors into "targets" — connected groups representing a single entity. When new data arrives, it must merge into existing groups correctly (e.g., a new phone number linked to a known email should join that email's target, not create a new one). "Targets" here means selector groups, not military targets.

**IMPORTANT**: This is a COPY of an Access database. The `.cls` and `.bas` files here are exported text — they have no connection to the actual Access runtime. Edit files here; the user will import and test in Access.

## No Build/Test Commands

There are no build, lint, or test commands. VBA code runs only inside the Access database. The user handles all importing and testing.

## Architecture

### Module Numbering Convention

Modules are numbered by layer:

| Prefix | Layer | Files |
|--------|-------|-------|
| `mod1*` | **Workflow orchestration** | `mod1a_RunSearch.bas`, `mod1b_RunAddData.bas` |
| `mod2*` | **Data access** | `mod2a_Sharepoint_Tables.bas`, `mod2b_S.bas` |
| `mod3*` | **Data processing** | `mod3a_CleanFix.bas`, `mod3b_DetectCheck.bas` |
| `mod4*` | **Infrastructure** | `mod4a_APICalls.bas`, `mod4b_DefineThings.bas`, `mod4c_ErrorHandler.bas`, `mod4d_JsonParser.bas` |
| `mod5*` | **Utilities** | `mod5_UTIL_Delete.bas` |

### Form Hierarchy

- **`Form_frmMainMenu`** — Entry point. Two tabs: Search and Add Data.
- **`Form_frmResultsDisplay`** — Results viewer with S/GW tabs and dynamic filters.
  - `Form_frmResults_SSubform` — S results subform
  - `Form_frmResults_GWSubform` — GW results subform
- **`Form_frmSchemaDetection`** — Column type confirmation during default import.
- **`Form_frmTargetDetails`** — Target editor (metadata + selectors).
  - `Form_frmTargetDetails_Subform` — Target selectors subform
- **`Form_frmMergeTargets`** — Merge two targets (keep one, absorb the other). Opened from Target Details with pre-filled targetId.

### Data Flow

```
User Input → Cleaning (mod3a) → Detection (mod3b) → Validation → Local Tables → SharePoint Sync (mod2a)
```

Internally, all delimiters are normalized to `"!!"` for processing, then reconverted for display. All selector types are lowercased internally, proper-cased for display.

### Key Workflows

**Search:** `frmMainMenu.btnSearch_Click()` → `RunSearch()` (mod1a) → `SearchGrayWolfe()` + `SearchS()` → fills temp tables → opens `frmResultsDisplay`

**Default Import:** `frmMainMenu.btnAdd_Click()` → `RunDefaultImport()` (mod1b) → auto-detects column types → `FillTempSchema()` → opens `frmSchemaDetection` → user confirms → `RunAddSchemaData()` → creates selectors + targets → syncs to SharePoint

**Unrelated Import:** `RunUnrelatedImport()` (mod1b) → direct type detection, skips relationship creation

**Target Edit:** `frmTargetDetails.Form_Current()` → loads target data → user edits → `UpdateTargetSelectors()` / `UpdateTargetStatsForm()` → syncs to SharePoint

## Database Tables

| Table | Purpose |
|-------|---------|
| `localNorks`, `localSelectors`, `localTargets` | Local mirrors of SharePoint lists |
| `tempSchema` | Staging for schema detection during import |
| `tempGWSearchResults`, `tempSSearchResults` | Staging for search results |

SharePoint lists: **Norks**, **Selectors**, **Targets**

## Naming Conventions

### Functions

- `RunX()` — Workflow entry points
- `SearchX()` — Search operations
- `FillX()` — Insert/populate data
- `UpdateX()` — Modify existing data
- `CheckX()` — Validation (returns Boolean)
- `FixX()` — Data cleaning/normalization
- `DetectX()` — Type/pattern detection
- `ClearX()` — Delete/reset operations
- `BuildX()` — Construct strings/data
- `DefineX()` — Return config arrays/constants

### Variables

- `Str` suffix for strings: `inputStr`, `filterStr`
- `Arr` suffix for arrays: `searchArr`, `stateArr`
- `Rs` for recordsets: `rs`, `rsSearch`

## Selector Types

11 types in detection priority order: `email`, `phone`, `ip`, `address`, `linkedin`, `github`, `telegram`, `discord`, `name`, `other`, `row`

Type detection uses `VBScript.RegExp` and lives in `mod3b_DetectCheck.bas`. Each type has a `CheckX()` validator and a `FixXStr()` cleaner in `mod3a_CleanFix.bas`.

## Error Handling

Custom error system in `mod4c_ErrorHandler.bas` using `ThrowError(errCode, errMsg)`. Error codes 1950-1972 cover specific scenarios (1950=empty search input, 1966=missing S API token, 1968=S auth failed, 1972=merge targets failed, 1998=user cancellation). See the file for the full list.

SQL operations use `db.Execute strSQL, dbFailOnError`. Cleanup operations use `On Error Resume Next`.

## Key Technical Details

- **HTTP:** WinINet API (`mod4a_APICalls.bas`) for S API calls; MSXML2 for simpler requests
- **JSON:** Borrowed VBA-JSON library in `mod4d_JsonParser.bas` (~42KB, do not modify)
- **S API rate limiting:** 500-item batches with 10-second pauses between calls
- **ID generation:** Timestamp-based format `YYMMDDHHNNSSMMM` via `DefineUniqueId()`
- **Config arrays:** All defined in `mod4b_DefineThings.bas` (selector types, states, street suffixes, delimiters, form defaults, table/column mappings)
- **Mapping functions** in `mod4b_DefineThings.bas` use `Scripting.Dictionary`: `TableMap()`, `ColumnSearchMap()`, `ColumnAddMap()`, `DetectFunctionMap()`, `TargetFormDisplayMap()`, `StateMap()`

## What Works Well

Duplicate detection via `selectorClean` (normalized form stored alongside display form), regex-based type detection, and the data model (selectors → targets with nork metadata) are solid foundations. The cleaning/detection pipeline in mod3a/mod3b is reliable and extensible.

## Target Bridging & Shared Selectors

Key design decisions for how targets and selectors interact:

- **Import default = join existing target**: When imported selectors match an existing target, they join it (assume same entity).
- **Bridging merges targets**: When an imported row's selectors span 2+ existing targets, `AddRelatedRowTargets()` calls `MergeTargets()` to combine them into one. `CollectTargetIdsForRow()` gathers all matching targetIds first.
- **Shared selectors allowed**: The same selector value can belong to multiple targets (e.g., shared office phone). `FillLocalSelectors()` checks selectorClean+targetId when a targetId is provided (manual add via Target Details), but blocks any duplicate during import.
- **Search shows all targets**: `SearchGrayWolfe()` counts distinct targets from `tempGWSearchResults` after the search loop, so shared selectors surface all their targets.
- **UPDATE scope**: `UpdateSelectorsTblTargetId()` and `UpdateGWSearchTargetId()` only fill in blank targetIds, preserving existing target assignments on shared selectors.

## Known Architectural Gaps

These are confirmed issues, not speculative. Future changes should account for them rather than working around them unknowingly.

- **Search doesn't expand to target group**: `SearchGrayWolfe()` returns exact selector matches only. It does not return sibling selectors belonging to the same target, so the user must manually navigate to the target to see the full picture.
- **No import transaction safety**: Import operations (`RunAddSchemaData`, `FillLocalSelectors`, etc.) are not wrapped in transactions. A mid-import failure (network error, SharePoint timeout) can leave partial data in local tables without corresponding SharePoint records.
