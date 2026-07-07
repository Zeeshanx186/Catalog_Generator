# Duplicate BOM Parts — Keep / Discard Design

**Date:** 2026-07-07
**Status:** Approved (user)

## Problem

The BOM parser silently deletes duplicate parts (`_dedup` in
`bom_downloader.py` drops any repeated manufacturer + part-number pair).
Users never see that a BOM row disappeared. Requirement: keep duplicates in
the list **in their original BOM position**, visibly flag them, and let the
user keep or discard each one. Kept duplicates must not trigger a second web
search — they reuse the document found for their first occurrence, and the
merged PDF repeats that document at the duplicate's BOM position.

## Approach (chosen: UI-computed duplicates, live)

Duplicate detection lives in the frontend and recomputes on every table
change, so flags stay correct when the user edits part numbers, deletes rows,
adds rows, or reorders. The backend stops dropping duplicates and gains one
behavior: a part marked as a duplicate is not searched; it inherits its
original's files.

Rejected alternatives: parse-time flagging only (flags go stale after edits);
keeping the silent dedup with a notice (cannot keep individual duplicates).

## Changes

### 1. Parser — `bom_downloader.py`

`_dedup(parts)` keeps its invalid-part filtering (`_looks_like_part`) but no
longer removes repeated (manufacturer, part_number) pairs. All valid rows are
returned in BOM order. (Function keeps its name; docstring updated.)

### 2. Parts step UI — `ui/index.html`

- `computeDuplicates()`: walks `S.parts`, key = `(manufacturer or '').trim().toUpperCase() + '|' + (part_number or '').trim().toUpperCase()`
  (empty part numbers are never duplicates). First occurrence of a key is the
  original; later occurrences get `dupOf = <0-based index of original>`.
  Runs inside `renderPartsTable()` (move/delete/add already re-render). For
  typing (`updatePart`, fired oninput) do NOT rebuild the table — that would
  steal focus from the input being edited. Instead call
  `refreshDuplicateMarks()`, which recomputes flags and updates each row's
  class/badge in place.
- Duplicate rows: amber left border + faint amber row tint + badge
  `DUP of #NNN` (NNN = original's displayed 1-based number) next to the part
  number. Existing per-row delete button doubles as "discard"; keeping means
  simply leaving the row.
- Banner above the table, visible only when duplicates exist:
  "N duplicate part(s) found — kept duplicates reuse the document of their
  first occurrence." with one button: **Discard all duplicates** (splices all
  flagged rows, re-renders).
- `startDownload()` sends each part with `duplicate_of` = original's
  **1-based** index (or absent/null for normal parts), computed at send time
  from the final table state.

### 3. Download pipeline — `api.py`

In `_run()`:

- Split parts into originals and duplicates by `duplicate_of`.
- Submit only originals to the executor (unchanged path).
- Duplicates are marked status `"searching"`→ resolved when their original's
  future completes (handled in the `as_completed` loop):
  - Original found → for each saved file of the original, copy it into the
    output folder under the duplicate's own index prefix using the existing
    naming scheme (e.g. `007_RITTAL_8108.245__COPY_OF_003.pdf`), record the
    copies in the duplicate's `files`, set status `"found"`, log
    `[Duplicate] #007 reuses #003's document`.
  - Original not found / crashed → duplicate becomes `"not_found"`, log line
    says so. (The retry pass already re-runs `not_found` parts one at a time;
    duplicates are excluded from `_retry_misses` by checking `duplicate_of`.)
  - Original skipped by user → duplicate resolves `"not_found"` (logged).
- Progress counters (`done`/`found`/`not_found`) updated for duplicates the
  same as normal parts.
- Merge needs **no changes**: `_collect_files()` scans the folder by index
  prefix, so the copied file appears at the duplicate's BOM position in
  `BOM_Manuals_Merged.pdf` / `BOM_Manuals_Selected.pdf` automatically, with
  its own divider page.
- `skip_part` on a duplicate works as for any pending part (status
  `"skipped"`, no copy made). `requeue_part`/`research_part` on a duplicate
  treats it as a normal part (independent search) — acceptable escape hatch.

### 4. Download/results UI — `ui/index.html`

Duplicate cards show their badge (`DUP of #NNN`) on the download grid and the
results list. Status label for a resolved duplicate reads `Found (duplicate)`.
Everything else (checkboxes, merge selection) behaves as for normal parts.

## Edge cases

- Duplicate whose original was discarded in the Parts step: it is now the
  first occurrence → treated as a normal part (recompute handles this).
- Chains (A, dup B→A, dup C→A): every later occurrence points at the FIRST
  occurrence, never at another duplicate.
- Original finishes `found` but with zero saved files (shouldn't happen):
  duplicate falls back to `not_found`.
- Copy failure (locked file, disk full): duplicate becomes `not_found` with an
  `[Error]` log line; never crashes the worker loop.
- Cancelled run: unresolved duplicates follow the existing rule (pending/
  searching → `not_found`).

## Testing

- Unit-style: `_dedup` returns repeated pairs in order; still filters invalid
  part numbers.
- Pipeline (frozen-compatible harness, as used for previous fixes): BOM of
  4 parts where #3 duplicates #1 → expect #1 searched once, #3 `found` with a
  copied file under `003_` prefix, merged PDF contains the document at both
  positions; a duplicate of a not-found part resolves `not_found`.
- UI smoke: duplicate badge appears, editing the PN away removes it, Discard
  all removes flagged rows only.
