# Duplicate BOM Parts (Keep / Discard) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Duplicate BOM rows survive parsing in BOM order, are visibly flagged in the UI with keep/discard controls, and kept duplicates reuse their original's downloaded document (no second search) at their own position in the merged PDF.

**Architecture:** The parser stops dropping repeated (manufacturer, part number) pairs. The frontend computes duplicate flags live from the parts table and sends `duplicate_of` (1-based index of the first occurrence) with each part. The download pipeline never searches duplicates; when an original finishes, each duplicate gets a copy of the original's PDF saved under the duplicate's own `NNN_` index prefix, which the existing folder-scan merge picks up automatically.

**Tech Stack:** Python 3.12 (`myvenv\Scripts\python.exe`), vanilla JS single-file UI (`ui/index.html`), no test framework — tests are plain-assert scripts under `tests/` run with the venv Python.

**Spec:** `docs/superpowers/specs/2026-07-07-duplicate-parts-design.md`

## Global Constraints

- Duplicate key: `manufacturer.trim().upper() + '|' + part_number.trim().upper()`; empty part numbers are never duplicates; later occurrences always point at the FIRST occurrence.
- `duplicate_of` is the **1-based** index of the original within the list actually sent to `start_download`.
- Copied file name: `{dup_idx:03d}_{sanitized 'MFR_PN'}__COPY_OF_{orig_idx:03d}.pdf` — must match the `_scan_folder` glob `f"{idx:03d}_*.pdf"`.
- Worker threads must never raise out of the duplicate-resolution code (log `[Error]` lines instead).
- Log strings may contain non-cp1252 characters ONLY via `tprint`/`self._log` (never bare `print`) — frozen build constraint.
- UI text: badge `DUP of #NNN`, status label `Found (duplicate)`, banner button `Discard all duplicates`.

---

### Task 1: Parser keeps duplicates

**Files:**
- Modify: `bom_downloader.py:410-416` (`_dedup`)
- Test: `tests/test_dedup.py` (create; also create empty `tests/` dir)

**Interfaces:**
- Produces: `_dedup(parts) -> list` now returns ALL valid rows in input order, including repeated (manufacturer, part_number) pairs. Invalid part numbers (`_looks_like_part` false) are still dropped. Callers (all four `_extract_*` functions) are unchanged.

- [ ] **Step 1: Write the failing test**

Create `tests/test_dedup.py`:

```python
import sys
sys.path.insert(0, r"D:\APP\BomDownloader")
import bom_downloader as bd

parts = [
    {"manufacturer": "RITTAL", "part_number": "8108.245", "description": "side panel"},
    {"manufacturer": "MOXA",   "part_number": "MB3270I-T", "description": ""},
    {"manufacturer": "rittal", "part_number": "8108.245", "description": "dup, case differs"},
    {"manufacturer": "RITTAL", "part_number": "??",        "description": "invalid pn"},
    {"manufacturer": "RITTAL", "part_number": "8108.245", "description": "second dup"},
]
out = bd._dedup(parts)

# All valid rows survive, in order — duplicates are KEPT
assert [p["part_number"] for p in out] == ["8108.245", "MB3270I-T", "8108.245", "8108.245"], out
# Invalid part numbers are still filtered
assert all(p["part_number"] != "??" for p in out)
print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dedup.py`
Expected: `AssertionError` (current `_dedup` drops the two duplicate rows).

- [ ] **Step 3: Implement**

In `bom_downloader.py`, replace the `_dedup` function:

```python
def _dedup(parts):
    """Filter out invalid part numbers but KEEP duplicates in BOM order —
    the UI flags them and lets the user keep or discard each one."""
    return [p for p in parts if _looks_like_part(p["part_number"])]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dedup.py`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add bom_downloader.py tests/test_dedup.py
git commit -m "feat: parser keeps duplicate BOM parts in order"
```

---

### Task 2: Backend duplicate resolution (api.py)

**Files:**
- Modify: `api.py` — `start_download` (progress part dicts, ~line 205), `_run` (~line 227: dup map, `_on_event`, submit loop, `as_completed` loop, after `_retry_misses`), `_retry_misses` (~line 454: exclude duplicates)
- Test: `tests/test_dup_pipeline.py` (create)

**Interfaces:**
- Consumes: parts dicts may carry `duplicate_of` (int 1-based, or None/absent) from the UI; Task 1's `_dedup` behavior.
- Produces: progress part dicts now include `"duplicate_of": int|None` (UI badges read this via `get_progress`). Duplicates are never submitted to the executor and never retried by `_retry_misses`. Resolved duplicates have status `found` with copied file(s) or `not_found`.

- [ ] **Step 1: Write the failing test**

Create `tests/test_dup_pipeline.py`. It monkeypatches `bd._download_part` with an offline fake (originals #1/#2 "find" a generated PDF, #4 finds nothing) and runs the real `api.API.start_download` pipeline:

```python
import sys, time, threading, tempfile, shutil
from pathlib import Path
sys.path.insert(0, r"D:\APP\BomDownloader")

import bom_downloader as bd
import api as api_mod

out_dir = Path(tempfile.mkdtemp(prefix="dup_test_"))

def _write_pdf(path):
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path)); c.drawString(100, 700, "dummy"); c.save()

def fake_download_part(idx, mfr, model, folder, max_dl, session, desc="", force=False):
    bd._emit("part_start", idx=idx, label=f"{mfr} {model}")
    if model == "NOPE-1":
        bd._emit("part_done", idx=idx, label=f"{mfr} {model}", status="not_found", files=[])
        return {"saved": [], "candidates": []}
    dest = Path(folder) / f"{idx:03d}_{mfr}_{model}.pdf"
    _write_pdf(dest)
    bd._emit("part_done", idx=idx, label=f"{mfr} {model}", status="found", files=[str(dest)])
    return {"saved": [dest], "candidates": []}

bd._download_part = fake_download_part

parts = [
    {"manufacturer": "RITTAL", "part_number": "8108.245", "description": ""},                       # 1 found
    {"manufacturer": "MOXA",   "part_number": "MB3270I-T", "description": ""},                      # 2 found
    {"manufacturer": "RITTAL", "part_number": "8108.245", "description": "", "duplicate_of": 1},    # 3 dup of 1
    {"manufacturer": "ACME",   "part_number": "NOPE-1",    "description": ""},                      # 4 not found
    {"manufacturer": "ACME",   "part_number": "NOPE-1",    "description": "", "duplicate_of": 4},   # 5 dup of 4
]

a = api_mod.API(); a.window = None
r = a.start_download(parts, {"output_folder": str(out_dir), "workers": 2,
                             "max_per_part": 1, "no_merge": True})
assert r["ok"], r

deadline = time.time() + 60
while time.time() < deadline:
    prog = a.get_progress()
    if prog.get("status") in ("done", "error"):
        break
    time.sleep(0.3)
assert prog["status"] == "done", prog["status"]

st = {p["idx"]: p for p in prog["parts"]}
assert st[1]["status"] == "found"
assert st[2]["status"] == "found"
# Duplicate of a found part: found, with a COPY file under its own prefix
assert st[3]["status"] == "found", st[3]
assert st[3]["files"] and "COPY_OF_001" in st[3]["files"][0], st[3]["files"]
assert Path(st[3]["files"][0]).is_file()
assert st[3]["duplicate_of"] == 1
# Duplicate of a not-found part: not found, no files
assert st[4]["status"] == "not_found"
assert st[5]["status"] == "not_found" and st[5]["files"] == []
# Counters include duplicates
assert prog["done"] == 5 and prog["found"] == 3 and prog["not_found"] == 2, \
    (prog["done"], prog["found"], prog["not_found"])
shutil.rmtree(out_dir, ignore_errors=True)
print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dup_pipeline.py`
Expected: FAIL — currently duplicates ARE searched (fake finds a doc for #3, so no
`COPY_OF_001` in its filename) and `duplicate_of` is missing from progress parts.

- [ ] **Step 3: Implement in api.py**

3a. In `start_download`, add `duplicate_of` to each progress part dict (the
`"parts": [...]` comprehension around line 205):

```python
                    {
                        "idx":          i + 1,
                        "manufacturer": p.get("manufacturer", ""),
                        "part_number":  p.get("part_number", ""),
                        "description":  p.get("description", ""),
                        "duplicate_of": (int(p["duplicate_of"])
                                         if p.get("duplicate_of") else None),
                        "status":       "pending",
                        "files":        [],
                    }
```

3b. In `_run`, right after the `bd.set_catalog_dirs(...)` call, build the
duplicate map (sanity-checked — a valid `duplicate_of` must point at an
EARLIER, non-duplicate part; anything else is treated as a normal part):

```python
        # Duplicates: dup_map[original_idx] = [duplicate_idx, ...] (all 1-based).
        # A duplicate is never searched — it reuses its original's document.
        dup_map: dict[int, list[int]] = {}
        dup_idxs: set[int] = set()
        for i, p in enumerate(parts):
            d = p.get("duplicate_of")
            try:
                d = int(d) if d else None
            except (TypeError, ValueError):
                d = None
            if d and 1 <= d < i + 1 and d not in dup_idxs:
                dup_map.setdefault(d, []).append(i + 1)
                dup_idxs.add(i + 1)
```

3c. In `_on_event`, inside the `part_start` branch, flip the original's
duplicates to "searching" too (they resolve when the original finishes).
After the existing `for p in ...: p["status"] = "searching"` loop add:

```python
                    for di in dup_map.get(ev["idx"], []):
                        if di in self._skipped:
                            continue
                        for p in self._progress["parts"]:
                            if p["idx"] == di and p["status"] == "pending":
                                p["status"] = "searching"
                                break
```

3d. Add a resolver closure in `_run`, after the `_tracked_download` definition:

```python
        def _resolve_duplicates(orig_idx: int, upgrade: bool = False) -> None:
            """Give orig's duplicates a copy of its document (or not_found).
            upgrade=True re-visits duplicates already marked not_found after
            the retry pass found the original late."""
            dups = dup_map.get(orig_idx, [])
            if not dups:
                return
            with self._lock:
                orig = next((p for p in self._progress["parts"]
                             if p["idx"] == orig_idx), None)
                src_files = [Path(f) for f in (orig.get("files") or [])] if orig else []
                orig_found = bool(orig and orig.get("status") == "found" and src_files)
            for d in dups:
                with self._lock:
                    dp = next((p for p in self._progress["parts"] if p["idx"] == d), None)
                    if dp is None:
                        continue
                    if upgrade:
                        if not (dp["status"] == "not_found" and not dp["files"]):
                            continue
                    elif dp["status"] not in ("pending", "searching"):
                        continue        # skipped by user or already resolved
                if upgrade and not orig_found:
                    continue
                files = []
                if orig_found:
                    safe = bd._sanitize(
                        f"{dp['manufacturer']}_{dp['part_number']}".strip("_").replace(" ", "_"))
                    for sf in src_files:
                        try:
                            if sf.is_file():
                                dest = folder / f"{d:03d}_{safe}__COPY_OF_{orig_idx:03d}.pdf"
                                shutil.copyfile(sf, dest)
                                files.append(str(dest))
                        except Exception as e:
                            with self._log_lock:
                                self._log.append(
                                    f"    [Error] duplicate #{d:03d} copy failed: {e!r}")
                st = "found" if files else "not_found"
                with self._lock:
                    dp["status"] = st
                    dp["files"]  = files
                    if upgrade:
                        self._progress["found"]     += 1
                        self._progress["not_found"] -= 1
                    else:
                        self._progress["done"] += 1
                        self._progress["found" if st == "found" else "not_found"] += 1
                with self._log_lock:
                    if files:
                        self._log.append(
                            f"    [Duplicate] #{d:03d} reuses #{orig_idx:03d}'s document")
                    else:
                        self._log.append(
                            f"    [Duplicate] #{d:03d}: original #{orig_idx:03d} "
                            f"found nothing - marked not found")
```

3e. Submit ONLY non-duplicates. Change the `fut_map` comprehension to skip
duplicates:

```python
            fut_map = {
                ex.submit(
                    _tracked_download,
                    i + 1,
                    p.get("manufacturer", ""),
                    p.get("part_number", ""),
                    folder,
                    max_dl,
                    self._session,
                    p.get("description", ""),
                ): i + 1
                for i, p in enumerate(parts)
                if (i + 1) not in dup_idxs
            }
```

3f. In the `as_completed` loop, resolve duplicates as each original lands.
After `part_results[idx] = fut.result()` AND at the end of the `except`
branch (both paths), add:

```python
                _resolve_duplicates(idx)
```

3g. Duplicates never produce `part_results` entries — give them empty ones so
downstream code that indexes `part_results` can't KeyError. After the
`as_completed` loop (before the `while not self._cancel_ev.is_set():` block):

```python
            for d in dup_idxs:
                part_results.setdefault(d, {"saved": [], "candidates": []})
```

3h. After the `self._retry_misses(parts, folder, max_dl)` call in `_run`, add
the late-upgrade pass:

```python
        # Retry pass may have found originals late — propagate to their dups.
        for oi in dup_map:
            _resolve_duplicates(oi, upgrade=True)
```

3i. In `_retry_misses`, exclude duplicates from the misses list:

```python
            misses = [p for p in self._progress["parts"]
                      if p.get("status") == "not_found"
                      and not p.get("duplicate_of")]
```

- [ ] **Step 4: Run tests**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dup_pipeline.py`
Expected: `OK`
Also re-run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dedup.py` → `OK`

- [ ] **Step 5: Commit**

```bash
git add api.py tests/test_dup_pipeline.py
git commit -m "feat: duplicates reuse their original's document instead of searching"
```

---

### Task 3: Parts-step UI — live duplicate flags, banner, discard

**Files:**
- Modify: `ui/index.html` — CSS (append before the closing `</style>` tag), banner HTML (before `<!-- Parts table -->` `<div class="table-wrap">` at ~line 1985), JS (`renderPartsTable`/`makePartRow` ~line 2519, `updatePart` ~line 2551, `deletePart`, `addPartRow`, `startDownload` ~line 2629)

**Interfaces:**
- Consumes: nothing new from backend (flags are computed client-side).
- Produces: `computeDuplicates(list) -> (int|null)[]` (0-based original index per row), `refreshDuplicateMarks()`, `discardAllDuplicates()`, `dupBadge(origIdx1Based)` returning badge HTML — Task 4 reuses `dupBadge`. `startDownload` sends `duplicate_of` (1-based, computed on the filtered `validParts`).

- [ ] **Step 1: Add CSS**

Insert immediately before the closing `</style>` tag (verify it is unique first:
`grep -c "</style>" ui/index.html` → must print `1`):

```css
    /* ── Duplicate part flags ── */
    tr.dup-row td { background: rgba(245, 158, 11, .06); }
    tr.dup-row td.idx { box-shadow: inset 3px 0 0 #f59e0b; }
    .dup-badge {
      display: inline-block; margin-left: 6px; padding: 1px 6px; border-radius: 4px;
      font-size: 10px; font-weight: 700; letter-spacing: .4px; white-space: nowrap;
      color: #f59e0b; background: rgba(245, 158, 11, .12);
      border: 1px solid rgba(245, 158, 11, .35); vertical-align: middle;
    }
    .dup-banner {
      display: flex; align-items: center; gap: 12px; margin: 0 0 12px;
      padding: 10px 14px; border: 1px solid rgba(245, 158, 11, .35);
      border-radius: 8px; background: rgba(245, 158, 11, .08); font-size: 13px;
    }
```

- [ ] **Step 2: Add the banner HTML**

Directly above `<!-- Parts table -->` / `<div class="table-wrap">` (inside
`#panel-parts`):

```html
        <div class="dup-banner" id="dup-banner" style="display:none">
          <span id="dup-banner-text"></span>
          <div style="flex:1"></div>
          <button class="btn btn-secondary btn-sm" onclick="discardAllDuplicates()">Discard all duplicates</button>
        </div>
```

- [ ] **Step 3: Add JS helpers + wire rendering**

Add above `renderPartsTable()`:

```js
    // ── Duplicate detection (live) ─────────────────────────────────────────────
    function dupKeyOf(p) {
      const pn = (p.part_number || '').trim().toUpperCase();
      if (!pn) return null;                       // empty PN is never a duplicate
      return (p.manufacturer || '').trim().toUpperCase() + '|' + pn;
    }

    // Returns per-row: null (normal / first occurrence) or the 0-based index
    // of the FIRST occurrence this row duplicates.
    function computeDuplicates(list) {
      const seen = new Map();
      return list.map((p, i) => {
        const k = dupKeyOf(p);
        if (k === null) return null;
        if (seen.has(k)) return seen.get(k);
        seen.set(k, i);
        return null;
      });
    }

    function dupBadge(origIdx1Based) {
      return `<span class="dup-badge">DUP of #${String(origIdx1Based).padStart(3, '0')}</span>`;
    }

    function refreshDuplicateMarks() {
      const dups = computeDuplicates(S.parts);
      const rows = document.querySelectorAll('#parts-tbody tr');
      rows.forEach((tr, i) => {
        const d = dups[i];
        tr.classList.toggle('dup-row', d !== null);
        let badge = tr.querySelector('.dup-badge');
        if (d !== null) {
          if (!badge) {
            badge = document.createElement('span');
            badge.className = 'dup-badge';
            tr.querySelector('td.pn').appendChild(badge);
          }
          badge.textContent = `DUP of #${String(d + 1).padStart(3, '0')}`;
        } else if (badge) {
          badge.remove();
        }
      });
      const n = dups.filter(d => d !== null).length;
      const banner = document.getElementById('dup-banner');
      banner.style.display = n ? 'flex' : 'none';
      if (n) document.getElementById('dup-banner-text').textContent =
        `${n} duplicate part${n !== 1 ? 's' : ''} found — kept duplicates reuse the document of their first occurrence.`;
    }

    function discardAllDuplicates() {
      const dups = computeDuplicates(S.parts);
      S.parts = S.parts.filter((_, i) => dups[i] === null);
      renderPartsTable();
      document.getElementById('parts-count').textContent = S.parts.length;
      toast('Duplicates discarded.', 'ok');
    }
```

At the END of `renderPartsTable()` (after the forEach) add:

```js
      refreshDuplicateMarks();
```

In `updatePart(i, field, val)` add a live re-check after the assignment
(NO table rebuild — it would steal focus from the input):

```js
    function updatePart(i, field, val) {
      if (S.parts[i]) S.parts[i][field] = val;
      if (field !== 'description') refreshDuplicateMarks();
    }
```

- [ ] **Step 4: Send duplicate_of from startDownload**

In `startDownload()`, replace the `validParts` line with:

```js
      const validParts = S.parts
        .filter(p => p.part_number && p.part_number.trim())
        .map(p => ({ ...p }));
      const dupIdx = computeDuplicates(validParts);
      validParts.forEach((p, i) => {
        p.duplicate_of = dupIdx[i] === null ? null : dupIdx[i] + 1;
      });
```

- [ ] **Step 5: Smoke-test in dev mode**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe main.py`
- Add parts manually (Add part): `RITTAL / 8108.245` twice and `MOXA / X1` once.
- Expect: second RITTAL row amber with `DUP of #001` badge; banner says
  "1 duplicate part found…".
- Edit the second row's PN to `8108.246` → badge and banner disappear while typing.
- Change it back → badge returns. Click **Discard all duplicates** → row removed,
  banner hidden, count updates.
Close the app.

- [ ] **Step 6: Commit**

```bash
git add ui/index.html
git commit -m "feat: flag duplicate parts in the review table with keep/discard"
```

---

### Task 4: Download & results badges

**Files:**
- Modify: `ui/index.html` — `buildDlGrid` (~line 2654), `pollProgress` status label (~line 2703), `resRowHTML` (~line 3025)

**Interfaces:**
- Consumes: `dupBadge()` from Task 3; `p.duplicate_of` on progress parts from Task 2; `buildDlGrid(validParts)` receives parts that already carry `duplicate_of` (Task 3).

- [ ] **Step 1: Badge on download cards**

In `buildDlGrid`, change the part-number line of the card template from:

```js
      <div class="card-pn">${esc(p.part_number)}</div>
```

to:

```js
      <div class="card-pn">${esc(p.part_number)}${p.duplicate_of ? ' ' + dupBadge(p.duplicate_of) : ''}</div>
```

- [ ] **Step 2: Status label for resolved duplicates**

In `pollProgress`, after the line
`card.querySelector('.card-status-txt').textContent = labels[p.status] || p.status;`
add:

```js
        if (p.duplicate_of && p.status === 'found')
          card.querySelector('.card-status-txt').textContent = 'Found (duplicate)';
```

- [ ] **Step 3: Badge on the results table**

In `resRowHTML`, change:

```js
      <td class="pn">${esc(p.part_number)}</td>
```

to:

```js
      <td class="pn">${esc(p.part_number)}${p.duplicate_of ? ' ' + dupBadge(p.duplicate_of) : ''}</td>
```

- [ ] **Step 4: Smoke-test in dev mode**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe main.py`
Add `RITTAL / 8108.245` twice, start the download (settings: default output
folder, catalog empty). Expect:
- Card #002 shows the `DUP of #001` badge; while #001 searches, #002 shows "Searching…".
- When #001 lands, #002 flips to **Found (duplicate)** almost instantly, with an Open button.
- Results page: row 002 badge + ✓ Found; output folder contains
  `002_RITTAL_8108.245__COPY_OF_001.pdf`; merged PDF has the document at both positions.
Close the app.

- [ ] **Step 5: Commit**

```bash
git add ui/index.html
git commit -m "feat: duplicate badges on download cards and results"
```

---

### Task 5: Portable build + frozen verification

**Files:**
- No source changes. Rebuild `dist\BomDownloader\` + `dist\BomDownloader-Portable.zip`.

**Interfaces:**
- Consumes: everything above; `BomDownloader.spec`, `build.bat` (already in repo).

- [ ] **Step 1: Re-run both test scripts**

```
D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dedup.py
D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dup_pipeline.py
```
Expected: `OK` twice.

- [ ] **Step 2: Rebuild the portable app**

Ensure no `BomDownloader.exe` is running (`tasklist | findstr BomDownloader`),
then:

```
D:\APP\BomDownloader\myvenv\Scripts\pyinstaller.exe BomDownloader.spec --noconfirm
```
Expected: `Build complete!` and `dist\BomDownloader\BomDownloader.exe` exists.

- [ ] **Step 3: Frozen smoke test**

Launch `dist\BomDownloader\BomDownloader.exe`. Repeat the Task 4 smoke test
(two identical RITTAL parts). Expect identical behavior to dev mode. Close the app.

- [ ] **Step 4: Refresh the portable ZIP**

```powershell
Remove-Item dist\BomDownloader-Portable.zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path dist\BomDownloader -DestinationPath dist\BomDownloader-Portable.zip
```

- [ ] **Step 5: Final commit**

```bash
git add -A
git status   # verify only intended files are staged (dist/ and build/ are gitignored)
git commit -m "chore: rebuild portable app with duplicate-part support"
```
