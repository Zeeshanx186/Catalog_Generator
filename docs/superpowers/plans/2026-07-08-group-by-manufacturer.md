# Group Parts by Manufacturer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An optional, reversible "Group by manufacturer" toggle button on the Parts step that stably reorders the parts list so same-manufacturer rows are bundled, in order of first appearance.

**Architecture:** Pure client-side change in `ui/index.html` (single-file frontend). A snapshot of the pre-grouping order enables undo; structural list edits invalidate the snapshot. Backend, download order, duplicates and merge all follow list order already — untouched.

**Tech Stack:** Vanilla JS in ui/index.html; runtime verification via pywebview `evaluate_js` harness (venv python `myvenv\Scripts\python.exe`).

**Spec:** `docs/superpowers/specs/2026-07-08-group-by-manufacturer-design.md`

## Global Constraints

- Group key: `(p.manufacturer || '').trim().toUpperCase()`; empty string is a valid group key.
- Stable sort: manufacturers in order of FIRST appearance; rows within a group keep relative order.
- Button text exactly: `Group by manufacturer` ↔ `Restore BOM order`.
- Snapshot invalidated by `movePart`, `deletePart`, `addPartRow`, `discardAllDuplicates`, `startOver` — NOT by `updatePart` (cell text edits).
- No backend changes; do not rebuild the exe (user builds it themselves).

---

### Task 1: Group-by-manufacturer toggle (ui/index.html)

**Files:**
- Modify: `ui/index.html` — state object `const S = {` (~line 2398), Parts action bar (~line 2028, the "Add part" button), JS functions `movePart` / `deletePart` / `addPartRow` / `discardAllDuplicates` / `startOver` (~line 2971)
- Test: runtime verification script `tests/ui_verify_grouping.py` (create; follows the pattern of the duplicates runtime check)

**Interfaces:**
- Consumes: `S.parts`, `renderPartsTable()`, `refreshDuplicateMarks()` (runs inside renderPartsTable), `computeDuplicates(list)` from the duplicates feature.
- Produces: `toggleGroupByMfr()`, `updateGroupBtn()`, `invalidateGroupSnapshot()`, state field `S.groupSnapshot: Array|null`, button element `id="btn-group-mfr"`.

- [ ] **Step 1: Add state field**

In `const S = { ... }` add after `outputFolder: null,`:

```js
      groupSnapshot: null,   // pre-grouping order (array of part refs) or null
```

- [ ] **Step 2: Add the button to the Parts action bar**

Immediately AFTER the closing `</button>` of the "Add part" button (before `<div style="flex:1"></div>`):

```html
          <button class="btn btn-secondary btn-sm" id="btn-group-mfr" onclick="toggleGroupByMfr()"
            title="Bundle same-manufacturer parts together (in order of first appearance)">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="14" y2="12"/>
              <line x1="4" y1="18" x2="20" y2="18"/>
            </svg>
            <span id="btn-group-mfr-txt">Group by manufacturer</span>
          </button>
```

- [ ] **Step 3: Add the JS (above `renderPartsTable()` definition)**

```js
    // ── Group by manufacturer (optional, reversible) ───────────────────────────
    function toggleGroupByMfr() {
      if (S.groupSnapshot) {                    // restore original BOM order
        S.parts = S.groupSnapshot;
        S.groupSnapshot = null;
      } else {                                  // group, keeping a snapshot
        S.groupSnapshot = S.parts.slice();
        const firstSeen = new Map();
        S.parts.forEach((p, i) => {
          const k = (p.manufacturer || '').trim().toUpperCase();
          if (!firstSeen.has(k)) firstSeen.set(k, i);
        });
        S.parts = S.parts
          .map((p, i) => ({ p, i, g: firstSeen.get((p.manufacturer || '').trim().toUpperCase()) }))
          .sort((a, b) => a.g - b.g || a.i - b.i)
          .map(x => x.p);
      }
      renderPartsTable();
      updateGroupBtn();
    }

    function updateGroupBtn() {
      const t = document.getElementById('btn-group-mfr-txt');
      if (t) t.textContent = S.groupSnapshot ? 'Restore BOM order' : 'Group by manufacturer';
    }

    // Structural edits make the saved order stale — drop the undo, keep the
    // current order. Cell TEXT edits (updatePart) deliberately do NOT do this.
    function invalidateGroupSnapshot() {
      S.groupSnapshot = null;
      updateGroupBtn();
    }
```

- [ ] **Step 4: Invalidate the snapshot on structural edits**

Add `invalidateGroupSnapshot();` as the FIRST line inside each of:
`movePart(i, d)` (after the bounds check `if (j < 0 || j >= S.parts.length) return;`),
`deletePart(i)`, `addPartRow()`, `discardAllDuplicates()`, and `startOver()`.

- [ ] **Step 5: Runtime verification (failing first)**

Create `tests/ui_verify_grouping.py`:

```python
"""Runtime check of the group-by-manufacturer toggle via pywebview evaluate_js."""
import json, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parents[1]))
import webview
from api import API

html = (Path(__file__).parents[1] / "ui" / "index.html").read_text(encoding="utf-8")
api = API()
window = webview.create_window("grouping verify", html=html, js_api=api,
                               width=1100, height=700)
api.window = window
out = {}

def run():
    time.sleep(3)
    ev = window.evaluate_js
    ev("""
      S.parts = [
        {manufacturer:'RITTAL',    part_number:'A1', description:''},
        {manufacturer:'MOXA',      part_number:'B1', description:''},
        {manufacturer:'rittal ',   part_number:'A2', description:''},
        {manufacturer:'HONEYWELL', part_number:'C1', description:''},
        {manufacturer:'MOXA',      part_number:'B2', description:''},
        {manufacturer:'',          part_number:'D1', description:''},
        {manufacturer:'RITTAL',    part_number:'A1', description:'dup of A1'}
      ];
      renderPartsTable(); 'seeded'
    """)
    out["btn_before"] = ev("document.getElementById('btn-group-mfr-txt').textContent")
    ev("toggleGroupByMfr(); 'grouped'")
    out["grouped"] = ev("S.parts.map(p => p.part_number)")
    out["btn_grouped"] = ev("document.getElementById('btn-group-mfr-txt').textContent")
    # duplicate badge still points at first occurrence (A1 dup right after originals)
    out["dup_after_group"] = ev("computeDuplicates(S.parts)")
    ev("toggleGroupByMfr(); 'restored'")
    out["restored"] = ev("S.parts.map(p => p.part_number)")
    out["btn_restored"] = ev("document.getElementById('btn-group-mfr-txt').textContent")
    # structural edit invalidates the snapshot
    ev("toggleGroupByMfr(); deletePart(0); 'edited'")
    out["btn_after_edit"] = ev("document.getElementById('btn-group-mfr-txt').textContent")
    out["snapshot_after_edit"] = ev("S.groupSnapshot === null")
    print("RESULTS:" + json.dumps(out))
    window.destroy()

webview.start(run, private_mode=False)
```

Expected assertions (check by eye or pipe through python):
- `btn_before` = `Group by manufacturer`
- `grouped` = `["A1","A2","A1","B1","B2","C1","D1"]` (RITTAL block incl. case-insensitive `rittal ` and the duplicate, then MOXA, HONEYWELL, empty-mfr)
- `dup_after_group`: index of the second `A1` (position 2) points at 0; all others null
- `btn_grouped` = `Restore BOM order`
- `restored` = `["A1","B1","A2","C1","B2","D1","A1"]` (original order)
- `btn_restored` = `Group by manufacturer`
- `btn_after_edit` = `Group by manufacturer` and `snapshot_after_edit` = true

- [ ] **Step 6: Run BEFORE implementing (expect failure)**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\ui_verify_grouping.py`
Expected: `btn_before` = null / JS errors (button and functions don't exist yet).
(Write the test first, run it, then apply Steps 1-4, then re-run.)

- [ ] **Step 7: Run AFTER implementing (expect pass)**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\ui_verify_grouping.py`
Expected: RESULTS line matching all assertions in Step 5.

- [ ] **Step 8: Regression check + commit**

Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dedup.py` → OK
Run: `D:\APP\BomDownloader\myvenv\Scripts\python.exe tests\test_dup_pipeline.py` → OK

```bash
git add ui/index.html tests/ui_verify_grouping.py
git commit -m "feat: optional group-parts-by-manufacturer toggle with undo"
```
