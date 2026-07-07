# Group Parts by Manufacturer — Design

**Date:** 2026-07-08
**Status:** Approved (user)

## Problem

Users want an OPTIONAL way to bundle same-manufacturer parts together in the
parts list (and therefore in the merged record book): if RITTAL appears first,
all remaining RITTAL rows follow it, then the next manufacturer's rows, etc.
It must be a button the user chooses to click — never automatic — and it must
be reversible.

## Approach

Pure client-side feature in `ui/index.html` (Parts step). No backend changes:
download order, duplicate resolution, merge and TOC all follow the list order
already.

## Behavior

- New button **“Group by manufacturer”** in the Parts-step action bar next to
  “Add part”.
- Click → stable sort of `S.parts` by manufacturer **in order of first
  appearance**:
  - Group key: `(p.manufacturer or '').trim().toUpperCase()` (empty string is
    a valid key — empty-manufacturer rows form their own group).
  - Manufacturers keep the order in which they first appear in the current
    list; rows within a group keep their relative order (stable).
- After grouping the button reads **“Restore BOM order”**; clicking it
  restores the exact pre-grouping order from a saved snapshot.
- Snapshot invalidation: any STRUCTURAL list change while grouped —
  `movePart`, `deletePart`, `addPartRow`, `discardAllDuplicates` — keeps the
  current order but discards the snapshot and resets the button label to
  “Group by manufacturer”. Cell text edits (`updatePart`) do NOT invalidate
  the snapshot (same objects, same order; the snapshot holds references, so
  edits appear in both orders).
- Grouping again after edits simply takes a fresh snapshot of the current
  order.
- Loading a new BOM / start-over clears the snapshot and resets the label.

## Interaction with duplicates

The sort is stable and duplicates share their original's manufacturer key, so
a duplicate can never move ahead of its first occurrence. Badges recompute on
re-render (`renderPartsTable` → `refreshDuplicateMarks`), and `duplicate_of`
is computed at send time in `startDownload` from the final order — no changes
needed.

## Implementation sketch

State: `S.groupSnapshot = null` (array of part references or null).

```js
function toggleGroupByMfr() {
  if (S.groupSnapshot) {                      // restore
    S.parts = S.groupSnapshot; S.groupSnapshot = null;
  } else {                                    // group
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
function invalidateGroupSnapshot() { S.groupSnapshot = null; updateGroupBtn(); }
```

`updateGroupBtn()` sets the button text/icon from `S.groupSnapshot`.
`invalidateGroupSnapshot()` is called from `movePart`, `deletePart`,
`addPartRow`, `discardAllDuplicates`, and `startOver`/BOM-load reset.

## Testing

Runtime verification via pywebview `evaluate_js` (same harness as the
duplicates feature): seed a mixed list (RITTAL, MOXA, RITTAL, HONEYWELL,
MOXA, empty-mfr), assert grouped order = [R,R,M,M,H,empty] with stable
within-group order; restore returns the original array; a `deletePart` while
grouped resets the button; duplicate badges still point at the first
occurrence after grouping.
