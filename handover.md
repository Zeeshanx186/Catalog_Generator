# BOM Downloader — Agent Handover

_Last updated: 2026-07-09_

## What this project is
`bom_downloader.py` is a Python CLI that reads a BOM (Excel/PDF/Markdown), and for
each part searches manufacturer sources + the web to download the correct
datasheet/manual, verifies each PDF's content, and merges them into one ordered PDF.

**Goal right now:** improve find-rate and reliability so the app can pass an
**enterprise approval** demo. `BOM.xlsx` (219 parts) is the reference test set.

- Platform: Windows 11, PowerShell primary (Bash/Git-Bash also available).
- Run env: a venv at `D:\APP\BomDownloader\myvenv` exists.
- Main file is ~4000 lines; tests in `tests/test_header_and_probe.py`.

---

## ⚠️ Current blocker (read first)
**The trial Zyte API key is exhausted.** As of 2026-07-09 it returns **HTTP 421**
and trips both the `Zyte` (SERP) and `Zyte-fetch` circuit breakers mid-run. When
Zyte dies and DDG is rate-limited, the cascade fast-skips parts ("search engines
cooling down — skipping"). **Any full run in this state produces an INVALID
find-rate** (last run: 118/219 raw, but ~53 are false negatives from the dead key).

**Do not trust find-rate numbers until a live search backend is in place.** Two paths:
1. Wait for the trial key to recover (if it's rate-limiting, not credit-exhaustion).
2. **SerpApi $150 plan** (user is buying it). `_serpapi` is already wired in
   `_BACKENDS` **ahead of** Zyte — just set `SERPAPI_KEY` and it activates
   automatically, no code change.

**Never run concurrent batches against one trial key** — that accelerated the
exhaustion. Verify the key with ONE probe before launching a full run.

---

## What was done this session (all committed to `main`)
```
454344e fix: transient Zyte errors no longer disable it (and skip parts) for the run
914bd8f fix: generic brochures no longer short-circuit the search (HIMA rescue)
8150121 perf: route site-scoped searches through the cascade, off the DDG lock
8ac1f64 fix: date in manufacturer column no longer leaks into search query
76ec989 perf: per-part wall-clock deadline + faster Zyte SERP timeout
9a1b49f feat: Zyte SERP/fetch integration + 6 search-accuracy bug fixes
```
Details:
- **Zyte integration** — `_zyte_serp` (Google SERP) + `_zyte_fetch` (unblock blocked
  downloads). Key read from `os.getenv("ZYTE_API_KEY")` — **never hardcode it.**
- **6 earlier accuracy fixes** — phantom-PDF probe (Weidmüller CDN serving HTML on
  .pdf), header-row leak, 20-min Retry-After hang, Cyrillic homoglyph part numbers,
  family-prefix verifier bug, brochure short-circuit v1.
- **Per-part deadline** — `PART_HARD_DEADLINE=150s` bounds the whole part
  (find+download+verify+rescue); `PART_SEARCH_BUDGET=75s` bounds finding only.
- **Date-in-manufacturer fix** — `_is_datelike()`; a datetime cell in the maker
  column (Excel auto-format) no longer becomes the search query prefix.
- **DDG-lock relief** — the `site:` / HIMA / Rittal-portal / manual-pairing searches
  now go through `_search()` (API backend first, DDG last) instead of pinning to the
  serialized DDG lock.
- **Brochure short-circuit v2** — `_own_doc_identifies()`: an own-domain PDF only
  short-circuits the search if it identifies the part (part# in URL/title, or an
  explicit manual/datasheet name). Generic marketing brochures no longer suppress the
  HIMA mirror tier.
- **Circuit-breaker resilience** — `record_soft()` + `SOFT_TRIP_THRESHOLD=6`:
  transient Zyte errors (520/timeout) don't disable it for the whole run; only
  decisive auth/credit (401/402/403) trip immediately.

---

## How the search works (mental model)
`_find_pdfs(model, mfr, ...)` walks tiers, accumulating candidates in `direct`:
1. **DirectProbe** — hardcoded CDN URL patterns (DigiKey media, Rockwell literature, Rittal pdf-creator).
2. **MfrPortal** — scrape the maker's own product/doc pages.
3. **Nexar** component DB (needs key, unset).
4. **SiteSearch** — `site:makerdomain "part"` via `_search()`. Short-circuits only on an *identifying* own-site doc (see `_own_doc_identifies`).
5. **HIMA tier** — family-aware queries (`_hima_queries`) → third-party mirrors (meloautotech, quicktimeonline, eic2, plc-module). Runs when nothing identifying is in hand.
6. **Distributor scrapers** (Mouser, RS — often bot-blocked) + **broad web search**.

`_search(q)` tries `_BACKENDS` in order and returns the first with results:
`SearXNG → Google CSE → Tavily → Exa → Serper → SerpApi → Zyte → DDG`.
Only DDG and DirectProbe are always-on; the rest need keys. **SerpApi sits ahead of Zyte.**

Verification: every downloaded PDF is opened and checked (`_verify_pdf`) — the part
number / doc-type must be present, else it's rejected as "different part."

### Gotchas that will bite you
- **CircuitBreaker**: `TRIP_THRESHOLD=3` (decisive), `SOFT_TRIP_THRESHOLD=6` (transient).
  A tripped backend **does not recover** for the rest of the run.
- **Fail-fast**: when DDG is paused AND no API backend is ready, parts are skipped
  ("cooling down — skipping, retry later"). **There is NO automatic retry pass** —
  `reset_search_throttle()` is defined but never called in the CLI flow. Skipped
  parts land in the interactive NOT-FOUND list only.
- **Zyte SERP** needs a **bare `?q=` URL** — extra params (num=, hl=) cause 520
  "website ban". 520s are intermittent → one retry built in.
- **Zyte can't unblock Mouser/DigiKey** (they serve challenge stubs to Zyte IPs);
  dex.cz / HIMA mirrors fetch fine.
- **Bing Search API is RETIRED** (2025-08-11, HTTP 410). `_bing` is dead code — do
  not suggest buying it. See `BING_API_KEY` comment in config.
- **End-of-run interactive prompt**: the app asks for input after processing. In a
  background run, redirect `< /dev/null` so it EOFs and exits cleanly. The
  DOWNLOAD SUMMARY prints *before* the prompt, so you still get the tally.

---

## Running & measuring
Full run (background, clean exit, UTF-8 log):
```bash
cd /d/APP/BomDownloader
export ZYTE_API_KEY="<key>" PYTHONUTF8=1 PYTHONIOENCODING=utf-8
python -u bom_downloader.py BOM.xlsx < /dev/null > scratchpad/run.log 2>&1
```
Direct parts (skip BOM parse): `python bom_downloader.py -P "HIMA" "F-CPU 01" -P "MOXA" "EDS-316-SS-SC-T"`

Tests: `python tests\test_header_and_probe.py` → prints `OK / OK2 / OK3 / OK4 / OK5`.

Count found/missed from a log (PowerShell — note `-> \d+ file` double-counts `0 files`,
so compute found = processed − missed):
```powershell
$g = Get-Content <log> | Select-String '^\[\d+/219\]'
$missed = ($g | ?{$_ -match '0 saved|0 files|no URLs found'}).Count
"Processed $($g.Count)  Found $($g.Count-$missed)  Missed $missed"
```

---

## Results so far
- **Old baseline (working Zyte): ~69%** (151/219).
- **Last run (dead key): 118/219 raw — INVALID**, ~53 false-negative skips.
- **Validated when Zyte was alive:** smoke test 3/3 (AB, MOXA, HIMA F6217 all verified
  MANUAL); isolated HIMA test found F-BASE RACK 01's 56-page HIMatrix Engineering
  Manual (eic2.com).
- **Schneider proof of the fixes:** A9F44xxx block went from *fully skipped (1/10)* in
  the broken run to *properly searched (5/10)* with no wholesale skips — even on a dead key.

---

## Suggested next steps
1. **Get a live search backend** (SerpApi key → set `SERPAPI_KEY`, or recovered trial
   Zyte), then run **one clean full 219** for a real, defensible find-rate.
2. **Hard buckets that need the backend:** HIMA F-modules (portal-walled; only public
   copies are on third-party mirrors that Google/SerpApi surfaces, not DDG). Many
   docs are permission-protected → need Zyte-fetch / equivalent to download.
3. **Genuinely portal-walled / no-public-doc parts** (accept as misses): niche valve &
   instrument makers — AUTOCLAVE, OLIVER, STEWART BUCHANAN, TESCOM, BUTECH, HYSTAT,
   PROSERV, etc.
4. **HONEYWELL FC-PDB-0824P**: exists on `mooreautomated.com` (a distributor product
   *page*, not a direct .pdf). Would need a page-spider / distributor scraper, not a
   direct-link search. Candidate feature.
5. **Cosmetic**: the DOWNLOAD SUMMARY banner prints a garbled separator (`='*W`) — a
   format-string bug, harmless, ~2-min fix if polishing for the demo.

---

## Repo / files
- `bom_downloader.py` — the app (uncommitted work: none; tree clean).
- `tests/test_header_and_probe.py` — regression tests (bugs 1–6).
- `BOM.xlsx` — 219-part reference BOM (untracked; do not commit blindly).
- `scratchpad/` (session temp) — run logs: `run_fresh.log` is the last full run.
- Agent memory lives in the Claude project memory dir (`MEMORY.md` + per-fact files),
  notably `zyte-trial-key.md` (key status & constraints) and
  `bomdownloader-portable-build.md` (build via `BomDownloader.spec`/`build.bat`, not
  bare pyinstaller).

## Hard rules
- Keep `ZYTE_API_KEY` in the env var — **never commit it hardcoded**.
- One trial key = one run at a time. No concurrent batches.
- Commit messages end with the Co-Authored-By trailer; branch off `main` before
  committing if the user hasn't asked to commit directly to it.
