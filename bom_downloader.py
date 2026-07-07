"""
BOM Manual Downloader
=====================
Downloads datasheets/manuals for every part in a BOM,
shows a summary, lets you resolve missing items interactively,
then merges everything into one ordered PDF.

Supported BOM formats : PDF, Excel (.xlsx / .xls), Markdown (.md)
Active search engines  : DirectProbe (manufacturer CDNs) + DuckDuckGo

Usage
-----
  # Load a BOM file:
  python bom_downloader.py BOM.pdf
  python bom_downloader.py BOM.xlsx
  python bom_downloader.py BOM.md

  # Enter parts directly (skip BOM extraction):
  python bom_downloader.py -P RITTAL 8808000 -P "ALLEN BRADLEY" 1756-A4K

  # Interactive startup menu (no args needed):
  python bom_downloader.py

Requirements
------------
  pip install ddgs requests beautifulsoup4 pdfplumber pypdf reportlab openpyxl
"""

import re, sys, os, time, argparse, threading, tempfile, shutil
from pathlib import Path
from urllib.parse import urljoin, quote
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from datetime import datetime

# ── Dependency check ──────────────────────────────────────────────────────────
_MISSING = []
# pyrefly: ignore [missing-import]
try: import pdfplumber
except ImportError: _MISSING.append("pdfplumber")
try: import requests; from bs4 import BeautifulSoup
except ImportError: _MISSING.append("requests beautifulsoup4")
# pyrefly: ignore [missing-import]
try: from ddgs import DDGS
except ImportError: _MISSING.append("ddgs")
try: from pypdf import PdfWriter, PdfReader
except ImportError: _MISSING.append("pypdf")
try: from reportlab.lib.pagesizes import A4; from reportlab.pdfgen import canvas as rl_canvas
except ImportError: _MISSING.append("reportlab")
try:
    from reportlab.lib.utils import simpleSplit
    from reportlab.pdfbase.pdfmetrics import stringWidth
except ImportError:
    pass
try: import openpyxl
except ImportError: _MISSING.append("openpyxl")
if _MISSING:
    print("Missing packages. Run:")
    print(f"  pip install {' '.join(_MISSING)}")
    sys.exit(1)

# =============================================================================
#  CONFIG  -  fill in API keys to enable extra search backends
# =============================================================================

DOWNLOAD_FOLDER  = "manuals"
MAX_PER_PART     = 1
DEFAULT_WORKERS  = 5
SCRAPE_WORKERS   = 4
PART_SEARCH_BUDGET = 75      # hard cap (s) on candidate-FINDING per part in batch
                             # runs. Whatever was found by then gets ranked and
                             # downloaded normally. Prevents any single part (e.g.
                             # Rittal articles whose manual lookup needs a slow
                             # search engine) from grinding for many minutes.
                             # Set to 0/None to disable.
RITTAL_TIMING    = True      # print per-phase timing for Rittal parts to the
                             # activity log: datasheet-sniff vs DDG-lookup vs
                             # total. Set False once diagnosis is done.

# ── Document-type verification (free — no API keys needed) ───────────────────
VERIFY_CONTENT      = True   # open every downloaded PDF and confirm it's a manual/datasheet
PREFER_MANUALS      = True   # keep searching candidates until a real MANUAL is found
ACCEPT_DATASHEETS   = True   # accept a datasheet when no manual can be located
ACCEPT_SCANNED      = True   # accept image-only scans (old manuals) when filename looks right
VERIFY_MAX_ATTEMPTS = 8      # hard cap on downloads per part
MANUAL_HUNT_EXTRA   = 2      # after a datasheet is accepted, try at most this many
                             # MORE candidates — and only ones whose NAME says 'manual'

NEXAR_CLIENT_ID     = ""   # nexar.com/api  (free 1k/month)
NEXAR_CLIENT_SECRET = ""

GOOGLE_CSE_KEY = ""        # console.cloud.google.com  (100 free/day)
GOOGLE_CSE_CX  = ""

SERPER_API_KEY = ""        # serper.dev  ($50/month 50k)
SERPAPI_KEY    = ""        # serpapi.com  ($50/month 5k)
BING_API_KEY   = ""        # RETIRED 2025-08-11 by Microsoft — endpoint returns HTTP 410, do not use
TAVILY_API_KEY = ""        # app.tavily.com  (1k free/month)
EXA_API_KEY    = ""        # dashboard.exa.ai  (1k free/month)
BRAVE_API_KEY  = ""        # brave.com/search/api  ($3/1k)

GDRIVE_API_KEY = ""        # console.cloud.google.com -> "Google Drive API" (free)
                           # OPTIONAL: makes Google Drive catalog listing more
                           # reliable for huge folders; without it the public
                           # folder view is scraped (works for normal catalogs)

MOUSER_API_KEY  = "d7fcc5ef-b4c4-4b57-b843-48b53921bd00"   # mouser.com/api-hub  (free — Search API; returns datasheet URLs)
FARNELL_API_KEY = ""       # partner.element14.com  (free)
FARNELL_STORE   = "uk.farnell.com"

DIGIKEY_CLIENT_ID     = ""  # developer.digikey.com  (free; OAuth2 client-credentials)
DIGIKEY_CLIENT_SECRET = ""  # best official datasheet coverage for electronic components

# ── Keyless / self-hosted search (NO signup) ─────────────────────────────────
# SearXNG is an open-source metasearch engine that aggregates Google/Bing/Brave/
# etc. and can return JSON. It needs no API key. Public instances are unreliable
# for JSON output (many disable it to deter abuse), so the robust path is to
# self-host one — it takes one command:
#     docker run -d --name searxng -p 8888:8080 searxng/searxng
# then add "http://localhost:8888" below. You must also enable the JSON format in
# the instance's settings.yml:  search.formats: [html, json]
# Multiple entries are tried in order; the first that returns results wins.
SEARX_INSTANCES = [
    # "http://localhost:8888",
    # "https://searx.be",
]

# =============================================================================
#  TERMINAL COLOURS
# =============================================================================

_COLOR = sys.stdout.isatty()
def _c(code, t): return f"\033[{code}m{t}\033[0m" if _COLOR else t
def _green(t):   return _c("92", t)
def _red(t):     return _c("91", t)
def _yellow(t):  return _c("93", t)
def _cyan(t):    return _c("96", t)
def _bold(t):    return _c("1",  t)
def _dim(t):     return _c("2",  t)

# =============================================================================
#  CIRCUIT BREAKER  -  only for optional/paid backends (never DDG / DirectProbe)
# =============================================================================

TRIP_THRESHOLD = 3

class CircuitBreaker:
    def __init__(self):
        self._lock    = threading.Lock()
        self._fails   = defaultdict(int)
        self._tripped = set()
        self._logged  = defaultdict(int)

    def is_open(self, name):
        with self._lock: return name in self._tripped

    def record_failure(self, name, err):
        with self._lock:
            self._fails[name] += 1
            n = self._logged[name]
            if n < 2:    tprint(f"    [{name}] {err}"); self._logged[name] += 1
            elif n == 2: tprint(f"    [{name}] (errors suppressed)"); self._logged[name] += 1
            if self._fails[name] >= TRIP_THRESHOLD and name not in self._tripped:
                self._tripped.add(name)
                tprint(f"    [{name}] disabled for this run.")

    def trip(self, name, reason=""):
        """Hard-disable a backend immediately (e.g. on a decisive bot-block 403).
        Prints exactly one message, even when called from many threads at once."""
        with self._lock:
            if name in self._tripped:
                return
            self._tripped.add(name)
            tprint(f"    [{name}] {reason or 'disabled for this run.'}")

    def record_success(self, name):
        with self._lock: self._fails[name] = 0; self._tripped.discard(name)

    def reset(self):
        """Clear every failure count and trip. Used before the end-of-run retry
        pass so backends disabled under parallel load (Mouser/RS bot-blocks etc.)
        get a genuinely fresh chance, mirroring a manual 'Search again'."""
        with self._lock:
            self._fails.clear()
            self._tripped.clear()
            self._logged.clear()

    def tripped_summary(self):
        with self._lock: return list(self._tripped)

_cb           = CircuitBreaker()
_print_lock   = threading.Lock()
_domain_locks = defaultdict(threading.Lock)
_domain_times = defaultdict(float)
MIN_DOMAIN_GAP = 1.5

# =============================================================================
#  PROGRESS / CANCELLATION HOOKS  (used by the GUI; no-ops in CLI mode)
# =============================================================================

_progress_cb : callable = None   # fn({"type": str, ...})
_cancel_ev   : threading.Event = None

def set_progress_callback(fn):
    global _progress_cb
    _progress_cb = fn

def set_cancel_event(ev):
    global _cancel_ev
    _cancel_ev = ev

def _emit(event_type: str, **kwargs):
    if _progress_cb:
        try: _progress_cb({"type": event_type, **kwargs})
        except Exception: pass

# ── Per-part skip (GUI: ✕ on a card while the run is live) ───────────────────
# _download_part records which part its worker thread is handling in a thread-
# local; _is_cancelled() then also reports True when the GUI has skipped THAT
# part. Every existing cancellation checkpoint in the search/download cascade
# thereby aborts just this one part without touching the others.
_skip_check : callable = None            # fn(idx) -> bool, installed by the GUI
_current_part_tl = threading.local()     # .idx of the part this thread works on

def set_skip_check(fn):
    global _skip_check
    _skip_check = fn

def _part_skipped() -> bool:
    if _skip_check is None:
        return False
    idx = getattr(_current_part_tl, "idx", None)
    if idx is None:
        return False
    try:
        return bool(_skip_check(idx))
    except Exception:
        return False

def _is_cancelled() -> bool:
    return (_cancel_ev is not None and _cancel_ev.is_set()) or _part_skipped()

# ── Search time budget (used by the GUI's interactive re-search) ─────────────
# When DDG rate-limits after a heavy run, the full search cascade can silently
# grind for many minutes. A deadline caps how long candidate-FINDING may take;
# ranking and downloading of already-found candidates are unaffected.
_search_deadline = [None]
_part_deadline_tl = threading.local()   # per-worker-thread part-search deadline
_ddg_part_calls_tl = threading.local()  # per-part count of REAL DDG attempts

DDG_MAX_PER_PART = 1   # hard cap on DDG calls per part. Each real call costs
                       # ~7s and is serialised globally — a part that issues 2
                       # (one in _rittal_portal for the product page, one in the
                       # trusted-branch manual top-up) was paying ~18s of pure
                       # DDG time EVEN AFTER it already had the official
                       # datasheet, and that volume is what trips the
                       # escalating 90s->600s rate-limit cooldowns. Whichever
                       # call site goes first per part gets it; the rest skip.

def _ddg_budget_left() -> bool:
    return getattr(_ddg_part_calls_tl, "n", 0) < DDG_MAX_PER_PART

def set_search_deadline(seconds):
    """Limit the search cascade to `seconds` from now (None = no limit)."""
    _search_deadline[0] = (time.monotonic() + seconds) if seconds else None

def _deadline_passed() -> bool:
    return _search_deadline[0] is not None and time.monotonic() > _search_deadline[0]

def _part_deadline_passed() -> bool:
    t = getattr(_part_deadline_tl, "t", None)
    return t is not None and time.monotonic() > t

def _should_stop_search() -> bool:
    return _is_cancelled() or _deadline_passed() or _part_deadline_passed()

def tprint(*a, **k):
    # Console output is cosmetic — it must NEVER kill a worker thread. In the
    # frozen windowed build (and any piped run) stdout is cp1252-encoded, so
    # characters like '→' raise UnicodeEncodeError; stdout can also be closed
    # or missing entirely. The GUI gets its copy via _emit below regardless.
    try:
        with _print_lock: print(*a, **k)
    except (UnicodeEncodeError, ValueError, OSError, AttributeError):
        pass
    _emit("log", message=" ".join(str(x) for x in a))

def _has(v):
    return bool(v and v.strip() and not v.strip().startswith("#"))

# =============================================================================
#  KNOWN MANUFACTURERS
# =============================================================================

_KNOWN_MFRS = sorted([
    ("ALLEN BRADLEY","ALLEN BRADLEY"),("PHOENIX CONTACT","PHOENIX CONTACT"),
    ("PHEONIX CONTACT","PHOENIX CONTACT"),("PHEONIX","PHOENIX CONTACT"),
    ("TRACO POWER","TRACO POWER"),("BARTON FIRTOP","BARTON FIRTOP"),
    ("GM INTERNATIONAL","GMI"),("EMERSON (ROSEMOUNT)","EMERSON"),
    ("APOLLO (ALJAC)","APOLLO"),("RITTAL","RITTAL"),("WEIDMULLER","WEIDMULLER"),
    ("SCHNEIDER","SCHNEIDER"),("HONEYWELL","HONEYWELL"),("PROSOFT","PROSOFT"),
    ("HIMA","HIMA"),("CISCO","CISCO"),("MOXA","MOXA"),("MTL","MTL"),
    ("EMERSON","EMERSON"),("DELTA","DELTA"),("AUTOCLAVE","AUTOCLAVE"),
    ("BUTECH","BUTECH"),("BIFOLD","BIFOLD"),("HYDAC","HYDAC"),
    ("LUBEDEVICE","LUBEDEVICE"),("EUROPRESS","EUROPRESS"),("DYNEX","DYNEX"),
    ("OLIVER","OLIVER"),("TESCOM","TESCOM"),("VERSA","VERSA"),
    ("MENNEKES","MENNEKES"),("BEKA","BEKA"),("CESP","CESP"),("CEAG","CEAG"),
    ("BIS","BIS"),("GMI","GMI"),("PARKER","PARKER"),("PROSERV","PROSERV"),
    ("SWAGELOK","SWAGELOK"),("APOLLO","APOLLO"),("ABB","ABB"),
], key=lambda x: len(x[0]), reverse=True)

_MFR_PAT = [
    (re.compile(r'(?<![A-Za-z])' + re.escape(n) + r'(?![A-Za-z])', re.I), c)
    for n, c in _KNOWN_MFRS
]

# Common BOM misspellings — fixed before ANY manufacturer-specific matching
# (direct probes, portal scrape, site: searches, own-domain ranking) so a typo
# in the source spreadsheet doesn't silently disable the whole official-docs
# pipeline for that part.
_MFR_SPELLFIX = [
    (re.compile(r'\bPHEONIX\b',    re.I), "PHOENIX"),
    (re.compile(r'\bPHONEIX\b',    re.I), "PHOENIX"),
    (re.compile(r'\bWIEDMULLER\b', re.I), "WEIDMULLER"),
    (re.compile(r'\bWEIDMULLAR\b', re.I), "WEIDMULLER"),
    (re.compile(r'\bSEIMENS\b',    re.I), "SIEMENS"),
    (re.compile(r'\bSYMENTAC\b',   re.I), "SYMANTEC"),
]

def _fix_mfr_spelling(mfr):
    if not mfr:
        return mfr
    for pat, repl in _MFR_SPELLFIX:
        mfr = pat.sub(repl, mfr)
    return mfr

# =============================================================================
#  BOM EXTRACTION  -  PDF / Excel / Markdown
# =============================================================================

def _is_cyr(t):
    cy = len(re.findall(r'[а-яА-ЯёЁ]', t))
    tot = len(re.findall(r'[a-zA-Zа-яА-ЯёЁ]', t))
    return tot > 5 and cy / tot > 0.4

def _looks_like_part(v):
    v = v.strip()
    return (bool(v) and 2 <= len(v) <= 80
            and not re.fullmatch(r'[\-_/\s\.]+', v)
            and bool(re.search(r'[A-Za-z0-9]', v)))

def _clean_model(m):
    p = m.split()
    if len(p) <= 1: return m
    last = p[-1].rstrip(".,")
    if re.fullmatch(r'[A-Z][A-Z0-9]{0,2}', last):
        rem = " ".join(p[:-1]).strip().rstrip("., ")
        if len(rem) >= 3 and re.search(r'[A-Za-z0-9]', rem): return rem
    return m

def _split_model_desc(text: str):
    """
    Given the text that follows a manufacturer name on a BOM line, split it into
    (part_number, description).

    In real BOMs the part number is always a single token (e.g. "8808000",
    "EN2T", "1756-A4K"). Everything after the first token is the description.

    Examples:
      "8808000 AX Enclosure 600x600x210"  → ("8808000",   "AX Enclosure 600x600x210")
      "EN2T EtherNet/IP Bridge Module"    → ("EN2T",      "EtherNet/IP Bridge Module")
      "1756-A4K ControlLogix 4-Slot"     → ("1756-A4K",  "ControlLogix 4-Slot")
      "8108245"                           → ("8108245",   "")
    """
    toks = text.strip().split()
    if not toks:
        return '', ''

    pn   = _clean_model(toks[0].rstrip('.,'))
    desc = ' '.join(toks[1:]).strip()

    # Only keep description when it contains real words (not just codes/numbers)
    if len(desc) >= 4 and re.search(r'[A-Za-z]{2,}', desc):
        return pn, desc
    return pn, ''

# Keep old name as a thin wrapper for any remaining call sites
def _model_after_mfr(text: str) -> str:
    return _split_model_desc(text)[0]

def _model_desc_around_mfr(text: str, m) -> tuple:
    """Part number + description for a line containing a known manufacturer,
    regardless of whether the BOM writes 'MFR PN desc' or 'PN MFR desc'.
    `m` is the regex match of the manufacturer inside `text`.
    Prefers the token AFTER the manufacturer (the historical layout); only
    when that token carries no digits while the token BEFORE the match is a
    strong part number does it flip to the 'PN first' reading."""
    after_model, after_desc = _split_model_desc(text[m.end():])
    before = text[:m.start()].strip().split()
    b = before[-1].rstrip('.,') if before else ''
    b_ok = (_looks_like_part(b) and bool(re.search(r'\d', b))
            and not re.fullmatch(r'\d{1,3}', b))       # a bare 1-3 digit token is a row counter
    a_ok = bool(after_model) and bool(re.search(r'\d', after_model))
    if b_ok and not a_ok:
        rest = text[m.end():].strip(' -–—:;,')
        desc = rest if (len(rest) >= 4 and re.search(r'[A-Za-z]{2,}', rest)) else ''
        return _clean_model(b), desc
    return after_model, after_desc

def _dedup(parts):
    """Filter out invalid part numbers but KEEP duplicates in BOM order —
    the UI flags them and lets the user keep or discard each one."""
    return [p for p in parts if _looks_like_part(p["part_number"])]

_HDR_MFR  = re.compile(r'(manufacturer|manuf\b|make\b|mfg|mfr|vendor|supplier|brand|maker|oem)', re.I)
_HDR_PART = re.compile(r'(part[\s_\-]?(no|num|code)|mpn|mfr[\s_\-]?pn|model[\s_\-]?(no|num)?\b|p/?n\b'
                       r'|article[\s_\-]?(no|num|code)?\b|order[\s_\-]?(no|num|code)'
                       r'|cat(alog(ue)?)?[\s_\-]?(no|num)|type[\s_\-]?(no|num|code))', re.I)
_HDR_SKIP = re.compile(r'(qty|quantity|ref|unit|price|tag|rev)', re.I)
_HDR_DESC = re.compile(r'(description|desc|item[\s_]?desc|material|component|service|function|notes?)', re.I)

def _find_cols(header):
    # A header like "Manufacturer Part Number" is a PART column, never the
    # manufacturer column — hence the "not _HDR_PART" guard on mc.
    mc = next((i for i,h in enumerate(header) if _HDR_MFR.search(str(h)) and not _HDR_PART.search(str(h)) and not _HDR_SKIP.search(str(h))), None)
    pc = [i for i,h in enumerate(header) if _HDR_PART.search(str(h)) and not _HDR_SKIP.search(str(h))]
    dc = next((i for i,h in enumerate(header) if _HDR_DESC.search(str(h)) and not _HDR_PART.search(str(h)) and not _HDR_MFR.search(str(h))), None)
    return mc, pc, dc

# ── Content-based column inference (headerless / unrecognised tables) ─────────
def _cell_is_code(v):
    """Looks like a part number: a compact alphanumeric code with digits —
    not a sentence, not a qty/line counter, not a dimension with units."""
    if not v or len(v) < 3 or len(v) > 40: return False
    if v.count(' ') > 1: return False
    if not re.search(r'\d', v): return False
    if re.fullmatch(r'\d{1,4}([.,]\d+)?', v): return False            # qty / line no / price
    if re.fullmatch(r'[\d.,]+\s*(mm|cm|m|kg|g|pcs?|packs?|ea|nos?|pc)\.?', v, re.I): return False
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_./+ #]*", v))

def _cell_is_known_mfr(v):
    if not v or len(v) > 40: return False
    return any(p.search(v) for p, _c in _MFR_PAT)

def _cell_is_wordy(v):
    return len(v) >= 12 and len(re.findall(r'[A-Za-z]{2,}', v)) >= 2

def _infer_cols(rows):
    """Work out which column holds what by scoring the CONTENT of every
    column — used when the header row is missing or its names aren't
    recognised, so any column order works.
    Returns (mfr_col, [pn_col], desc_col, first_data_row)."""
    cl = lambda v: re.sub(r'\s+', ' ', str(v or '')).strip()
    sample = [r for r in rows if any(cl(c) for c in r)][:80]
    if not sample: return None, [], None, 0
    ncols = max(len(r) for r in sample)

    # Skip the first row when it reads like a header (words only, no codes)
    first = [cl(c) for c in sample[0]]
    data_start = 0
    if first and not any(_cell_is_code(c) for c in first if c) \
       and any(re.fullmatch(r'[A-Za-z ()/#.%-]{2,30}', c) for c in first if c):
        data_start = 1
    data = sample[data_start:]
    if not data: return None, [], None, 0

    stats = []
    for col in range(ncols):
        filled = [cl(r[col]) for r in data if col < len(r) and cl(r[col])]
        n = len(filled)
        if not n:
            stats.append({"pn": 0, "mfr": 0, "desc": 0, "alpha": 0, "uniq": 1}); continue
        stats.append({
            "pn":    sum(_cell_is_code(c) for c in filled) / n,
            "mfr":   sum(_cell_is_known_mfr(c) and not _cell_is_code(c) for c in filled) / n,
            "desc":  sum(_cell_is_wordy(c) for c in filled) / n,
            # short alphabetic labels with lots of repetition → manufacturer-ish
            "alpha": sum(bool(re.fullmatch(r"[A-Za-z][A-Za-z&.\- ]{1,24}", c))
                         and len(c.split()) <= 3 for c in filled) / n,
            "uniq":  len(set(filled)) / n,
        })

    # Manufacturer: the column that keeps naming KNOWN manufacturers
    mc = max(range(ncols), key=lambda i: stats[i]["mfr"], default=None)
    if mc is None or stats[mc]["mfr"] < 0.3: mc = None

    # Part number: the strongest code-like column (excluding the mfr column)
    pn_ranked = sorted((i for i in range(ncols) if i != mc),
                       key=lambda i: stats[i]["pn"], reverse=True)
    pc = pn_ranked[0] if pn_ranked and stats[pn_ranked[0]]["pn"] >= 0.5 else None

    # Description: the wordiest remaining column
    d_ranked = sorted((i for i in range(ncols) if i != mc and i != pc),
                      key=lambda i: stats[i]["desc"], reverse=True)
    dc = d_ranked[0] if d_ranked and stats[d_ranked[0]]["desc"] >= 0.3 else None

    # No known manufacturer found: fall back to a short-word, low-variety column
    if mc is None:
        m_ranked = sorted((i for i in range(ncols)
                           if i != pc and i != dc
                           and stats[i]["alpha"] >= 0.6 and stats[i]["uniq"] <= 0.6),
                          key=lambda i: stats[i]["alpha"] * (1 - stats[i]["uniq"]),
                          reverse=True)
        mc = m_ranked[0] if m_ranked else None

    return mc, ([pc] if pc is not None else []), dc, data_start

def _rows_to_parts(rows):
    if not rows: return []
    cl = lambda v: re.sub(r'\s+', ' ', str(v or '')).strip()
    hdr = [cl(c) for c in rows[0]]; mc, pcs, dc = _find_cols(hdr)
    start = 1
    if not pcs:
        # Header row missing or unrecognised — infer the columns from the
        # data itself so the file works no matter how columns are ordered.
        mc, pcs, dc, start = _infer_cols(rows)
        if not pcs: return []
    parts = []
    for row in rows[start:]:
        cells = [cl(c) for c in row]
        for pc in pcs:
            pn   = cells[pc] if pc < len(cells) else ""
            mfr  = cells[mc] if mc is not None and mc < len(cells) else ""
            desc = cells[dc] if dc is not None and dc < len(cells) else ""
            if _looks_like_part(pn): parts.append({"manufacturer": mfr, "part_number": pn, "description": desc})
    return parts

def _desc_keywords(desc):
    """Extract meaningful search keywords from a BOM item description.
    Keeps short UPPERCASE family codes (TS, VX) and alphanumeric family
    tokens (TS8, VX25, IP55) — these are often the key to family-level docs.
    Drops dimension-style tokens (H2000, W1200, D800mm) and counting words."""
    if not desc: return []
    _STOP = {'the','a','an','and','or','for','to','of','in','with','by','at','from',
             'is','are','be','as','on','no','not','type','unit','module','assembly',
             'system','device','series','standard','general','purpose',
             'one','two','three','four','per','pcs','pack','incl','including'}
    out = []
    for w in re.findall(r'[A-Za-z][A-Za-z0-9]+', desc):
        wl = w.lower()
        if wl in _STOP or wl in out:
            continue
        if re.fullmatch(r'[A-Za-z]\d+[A-Za-z]*', w) and len(w) > 3:
            continue                                  # H2000 / W1200 / D800mm
        if (len(w) >= 3 and w.isalpha()) \
           or (w.isupper() and 2 <= len(w) <= 4) \
           or (any(ch.isdigit() for ch in w) and len(w) <= 4):
            out.append(wl)
        if len(out) >= 5:
            break
    return out

def _model_in_text(text, model):
    """True only if model appears as a WHOLE token — not as a prefix of a longer part number.
    e.g. 'EN2T' must NOT match 'EN2TSC'.
    Handles: "1756-EN2T-nd" (separators), "1756EN2T.pdf" (numeric prefix), "EN2T.pdf",
    and dotted part numbers both ways: '8108.235' matches '8108235' and vice versa.
    """
    if '|' in model:
        return any(_model_in_text(text, m) for m in model.split('|'))
    ml = model.lower()
    tl = text.lower()
    mn = re.sub(r'[-_\s.]+', '', ml)          # dot-stripped compact form
    # Check 1: original text with full word boundaries (handles hyphenated/spaced forms)
    if re.search(r'(?<![A-Za-z0-9])' + re.escape(ml) + r'(?![A-Za-z0-9])', tl):
        return True
    # Check 1b: dotless form of a dotted model as a whole token (8108.235 → 8108235)
    if '.' in ml and re.search(r'(?<![A-Za-z0-9])' + re.escape(mn) + r'(?![A-Za-z0-9])', tl):
        return True
    # Check 2: split the text on URL/path separators and match each segment.
    # Two passes: with '.' as a separator (handles "1756en2t.pdf") and without it
    # (keeps "8108.235" together so its dot-stripped form can match).
    tokens = set(re.split(r'[-_\s./+,;:()\[\]{}]', tl))
    tokens |= set(re.split(r'[-_\s/+,;:()\[\]{}]', tl))
    for token in tokens:
        tn = re.sub(r'[-_\s.]+', '', token)
        if tn == mn:
            return True
        # Allow wildcard matching for part numbers (e.g. MTLx544 matching MTL 5544)
        if 'x' in tn or '*' in tn:
            if len(tn) >= 5 and any(c.isdigit() for c in tn):
                pat = '^' + re.sub(r'(?<=\d)x|x(?=\d)', r'\\d', re.escape(tn)).replace(r'\*', r'.*') + '$'
                try:
                    if re.match(pat, mn):
                        return True
                except Exception:
                    pass
        if 'x' in mn or '*' in mn:
            if len(mn) >= 5 and any(c.isdigit() for c in mn):
                pat = '^' + re.sub(r'(?<=\d)x|x(?=\d)', r'\\d', re.escape(mn)).replace(r'\*', r'.*') + '$'
                try:
                    if re.match(pat, tn):
                        return True
                except Exception:
                    pass
        # Allow a leading digits-only prefix (e.g. "1756en2t" → model is "en2t")
        if tn.endswith(mn) and len(tn) > len(mn) and tn[:-len(mn)].isdigit():
            return True
    return False

# ── PDF ───────────────────────────────────────────────────────────────────────
def _extract_pdf(path):
    parts, seen = [], set()
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if _is_cyr(text[:500]): continue
            for raw in text.splitlines():
                line = re.sub(r'\s+', ' ', raw).strip()
                if not re.match(r'^\d+\s', line) or _is_cyr(line): continue
                for pat, can in _MFR_PAT:
                    m = pat.search(line)
                    if not m: continue
                    model, desc = _model_desc_around_mfr(line, m)
                    if (not model or len(model) < 2 or len(model) > 50
                            or not re.search(r'[A-Za-z0-9]', model)
                            or model.upper() in ("NA","N/A","-","CUSTOM","TBD")): continue
                    key = (can, model.upper())
                    if key not in seen: seen.add(key); parts.append({"manufacturer": can, "part_number": model, "description": desc})
                    break
    if not parts:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables(): parts.extend(_rows_to_parts(table))
    return _dedup(parts)

# ── Excel ─────────────────────────────────────────────────────────────────────
def _extract_excel(path):
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:
        print(f"  Could not open Excel file: {e}"); return []
    parts = []
    for sheet in wb.worksheets:
        rows = list(sheet.values)
        if not rows: continue
        for hi in range(min(5, len(rows))):
            _, pcs, _ = _find_cols([str(c or '') for c in rows[hi]])
            if pcs: parts.extend(_rows_to_parts(rows[hi:])); break
        else:
            # No recognisable header — infer columns from the cell CONTENT
            # (works with any column order, with or without a header row).
            inferred = _rows_to_parts(rows)
            if inferred:
                parts.extend(inferred)
                continue
            for row in rows:
                text = " ".join(str(c or '') for c in row)
                if _is_cyr(text): continue
                for pat, can in _MFR_PAT:
                    m = pat.search(text)
                    if not m: continue
                    model, desc = _model_desc_around_mfr(text, m)
                    if model and 2 <= len(model) <= 50:
                        parts.append({"manufacturer": can, "part_number": model, "description": desc})
                    break
    wb.close(); return _dedup(parts)

# ── Word (.docx) ──────────────────────────────────────────────────────────────
def _extract_docx(path):
    """BOM extraction from Word documents: tables first (any column order —
    header names when recognised, content inference otherwise), then a
    known-manufacturer scan over plain paragraphs."""
    try:
        import docx  # python-docx
    except ImportError:
        raise RuntimeError(
            "Reading Word files needs the 'python-docx' package — "
            "run:  pip install python-docx")
    doc = docx.Document(str(path))
    parts = []
    for tbl in doc.tables:
        rows = [[cell.text for cell in row.cells] for row in tbl.rows]
        parts.extend(_rows_to_parts(rows))
    if not parts:
        for para in doc.paragraphs:
            line = re.sub(r'\s+', ' ', para.text).strip()
            if not line or _is_cyr(line): continue
            for pat, can in _MFR_PAT:
                m = pat.search(line)
                if not m: continue
                model, desc = _model_desc_around_mfr(line, m)
                if model and 2 <= len(model) <= 50:
                    parts.append({"manufacturer": can, "part_number": model, "description": desc})
                break
    return _dedup(parts)

# ── Markdown ──────────────────────────────────────────────────────────────────
def _extract_markdown(path):
    text = Path(path).read_text(encoding="utf-8", errors="replace")
    parts = []; in_tbl = False; hdr = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("|") and s.endswith("|"):
            cells = [c.strip() for c in s.split("|")[1:-1]]
            if not in_tbl: in_tbl = True; hdr = cells; continue
            if re.fullmatch(r'[\-\:\|\s]+', s): continue
            if len(cells) == len(hdr):
                rd = dict(zip(hdr, cells))
                mfr  = next((rd[k] for k in rd if _HDR_MFR.search(k) and not _HDR_SKIP.search(k)), "")
                pn   = next((rd[k] for k in rd if _HDR_PART.search(k) and not _HDR_SKIP.search(k)), "")
                desc = next((rd[k] for k in rd if _HDR_DESC.search(k) and not _HDR_PART.search(k) and not _HDR_MFR.search(k)), "")
                if _looks_like_part(pn): parts.append({"manufacturer": mfr.strip(), "part_number": pn.strip(), "description": desc.strip()})
        else:
            in_tbl = False; hdr = []
    if not parts:
        for line in text.splitlines():
            if _is_cyr(line): continue
            for pat, can in _MFR_PAT:
                m = pat.search(line)
                if not m: continue
                model, desc = _model_desc_around_mfr(line, m)
                if model and 2 <= len(model) <= 50: parts.append({"manufacturer": can, "part_number": model, "description": desc})
                break
    return _dedup(parts)

# ── dispatcher ────────────────────────────────────────────────────────────────
def extract_parts_from_bom(bom_path):
    path = Path(bom_path); ext = path.suffix.lower()
    print(f"\n  Reading BOM: {path.name}")
    if ext == ".pdf":               return _extract_pdf(path)
    elif ext in (".xlsx", ".xls"):  return _extract_excel(path)
    elif ext == ".docx":            return _extract_docx(path)
    elif ext == ".doc":
        raise RuntimeError("Legacy .doc files aren't supported — save the "
                           "document as .docx (or PDF) in Word and try again.")
    elif ext in (".md", ".markdown", ".txt"): return _extract_markdown(path)
    else:
        print(f"  Unknown extension '{ext}', trying PDF then Excel …")
        parts = _extract_pdf(path)
        return parts if parts else _extract_excel(path)

# ── manual part entry ─────────────────────────────────────────────────────────
def _enter_parts_manually():
    print(f"""
  Enter one part per line:  {_cyan('MANUFACTURER')}  {_cyan('PART_NUMBER')}
  Use  {_dim('|')}  if manufacturer name has spaces:  {_cyan('ALLEN BRADLEY | 1756-A4K')}
  Type {_bold('done')} when finished.
""")
    parts = []
    while True:
        try: raw = input(f"  Part {_bold(str(len(parts)+1))} > ").strip()
        except (EOFError, KeyboardInterrupt): break
        if not raw: continue
        if raw.lower() == "done": break
        if "|" in raw:
            mfr, _, pn = raw.partition("|")
        else:
            toks = raw.split()
            if len(toks) < 2: print(f"  {_red('Need: MANUFACTURER PART_NUMBER')}"); continue
            pn = toks[-1]; mfr = " ".join(toks[:-1])
        mfr = mfr.strip(); pn = pn.strip()
        if not _looks_like_part(pn): print(f"  {_red(f'Invalid part number: {pn!r}')}"); continue
        parts.append({"manufacturer": mfr.upper(), "part_number": pn})
        print(f"  {_green('Added:')} {mfr.upper()}  {pn}")
    return parts

# =============================================================================
#  TIER 0 - DIRECT URL PROBING  (always active, no circuit breaker)
# =============================================================================

BROWSER_UA  = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
               "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
BROWSER_HDR = {"User-Agent": BROWSER_UA,
               "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
               "Accept-Language": "en-US,en;q=0.9"}

def _norm(pn):
    c = pn.strip()
    nd  = re.sub(r'[-_\s.]', '', c)      # compact: dashes, spaces AND dots removed (Rittal 8108.235 → 8108235)
    dot = re.sub(r'[-_\s]', '.', c)
    return c, nd, dot, c.lower(), nd.lower()

def _direct_candidates(model, mfr):
    model = model.split('|')[0]
    c, nd, dot, lo, lond = _norm(model); ml = (mfr or "").lower(); urls = []
    # Generic multi-brand CDNs
    urls += [(f"https://media.sminor.is/media/vorur_pdf/{lond}.pdf", f"{model} (sminor)"),
             (f"https://media.sminor.is/media/vorur_pdf/{lo}.pdf",   f"{model} (sminor)"),
             (f"https://media.sminor.is/media/vorur_pdf/{dot}.pdf",  f"{model} (sminor)")]
    if mfr:
        sl = ml.replace(" ","-"); fc = model[0].upper() if model else "0"
        urls += [(f"https://www.bolenscontrol.com/media/assets/product/documents/{sl}/{sl}-{lond}.pdf", f"{model} (bolens)"),
                 (f"https://www.bolenscontrol.com/media/assets/product/documents/{sl}/{sl}-{lo}.pdf",   f"{model} (bolens)"),
                 (f"https://rspsupply.com/images/downloads/{mfr.title()}/{fc}/{mfr.title()} {c}/Specification Sheet.pdf", f"{model} (rspsupply)"),
                 (f"https://shop.electech.com.eg/wp-content/uploads/2025/Brands/{mfr.title()}/datasheets/{lond}.pdf", f"{model} (electech)")]
    # Manufacturer-specific direct patterns
    if "rittal" in ml:
        urls += [
            (f"https://www.rittal.com/img/products/DATA/{nd}.pdf",                    f"{model} (Rittal)"),
            (f"https://www.rittal.com/img/products/DATA/{c}.pdf",                     f"{model} (Rittal)"),
            (f"https://static.rittal.com/dokumente/{nd}.pdf",                         f"{model} (Rittal)"),
        ]
    if "allen" in ml or "rockwell" in ml:
        # um=user manual, in=installation, rm=reference manual first — td (tech data) later
        for dt in ("um","in","rm","td","qr","sg","pp"):
            urls += [
                (f"https://literature.rockwellautomation.com/idc/groups/literature/documents/{dt}/{lond}-{dt}001_-en-p.pdf", f"{model} {dt.upper()} (Rockwell)"),
                (f"https://literature.rockwellautomation.com/idc/groups/literature/documents/{dt}/{lo}-{dt}001_-en-p.pdf",   f"{model} {dt.upper()} (Rockwell)"),
            ]
    if "weidm" in ml:
        urls += [
            (f"https://catalog.weidmueller.com/assets/{nd}.pdf",       f"{model} (Weidmuller)"),
            (f"https://catalog.weidmueller.com/assets/{c.upper()}.pdf", f"{model} (Weidmuller)"),
            (f"https://catalog.weidmueller.com/assets/{c}.pdf",         f"{model} (Weidmuller)"),
        ]
    if "phoenix" in ml:
        urls += [
            (f"https://www.phoenixcontact.com/assets/downloads_ed/global/web_dwl_technical_info/{nd}_en_xx.pdf",    f"{model} (Phoenix Contact)"),
            (f"https://www.phoenixcontact.com/assets/downloads_ed/global/web_dwl_installation/{nd}_en_xx.pdf",      f"{model} install (Phoenix Contact)"),
            (f"https://www.phoenixcontact.com/assets/downloads_ed/global/web_dwl_technical_info/{c}_en_xx.pdf",     f"{model} (Phoenix Contact)"),
            # DigiKey mirrors Phoenix Contact datasheets at a deterministic path
            # (e.g. …/Data%20Sheets/Phoenix%20Contact%20PDFs/2866789_DS.pdf).
            # NOTE: spaces in a URL *path* must be %20 — '+' only means space in
            # query strings, so '+' here 404s. Verified live: 2866789_DS.pdf,
            # 2320157.pdf both resolve with %20.
            (f"https://media.digikey.com/pdf/Data%20Sheets/Phoenix%20Contact%20PDFs/{nd}_DS.pdf",  f"{model} datasheet (Phoenix Contact via DigiKey)"),
            (f"https://media.digikey.com/pdf/Data%20Sheets/Phoenix%20Contact%20PDFs/{c}_DS.pdf",   f"{model} datasheet (Phoenix Contact via DigiKey)"),
            (f"https://media.digikey.com/pdf/Data%20Sheets/Phoenix%20Contact%20PDFs/{nd}.pdf",     f"{model} datasheet (Phoenix Contact via DigiKey)"),
            (f"https://media.digikey.com/pdf/Data%20Sheets/Phoenix%20Contact%20PDFs/{c}.pdf",      f"{model} datasheet (Phoenix Contact via DigiKey)"),
        ]
    if mfr and "phoenix" not in ml:
        # Same DigiKey media CDN convention works for many manufacturers.
        # Path spaces must be %20 (quote), never '+'.
        dk = quote(mfr.title())
        urls += [
            (f"https://media.digikey.com/pdf/Data%20Sheets/{dk}%20PDFs/{nd}_DS.pdf", f"{model} datasheet ({mfr} via DigiKey)"),
            (f"https://media.digikey.com/pdf/Data%20Sheets/{dk}%20PDFs/{nd}.pdf",    f"{model} datasheet ({mfr} via DigiKey)"),
        ]
    if "schneider" in ml:
        urls += [
            (f"https://download.schneider-electric.com/files?p_Doc_Ref={c}",          f"{model} (Schneider)"),
            (f"https://download.schneider-electric.com/files?p_Doc_Ref={nd}",         f"{model} (Schneider)"),
        ]
    if "mtl" in ml or "eaton" in ml:
        urls += [
            (f"https://www.mtl-inst.com/images/uploads/datasheets/{nd}.pdf",          f"{model} (MTL)"),
            (f"https://www.mtl-inst.com/images/uploads/datasheets/{c}.pdf",           f"{model} (MTL)"),
            (f"https://www.mtl-inst.com/images/uploads/manuals/{nd}.pdf",             f"{model} manual (MTL)"),
        ]
        # DirectProbe suffix/wildcard matching for MTL range products
        m_mtl = re.search(r'\b(?:MTL|EATON)?\s*(55|45|77|50)(\d{2})([A-Za-z0-9/]*)\b', model, re.I)
        if m_mtl:
            prefix = m_mtl.group(1)
            suffix = m_mtl.group(2)
            extra = m_mtl.group(3)
            extra_clean = re.sub(r'[^A-Za-z0-9]', '', extra)
            if prefix in ('45', '55'):
                urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}.pdf", f"{model} (MTL)"))
                urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}-S.pdf", f"{model} (MTL)"))
                if extra_clean:
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}{extra_clean}.pdf", f"{model} (MTL)"))
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}-C-{extra_clean}.pdf", f"{model} (MTL)"))
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}-{extra_clean}.pdf", f"{model} (MTL)"))
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/4500_5500/MTLx5{suffix}_{extra_clean}.pdf", f"{model} (MTL)"))
            elif prefix == '77':
                urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/7700/MTL77{suffix}.pdf", f"{model} (MTL)"))
                urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/7700/INM77{suffix}.pdf", f"{model} (MTL)"))
                if extra_clean:
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/7700/MTL77{suffix}{extra_clean}.pdf", f"{model} (MTL)"))
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/7700/INM77{suffix}{extra_clean}.pdf", f"{model} (MTL)"))
                    urls.append((f"https://www.mtl-inst.com/images/uploads/datasheets/7700/INM77{suffix}{extra_clean}+Rev+11.pdf", f"{model} (MTL)"))
    if "moxa" in ml:
        urls += [
            (f"https://www.moxa.com/getmedia/{lond}-datasheet.pdf",                  f"{model} (Moxa)"),
            (f"https://www.moxa.com/getmedia/{lo}-datasheet.pdf",                    f"{model} (Moxa)"),
        ]
    if "prosoft" in ml:
        urls += [
            (f"https://www.prosoft-technology.com/content/download/{lond}/{lond}.pdf", f"{model} (Prosoft)"),
        ]
    if "traco" in ml:
        urls += [
            (f"https://www.tracopower.com/models/{lo}.pdf",                           f"{model} (Traco)"),
            (f"https://www.tracopower.com/models/{lond}.pdf",                         f"{model} (Traco)"),
        ]
    if "hima" in ml:
        urls += [
            (f"https://www.hima.com/mediathek/download/{lond}.pdf",                  f"{model} (HIMA)"),
        ]
    if "parker" in ml:
        urls += [
            (f"https://ph.parker.com/inetcs/literature/{nd}/{nd}.pdf",               f"{model} (Parker)"),
        ]
    if "honeywell" in ml:
        urls += [
            (f"https://sensing.honeywell.com/honeywell-sensing-{lo}-product-sheet.pdf", f"{model} (Honeywell)"),
        ]
    return urls

def _probe_url_list(url_title_pairs, session):
    """HEAD-probe a list of (url, title) pairs in parallel; return live PDF hits.
    Some CDNs (Akamai/CloudFront configs) reject HEAD with 403/405 while serving
    GET normally — for those we fall back to a 1 KB ranged GET and sniff %PDF."""
    # Dedupe — for all-numeric parts the {nd} and {c} pattern expansions are the
    # SAME URL, which previously got probed (and counted) twice.
    seen, pairs = set(), []
    for u, t in url_title_pairs:
        if u not in seen:
            seen.add(u); pairs.append((u, t))
    url_title_pairs = pairs
    found = []
    # Only these hosts are known to reject HEAD while serving GET. For everything
    # else, a 403/405 means "no" — don't waste a 12s GET re-fetch on it.
    HEAD_BLOCKED_HOSTS = ("media.digikey.com", "literature.rockwellautomation.com")
    def _p(url, title):
        try:
            _throttle_probe(url)
            r = session.head(url, timeout=5, allow_redirects=True,
                             headers={"User-Agent": BROWSER_UA})
            if r.status_code == 200:
                ct = r.headers.get("Content-Type", "")
                if "pdf" in ct.lower() or _is_pdf_url(url):
                    return {"url": url, "title": title, "referer": None}
            elif (r.status_code in (403, 405, 501)
                  and any(h in url for h in HEAD_BLOCKED_HOSTS)):
                head = _get_first_bytes(url, session,
                                        {"User-Agent": BROWSER_UA, "Range": "bytes=0-1023"},
                                        n=512, total_budget=8)
                if head[:5] == b"%PDF-":
                    return {"url": url, "title": title, "referer": None}
        except Exception: pass
        return None
    # Inner pool is intentionally SMALL. This runs inside the outer per-part
    # download pool (up to DEFAULT_WORKERS threads), so a wide inner pool
    # multiplies out: 5 outer x 32 inner = 160 threads all sharing ONE requests
    # Session whose connection pool holds ~20 — the surplus threads block
    # forever waiting for a connection that never frees. Capping the inner pool
    # at 6 keeps worst-case concurrency (5 x 6 = 30) close to the pool size and
    # eliminates the deadlock. Probes are short, so 6-wide is still fast.
    workers = min(6, max(2, len(url_title_pairs)))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for fut in as_completed({ex.submit(_p, u, t): (u, t) for u, t in url_title_pairs}):
            r = fut.result()
            if r: found.append(r)
    return found

def make_pooled_session():
    """A requests.Session whose connection pool is sized for this app's peak
    concurrency. Outer download workers (≤ ~8) x inner probe pool (≤ 6) ≈ 48
    threads may want a connection at once; a default Session (pool_maxsize=10)
    starves under that and threads block forever inside .head()/.get() — the
    'stuck at checking manufacturer sources' freeze. Use this everywhere."""
    s = requests.Session()
    s.headers.update({"User-Agent": BROWSER_UA})
    s.max_redirects = 3
    try:
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=50, pool_maxsize=64, pool_block=False,
            max_retries=requests.adapters.Retry(total=1, backoff_factor=0.3),
        )
        s.mount("http://", adapter)
        s.mount("https://", adapter)
    except Exception:
        pass
    return s

def _direct_probe(model, mfr, session):
    """Always runs — no circuit breaker."""
    return _probe_url_list(_direct_candidates(model, mfr), session)

# =============================================================================
#  MANUFACTURER-SPECIFIC PORTAL SEARCH
#  Goes directly to each manufacturer's documentation portal — no search engine.
#  Runs after direct URL probing but before any DDG search.
# =============================================================================

_MFR_PORTAL_PATTERNS = {
    # key: substring(s) checked in UPPER manufacturer name
    # value: list of (product_page_url_template, search_url_template)
    # Placeholders: {model}, {nd} (no-dash), {lo} (lowercase), {lond} (lc no-dash)
    # RITTAL is handled by the dedicated _rittal_portal() branch — their old
    # int_en/product and search.jsp URLs are dead (site migrated to /com-en/).
    "ALLEN BRADLEY": [
        "https://www.rockwellautomation.com/en-us/products/details.{model}.html",
        "https://www.rockwellautomation.com/en-us/search.html#q={model}&t=All",
    ],
    "ROCKWELL": [
        "https://www.rockwellautomation.com/en-us/products/details.{model}.html",
    ],
    "PHOENIX CONTACT": [
        "https://www.phoenixcontact.com/en/products/{model}",
        "https://www.phoenixcontact.com/en/products/{nd}",
        "https://www.phoenixcontact.com/gb/products/{nd}",
    ],
    "SCHNEIDER": [
        "https://www.se.com/en/search/#q={model}&t=All",
        "https://www.schneider-electric.com/en/search/#q={model}&t=All",
    ],
    "WEIDM": [
        "https://catalog.weidmueller.com/catalog/Start.do?ObjectID={model}",
        "https://catalog.weidmueller.com/catalog/Start.do?ObjectID={nd}",
    ],
    "MOXA": [
        "https://www.moxa.com/en/products/?search={model}",
        "https://www.moxa.com/en/products/?search={nd}",
    ],
    "HIMA": [
        "https://www.hima.com/en/product/{lond}",
        "https://www.hima.com/en/search?q={model}",
    ],
    "TRACO": [
        "https://www.tracopower.com/products/{lond}/",
        "https://www.tracopower.com/search/?q={model}",
    ],
    "HONEYWELL": [
        "https://process.honeywell.com/us/en/products/{lond}",
        "https://sensing.honeywell.com/en/search?keyword={model}",
    ],
    "PROSOFT": [
        "https://www.prosoft-technology.com/search/?q={model}",
    ],
    "CISCO": [
        "https://www.cisco.com/c/en/us/search/index.html#q={model}&t=All",
    ],
    "PARKER": [
        "https://ph.parker.com/us/en/search#q={model}&t=All",
    ],
    "BIFOLD": [
        "https://www.bifold.com/search/?q={model}",
    ],
    "HYDAC": [
        "https://www.hydac.com/en/products.html?q={model}",
    ],
    "DELTA": [
        "https://www.deltaww.com/en-US/Search?q={model}",
    ],
    "MTL": [
        "https://www.mtl-inst.com/search/?q={model}",
    ],
}


def _rittal_portal(model, mfr, session, force=False):
    """Official Rittal documents straight from rittal.com.
    1. https://www.rittal.com/pdf-creator/variant/com-en/{nd}
       — deterministic endpoint that generates the OFFICIAL datasheet for any
       article number (8108.235 → 8108235). One GET, no search engine.
    2. The product page (found via one DDG lookup, since its path contains
       unguessable catalogue IDs) carries the OFFICIAL assembly-instructions
       links under rittal.com/imf/… — that is the manual."""
    c, nd, dot, lo, lond = _norm(model)
    found = []
    _t0 = time.monotonic()

    # (1) deterministic official datasheet. pdf-creator RENDERS on the fly and
    # drip-feeds bytes, so a plain timeout= can hang for minutes — use the
    # wall-clock-bounded sniff (≤18s) instead.
    ds_url = f"https://www.rittal.com/pdf-creator/variant/com-en/{nd}"
    try:
        _throttle(ds_url)
        head = _get_first_bytes(ds_url, session, BROWSER_HDR, n=512,
                                total_budget=20)
        if head[:4] == b"%PDF":
            found.append({"url": ds_url,
                          "title": f"{model} official Rittal datasheet (technical data)",
                          "referer": "https://www.rittal.com/com-en/products"})
    except Exception:
        pass
    if RITTAL_TIMING:
        tprint(f"    [rittal-timing] {model}: datasheet sniff "
               f"{time.monotonic()-_t0:.1f}s "
               f"({'hit' if found else 'miss'})")

    # (2) product page → official assembly instructions (the MANUAL).
    # If we already have the official datasheet, DON'T spend this part's ONE
    # DDG call here on "rittal.com com-en products {nd}" — live testing shows
    # this query frequently returns no rittal.com hit at all (Rittal's product
    # pages don't reliably rank for it). The trusted-branch manual-chase query
    # in _find_pdfs ('"{model}" Rittal manual') is broader and just as likely
    # to surface the manual, so we defer the budget to that single call. THIS
    # is what stops every Rittal part from making 2 serialised ~7s DDG calls.
    want_manual = PREFER_MANUALS and not any(
        "Instructions" in f["url"] for f in found)
    if found:
        suffix = "manual lookup deferred" if want_manual else "manual already present"
        tprint(f"    [MfrPortal:rittal.com] {len(found)} official doc(s) "
               f"for {model} ({suffix})")
        return found

    # NOTHING from the deterministic datasheet — this is the part's one shot
    # at finding anything via the portal, so spend the DDG budget here if any
    # remains and search engines aren't paused/exhausted (force= bypasses the
    # pause for interactive re-search, but never the budget/cancel checks).
    page_url = None
    ddg_paused = (not force) and time.monotonic() < _DDG_PAUSED_UNTIL[0]
    if (_ddg_budget_left() and not _is_cancelled() and not _should_stop_search()
            and not ddg_paused):
        _td = time.monotonic()
        try:
            for res in _ddg(f"rittal.com com-en products {nd}", n=6, force=force, critical=True):
                u = (res.get("href") or "").strip()
                if "rittal.com" in u and "/products/" in u and nd in u:
                    page_url = u
                    break
        except Exception:
            pass
        if RITTAL_TIMING:
            tprint(f"    [rittal-timing] {model}: DDG page lookup "
                   f"{time.monotonic()-_td:.1f}s "
                   f"({'found page' if page_url else 'no page'})")
    if page_url and not _is_cancelled():
        try:
            _throttle(page_url)
            html = _get_text_capped(page_url, session,
                                    {**BROWSER_HDR, "Referer": page_url})
            if html:
                for m in re.finditer(
                        r'https?://www\.rittal\.com/imf/[^\s"\'<>]*?'
                        r'(?:Instructions|Technical_details_EN)[^\s"\'<>]*',
                        html):
                    u = m.group(0)
                    kind = ("assembly instructions manual" if "Instructions" in u
                            else "technical details datasheet")
                    if not any(f["url"] == u for f in found):
                        found.append({"url": u,
                                      "title": f"{model} official Rittal {kind}",
                                      "referer": page_url})
        except Exception:
            pass

    if found:
        tprint(f"    [MfrPortal:rittal.com] {len(found)} official doc(s) for {model}")
    if RITTAL_TIMING:
        tprint(f"    [rittal-timing] {model}: TOTAL {time.monotonic()-_t0:.1f}s, "
               f"{len(found)} doc(s)")
    return found


def _mfr_portal_search(model, mfr, session, force=False):
    """
    Scrape manufacturer product pages directly — no search engine required.
    For each known manufacturer we try their product detail page URL first,
    then their on-site search, returning any PDF links found.
    """
    if not mfr or _is_cancelled():
        return []

    mu = mfr.upper()
    if "RITTAL" in mu:
        return _rittal_portal(model, mfr, session, force=force)
    # Find matching config entry
    pages = None
    for key, urls in _MFR_PORTAL_PATTERNS.items():
        if key in mu:
            pages = urls
            break
    if not pages:
        return []

    c, nd, dot, lo, lond = _norm(model)
    found = []

    def _expand(tmpl):
        return tmpl.format(model=model, nd=nd, lo=lo, lond=lond, c=c)

    for page_tmpl in pages:
        if _is_cancelled():
            break
        url = _expand(page_tmpl)
        hits = _scrape_page(url, model, mfr, session)
        if hits:
            tprint(f"    [MfrPortal:{_dom(url)}] {len(hits)} hit(s) for {model}")
            found.extend(hits)
            break   # found on first working page — don't need to try the rest

    return found



# =============================================================================
#  OPTIONAL BACKENDS  -  activate by filling in keys in CONFIG above
# =============================================================================

def _nexar_find(model, mfr, _tok=[None], _dead=[False]):
    if _dead[0] or not (_has(NEXAR_CLIENT_ID) and _has(NEXAR_CLIENT_SECRET)): return []
    if _cb.is_open("Nexar"): return []
    try:
        if not _tok[0]:
            r = requests.post("https://identity.nexar.com/connect/token",
                data={"grant_type":"client_credentials","client_id":NEXAR_CLIENT_ID,"client_secret":NEXAR_CLIENT_SECRET},timeout=15)
            r.raise_for_status(); _tok[0] = r.json()["access_token"]
        r = requests.post("https://api.nexar.com/graphql",
            json={"query":"query DS($q:String!){supSearchMpn(q:$q,limit:5){hits{part{mpn bestDatasheet{url name}documents{url name}}}}}","variables":{"q":f"{mfr} {model}".strip() if mfr else model}},
            headers={"Authorization":f"Bearer {_tok[0]}"},timeout=15)
        r.raise_for_status(); _cb.record_success("Nexar"); out=[]
        for hit in r.json().get("data",{}).get("supSearchMpn",{}).get("hits",[]):
            p=hit.get("part",{}); best=p.get("bestDatasheet")
            if best and best.get("url"): out.append({"url":best["url"],"title":best.get("name") or f"{model} datasheet","referer":None})
            for doc in p.get("documents",[]):
                if doc.get("url") and _is_pdf_url(doc["url"]) and doc["url"] not in {x["url"] for x in out}:
                    out.append({"url":doc["url"],"title":doc.get("name") or f"{model} doc","referer":None})
        return out
    except Exception as e:
        st=getattr(getattr(e,"response",None),"status_code",None)
        if st in (400,401,403): _dead[0]=True
        _cb.record_failure("Nexar",str(e)[:80]); return []

def _gcse(q,n=10):
    if not(_has(GOOGLE_CSE_KEY) and _has(GOOGLE_CSE_CX)) or _cb.is_open("Google CSE"): return []
    try:
        r=requests.get("https://www.googleapis.com/customsearch/v1",params={"key":GOOGLE_CSE_KEY,"cx":GOOGLE_CSE_CX,"q":q,"num":min(n,10)},timeout=12)
        r.raise_for_status(); _cb.record_success("Google CSE")
        return [{"href":i.get("link",""),"title":i.get("title","")} for i in r.json().get("items",[])]
    except Exception as e: _cb.record_failure("Google CSE",str(e)[:80]); return []

def _serper(q,n=10):
    if not _has(SERPER_API_KEY) or _cb.is_open("Serper"): return []
    try:
        r=requests.post("https://google.serper.dev/search",headers={"X-API-KEY":SERPER_API_KEY,"Content-Type":"application/json"},json={"q":q,"num":n},timeout=12)
        r.raise_for_status(); _cb.record_success("Serper")
        return [{"href":i.get("link",""),"title":i.get("title","")} for i in r.json().get("organic",[])]
    except Exception as e: _cb.record_failure("Serper",str(e)[:80]); return []

def _serpapi(q,n=10):
    if not _has(SERPAPI_KEY) or _cb.is_open("SerpApi"): return []
    try:
        r=requests.get("https://serpapi.com/search",params={"api_key":SERPAPI_KEY,"q":q,"engine":"google","num":n},timeout=12)
        r.raise_for_status(); _cb.record_success("SerpApi")
        return [{"href":i.get("link",""),"title":i.get("title","")} for i in r.json().get("organic_results",[])]
    except Exception as e: _cb.record_failure("SerpApi",str(e)[:80]); return []

def _bing(q,n=10):
    if not _has(BING_API_KEY) or _cb.is_open("Bing"): return []
    try:
        r=requests.get("https://api.bing.microsoft.com/v7.0/search",params={"q":q,"count":n},headers={"Ocp-Apim-Subscription-Key":BING_API_KEY},timeout=12)
        r.raise_for_status(); _cb.record_success("Bing")
        return [{"href":i.get("url",""),"title":i.get("name","")} for i in r.json().get("webPages",{}).get("value",[])]
    except Exception as e: _cb.record_failure("Bing",str(e)[:80]); return []

def _tavily(q,n=10):
    if not _has(TAVILY_API_KEY) or _cb.is_open("Tavily"): return []
    try:
        r=requests.post("https://api.tavily.com/search",json={"api_key":TAVILY_API_KEY,"query":q,"max_results":n,"search_depth":"basic"},timeout=15)
        r.raise_for_status(); _cb.record_success("Tavily")
        return [{"href":i.get("url",""),"title":i.get("title","")} for i in r.json().get("results",[])]
    except Exception as e: _cb.record_failure("Tavily",str(e)[:80]); return []

def _exa(q,n=10):
    if not _has(EXA_API_KEY) or _cb.is_open("Exa"): return []
    try:
        r=requests.post("https://api.exa.ai/search",json={"query":q,"numResults":n,"type":"keyword","contents":{"text":False}},headers={"x-api-key":EXA_API_KEY,"Content-Type":"application/json"},timeout=15)
        r.raise_for_status(); _cb.record_success("Exa")
        return [{"href":i.get("url",""),"title":i.get("title","")} for i in r.json().get("results",[])]
    except Exception as e: _cb.record_failure("Exa",str(e)[:80]); return []

def _brave(q,n=10):
    if not _has(BRAVE_API_KEY) or _cb.is_open("Brave"): return []
    try:
        r=requests.get("https://api.search.brave.com/res/v1/web/search",params={"q":q,"count":n,"result_filter":"web"},
            headers={"Accept":"application/json","Accept-Encoding":"gzip","X-Subscription-Token":BRAVE_API_KEY},timeout=12)
        r.raise_for_status(); _cb.record_success("Brave")
        return [{"href":i.get("url",""),"title":i.get("title","")} for i in r.json().get("web",{}).get("results",[])]
    except Exception as e: _cb.record_failure("Brave",str(e)[:80]); return []

def _searx(q, n=10):
    """SearXNG metasearch — KEYLESS, self-hostable. Aggregates many engines and
    returns JSON, so it's an excellent unlimited substitute for DDG that doesn't
    suffer DuckDuckGo's per-IP rate limiting. Tries each configured instance in
    turn; a flaky instance is NOT circuit-broken on its own (we just move to the
    next, then fall through to DDG)."""
    for base in SEARX_INSTANCES:
        if not _has(base):
            continue
        base = base.rstrip("/")
        try:
            r = requests.get(
                f"{base}/search",
                params={"q": q, "format": "json", "language": "en",
                        "safesearch": 0, "categories": "general"},
                headers={"User-Agent": BROWSER_UA, "Accept": "application/json"},
                timeout=12)
            if r.status_code != 200:
                continue
            hits = [{"href": it.get("url", ""), "title": it.get("title", "")}
                    for it in (r.json().get("results") or []) if it.get("url")]
            if hits:
                _cb.record_success("SearXNG")
                return hits[:n]
        except Exception as e:
            _cb.record_failure("SearXNG", f"{_dom(base)}: {str(e)[:60]}")
    return []

def _mouser_api(model, mfr):
    """Mouser Search API (free key). Returns DIRECT official datasheet URLs —
    far better than guessing via a web search. Shape matches _nexar_find:
    [{url,title,referer}]. Off unless MOUSER_API_KEY is set."""
    if not _has(MOUSER_API_KEY) or _cb.is_open("Mouser API"): return []
    kw = f"{mfr} {model}".strip() if mfr else model
    try:
        r = requests.post(
            "https://api.mouser.com/api/v1/search/keyword",
            params={"apiKey": MOUSER_API_KEY},
            json={"SearchByKeywordRequest": {"keyword": kw, "records": 10,
                  "searchOptions": "", "startingRecord": 0}},
            timeout=15)
        r.raise_for_status(); _cb.record_success("Mouser API")
        out, seen = [], set()
        for p in ((r.json().get("SearchResults") or {}).get("Parts") or []):
            ds = (p.get("DataSheetUrl") or "").strip()
            if ds.lower().startswith("http") and ds not in seen:
                seen.add(ds)
                mpn = p.get("ManufacturerPartNumber") or model
                out.append({"url": ds, "title": f"{mpn} datasheet", "referer": None})
        return out
    except Exception as e:
        _cb.record_failure("Mouser API", str(e)[:80]); return []

def _digikey_find(model, mfr, _tok=[None, 0.0]):
    """DigiKey Product Information API (free dev account, OAuth2 client-creds).
    Best official-datasheet coverage for electronic components. Returns
    [{url,title,referer}]. Off unless DIGIKEY_CLIENT_ID/SECRET are set.
    NOTE: DigiKey's API is finicky (OAuth + exact v4 endpoint + locale headers).
    If it returns HTTP 400/401, verify your app's configured endpoint and locale
    in the DigiKey developer dashboard."""
    if not (_has(DIGIKEY_CLIENT_ID) and _has(DIGIKEY_CLIENT_SECRET)): return []
    if _cb.is_open("DigiKey"): return []
    kw = f"{mfr} {model}".strip() if mfr else model
    try:
        if not _tok[0] or time.monotonic() > _tok[1]:
            tr = requests.post("https://api.digikey.com/v1/oauth2/token",
                data={"grant_type": "client_credentials",
                      "client_id": DIGIKEY_CLIENT_ID,
                      "client_secret": DIGIKEY_CLIENT_SECRET}, timeout=15)
            tr.raise_for_status(); j = tr.json()
            _tok[0] = j["access_token"]
            _tok[1] = time.monotonic() + int(j.get("expires_in", 600)) - 30
        r = requests.post(
            "https://api.digikey.com/products/v4/search/keyword",
            headers={"Authorization": f"Bearer {_tok[0]}",
                     "X-DIGIKEY-Client-Id": DIGIKEY_CLIENT_ID,
                     "X-DIGIKEY-Locale-Site": "US",
                     "X-DIGIKEY-Locale-Language": "en",
                     "X-DIGIKEY-Locale-Currency": "USD",
                     "Content-Type": "application/json"},
            json={"Keywords": kw, "Limit": 10, "Offset": 0}, timeout=15)
        r.raise_for_status(); _cb.record_success("DigiKey")
        out, seen = [], set()
        for p in (r.json().get("Products") or []):
            ds = (p.get("DatasheetUrl") or "").strip()
            if ds.startswith("//"): ds = "https:" + ds
            if ds.lower().startswith("http") and ds not in seen:
                seen.add(ds)
                mpn = p.get("ManufacturerProductNumber") or model
                out.append({"url": ds, "title": f"{mpn} datasheet", "referer": None})
        return out
    except Exception as e:
        st = getattr(getattr(e, "response", None), "status_code", None)
        if st in (401, 403): _tok[0] = None
        _cb.record_failure("DigiKey", str(e)[:80]); return []

def _component_db_docs(model, mfr):
    """Direct datasheet/manual URLs from component databases (Nexar/Octopart,
    Mouser, DigiKey). Each is independently key-gated and circuit-broken; any
    that isn't configured returns []. De-duplicated official-doc URLs, shaped
    like _direct_probe hits ([{url,title,referer}]) and fed straight into the
    candidate pool. Replaces the previous lone _nexar_find() call sites."""
    out, seen = [], set()
    for fn in (_nexar_find, _mouser_api, _digikey_find):
        try:
            hits = fn(model, mfr)
        except Exception:
            hits = []
        for h in (hits or []):
            u = h.get("url")
            if u and u not in seen:
                seen.add(u); out.append(h)
    return out

# -----------------------------------------------------------------------------
# Active search backends.  _search() tries them IN ORDER and returns the first
# non-empty result, so cheap/unlimited engines go first to spare the metered
# ones; DDG is always the final fallback and is NEVER circuit-broken.
#   - SearXNG : keyless, unlimited (self-host) — enabled when SEARX_INSTANCES set
#   - Google CSE / Tavily : genuinely free tiers (100/day, 1k/month) — set key to enable
#   - Exa / Serper / SerpApi : free credit / trial — set key to enable
#   - Bing : RETIRED 2025-08-11 (HTTP 410), removed
#   - Brave : free tier removed 2026-02, now needs a card — left disabled
# Component databases (Nexar/Mouser/DigiKey) are NOT here; they return direct
# document URLs and are queried in _find_pdfs via _component_db_docs().
# -----------------------------------------------------------------------------
_BACKENDS = [
    ("SearXNG",    lambda: bool(SEARX_INSTANCES),                    _searx),
    ("Google CSE", lambda: _has(GOOGLE_CSE_KEY) and _has(GOOGLE_CSE_CX), _gcse),
    ("Tavily",     lambda: _has(TAVILY_API_KEY),                     _tavily),
    ("Exa",        lambda: _has(EXA_API_KEY),                        _exa),
    ("Serper",     lambda: _has(SERPER_API_KEY),                     _serper),
    ("SerpApi",    lambda: _has(SERPAPI_KEY),                        _serpapi),
    # ("Brave",    lambda: _has(BRAVE_API_KEY),  _brave),  # free tier removed 2026-02 (card required)
    ("DDG", lambda: True, None),   # handler assigned below — ALWAYS the last fallback
]

# DDG rate-limit circuit-breaker: after 2 queries fail completely, pause DDG so
# the cascade flies through to direct probes / portal scrapers instead of
# burning the whole search budget on sleeps. Newer `ddgs` versions also accept
# backend="auto" which rotates engines (Bing/Brave/Google…) — that alone often
# dodges DuckDuckGo-specific rate limits.
_DDG_FAILS    = [0]
_DDG_PAUSED_UNTIL = [0.0]
_DDG_PAUSE_COUNT  = [0]          # escalates the cooldown: 90s -> 180s -> 360s -> 600s cap
DDG_COOLDOWN_SECS = 90
DDG_COOLDOWN_MAX  = 600
DDG_CALL_TIMEOUT  = 12           # hard cap (s) per HTTP call INSIDE the ddgs lib.
                                 # Measured real-world latency for backend="auto"
                                 # is ~7.3s — 8s was clipping legitimate slow
                                 # calls, counting them as FAILS and accelerating
                                 # the 90s->600s rate-limit cooldowns. 12s gives
                                 # headroom while still well under the 30-60s
                                 # hang this was originally added to prevent.

# All DDG queries from EVERY worker thread are serialised through this lock and
# spaced DDG_MIN_INTERVAL seconds apart. With N parallel workers each firing
# multiple queries per part, DDG sees a burst of simultaneous requests from one
# IP and rate-limits almost immediately -- pacing globally prevents that.
_DDG_QUERY_LOCK  = threading.Lock()
_DDG_LAST_QUERY  = [0.0]
DDG_MIN_INTERVAL = 1.5    # global spacing between DDG calls. Lowered from 3.0
                          # because DDG volume is now a fraction of what it was:
                          # DirectProbe/portal short-circuits mean only parts
                          # with NO official mirror ever reach DDG, so the burst
                          # risk that needed 3s spacing is gone.
DDG_LOCK_WAIT_MAX = 8.0   # max seconds any single query waits in line for the
                          # global lock before giving up (was 20s). A part that
                          # already holds an official doc should never block for
                          # long chasing an optional manual.

_DDG_CACHE = {}
_DDG_CACHE_LOCK = threading.Lock()

def _any_search_backend_ready() -> bool:
    """True if any non-DDG search backend (Google CSE, Bing, Serper, …) is
    configured and not circuit-broken. When one exists, the search cascade can
    still work even while DDG is paused, so we must NOT fail-fast."""
    for name, enabled, _fn in _BACKENDS:
        if name == "DDG":
            continue
        try:
            if enabled() and not _cb.is_open(name):
                return True
        except Exception:
            pass
    return False


def _ddg(q, n=10, retries=2, force=False, critical=False):
    """DuckDuckGo - always active, never circuit-broken (but it does cool down
    when rate-limited, with an escalating pause). Calls are serialised and
    paced across threads to stay under DDG's per-IP rate limit."""
    if not force:
        with _DDG_CACHE_LOCK:
            if q in _DDG_CACHE:
                return _DDG_CACHE[q]

    if not force and time.monotonic() < _DDG_PAUSED_UNTIL[0]:
        return []
    if not force:
        max_calls = 3 if critical else DDG_MAX_PER_PART
        if getattr(_ddg_part_calls_tl, "n", 0) >= max_calls:
            return []
    for attempt in range(1, retries + 1):
        if _should_stop_search(): return []
        # Bound how long this part will WAIT IN LINE for the global lock: its
        # remaining budget, capped at 20s. Without this, with N workers
        # funnelling queries through one lock, a part could queue for minutes —
        # invisible to the budget checks between cascade tiers.
        limit = 45.0 if critical else DDG_LOCK_WAIT_MAX
        max_wait = limit
        t_start_wait = time.monotonic()
        acquired = _DDG_QUERY_LOCK.acquire(timeout=max_wait)
        wait_duration = time.monotonic() - t_start_wait
        if wait_duration > 0.05:
            t_dl = getattr(_part_deadline_tl, "t", None)
            if t_dl is not None:
                _part_deadline_tl.t = t_dl + wait_duration
        if not acquired:
            return []
        try:
            # Re-check the cache and pause now that we hold the lock -- another thread may
            # have populated the cache or tripped the rate-limit while we were waiting in line.
            if not force:
                with _DDG_CACHE_LOCK:
                    if q in _DDG_CACHE:
                        return _DDG_CACHE[q]
            if not force and time.monotonic() < _DDG_PAUSED_UNTIL[0]:
                return []
            if _should_stop_search(): return []
            gap = DDG_MIN_INTERVAL - (time.monotonic() - _DDG_LAST_QUERY[0])
            if gap > 0:
                time.sleep(gap)
            # This is the real, costly attempt (~7s observed) — count it toward
            # this part's DDG_MAX_PER_PART cap regardless of outcome.
            _ddg_part_calls_tl.n = getattr(_ddg_part_calls_tl, "n", 0) + 1
            err = None
            backend = "api" if attempt == 1 else "html"
            try:
                try:
                    dd = DDGS(timeout=DDG_CALL_TIMEOUT)
                except TypeError:               # very old lib without `timeout`
                    dd = DDGS()
                with dd as d:
                    try:
                        r = list(d.text(q, max_results=n, backend=backend))
                    except TypeError:           # older lib without `backend` param
                        r = list(d.text(q, max_results=n))
            except Exception as e:
                r, err = [], e
            _DDG_LAST_QUERY[0] = time.monotonic()
        finally:
            _DDG_QUERY_LOCK.release()
        if r:
            _DDG_FAILS[0] = 0
            _DDG_PAUSE_COUNT[0] = 0
            _DDG_PAUSED_UNTIL[0] = 0.0
            if not force:
                with _DDG_CACHE_LOCK:
                    _DDG_CACHE[q] = r
            return r
        if _should_stop_search(): return []
        if err:
            err_str = str(err).lower()
            if any(x in err_str for x in ("dns error", "query refused", "name resolution", "non-recoverable failure", "connection refused")):
                _DDG_FAILS[0] = 2
                _DDG_PAUSE_COUNT[0] = max(_DDG_PAUSE_COUNT[0], 2)
                break
        time.sleep(attempt * (3 if err else 2))
    if err:
        _DDG_FAILS[0] += 1
    if _DDG_FAILS[0] >= 2:
        cooldown = min(DDG_COOLDOWN_SECS * (2 ** _DDG_PAUSE_COUNT[0]), DDG_COOLDOWN_MAX)
        _DDG_PAUSE_COUNT[0] += 1
        _DDG_PAUSED_UNTIL[0] = time.monotonic() + cooldown
        _DDG_FAILS[0] = 0
        tprint(f"    [DDG] {_yellow('search engine rate-limited')} -- pausing it for "
               f"{int(cooldown)}s (direct manufacturer sources still active)")
    return []


_BACKENDS[-1] = ("DDG", lambda: True, _ddg)   # wire in now that fn is defined

def _search(q, n=10, force=False, critical=False):
    for name, enabled, fn in _BACKENDS:
        if _should_stop_search(): return []
        if not enabled(): continue
        if name != "DDG" and _cb.is_open(name): continue   # DDG never blocked
        r = fn(q, n, force=force, critical=critical) if name == "DDG" else fn(q, n)
        if r: return r
    return []


def reset_search_throttle():
    """Clear ALL self-imposed search-engine cooldowns and circuit-breaker trips
    so the next queries start fresh.

    The parallel main pass routinely rate-limits DuckDuckGo into an escalating
    global pause (_DDG_PAUSED_UNTIL) and trips the Mouser/RS scrapers for the
    whole run. That silently starves parts which depend purely on search engines
    (e.g. ABB — no DirectProbe/portal mirror): they come back 'not found' even
    though the document is reachable. Calling this before the sequential retry
    pass makes that pass behave like a manual 'Search again'."""
    _DDG_PAUSED_UNTIL[0] = 0.0
    _DDG_PAUSE_COUNT[0]  = 0
    _DDG_FAILS[0]        = 0
    try:
        _MOUSER_EMPTY[0] = 0          # defined later in the module; safe at call time
    except Exception:
        pass
    _cb.reset()
    with _DDG_CACHE_LOCK:
        _DDG_CACHE.clear()

# =============================================================================
#  DOWNLOAD UTILITIES
# =============================================================================

MFR_DOMS  = ["rockwellautomation.com","rittal.com","schneider-electric.com",
             "phoenixcontact.com","weidmuller.com","hima.com","cisco.com",
             "moxa.com","honeywell.com","mtl-inst.com","prosoft-technology.com"]
DIST_DOMS = ["farnell.com","mouser.com","digikey.com","rs-online.com",
             "element14.com","avnet.com","arrow.com","tme.eu",
             "sminor.is","bolenscontrol.com","rspsupply.com"]

# Manufacturer documentation portals for site-specific DDG searches.
# DDG honours 'site:' even though it ignores 'filetype:' — this is the
# single most effective free improvement for finding exact manuals.
_MFR_DOC_SITES = {
    "ALLEN BRADLEY":   "literature.rockwellautomation.com",
    "ROCKWELL":        "literature.rockwellautomation.com",
    "PHOENIX CONTACT": "phoenixcontact.com",
    "SCHNEIDER":       "se.com",
    "RITTAL":          "rittal.com",
    "WEIDMULLER":      "weidmueller.com",
    "WEIDM":           "weidmueller.com",
    "HIMA":            "hima.com",
    "MOXA":            "moxa.com",
    "MTL":             "mtl-inst.com",
    "CISCO":           "cisco.com",
    "HONEYWELL":       "process.honeywell.com",
    "PROSOFT":         "prosoft-technology.com",
    "TRACO":           "tracopower.com",
    "DELTA":           "deltaww.com",
    "PARKER":          "parker.com",
    "BIFOLD":          "bifold.com",
    "HYDAC":           "hydac.com",
    "MENNEKES":        "mennekes.de",
    "ABB":             "abb.com",
}

def _mfr_doc_site(mfr):
    """Return the manufacturer's primary documentation domain, or None."""
    if not mfr: return None
    mu = mfr.upper()
    for key, site in _MFR_DOC_SITES.items():
        if key in mu:
            return site
    return None

def _dom(url): return re.sub(r'^https?://(www\.)?','',url.lower()).split('/')[0]

def _base_dom(d):
    """'literature.rockwellautomation.com' → 'rockwellautomation.com'"""
    p = d.split('.')
    return '.'.join(p[-2:]) if len(p) >= 2 else d

# Manufacturer names whose own website doesn't contain their name
_MFR_DOM_ALIASES = {
    "ALLEN":     ["rockwellautomation.com"],
    "ROCKWELL":  ["rockwellautomation.com"],
    "SCHNEIDER": ["se.com", "schneider-electric.com"],
    "MTL":       ["mtl-inst.com", "eaton.com"],
    "EATON":     ["mtl-inst.com", "eaton.com"],
    "PHOENIX":   ["phoenixcontact.com"],
    "WEIDM":     ["weidmueller.com"],
    "ROSEMOUNT": ["emerson.com"],
    "TRACO":     ["tracopower.com"],
    "PROSOFT":   ["prosoft-technology.com"],
    "GMI":       ["gmintsrl.com"],
    "ABB":       ["abb.com", "library.abb.com", "library.e.abb.com"],
}

# ── ABB dual-identity resolution ────────────────────────────────────────────
# Every ABB part carries TWO identifiers: an order/article number (the BOM
# value, e.g. 3BSE052605R1) and a product/type designation (e.g. AO815). Real
# datasheets/manuals are titled and indexed by the DESIGNATION, and ABB's own
# library PDFs are picture-heavy and frequently never spell out the order number
# as extractable text — so matching on the order number alone (a) can't find the
# right document and (b) wrongly rejects it after download. We resolve the
# designation from result titles / distributor listings, cache it per order
# number, and then treat it as an equally-valid identity everywhere.
_ABB_ALIAS = {}                       # normalised order number -> designation
_ABB_ALIAS_LOCK = threading.Lock()

# Tokens that match the designation pattern but are generic FAMILY / umbrella
# names — never accept these as a specific part alias.
_ABB_DESIG_STOP = {
    "AC800M", "AC800", "AC700", "800XA", "S800", "S900", "S800IO", "S900IO",
    "RDIO", "RAIO", "RTAC", "IRB", "DSQC",
    "ACS380", "ACS580", "ACS880", "ACS800", "ACS550",
}

# Title/URL markers of generic ABB documents (not a single-part datasheet).
_ABB_FAMILY_RX = re.compile(
    r'integrity|global|sample\s+specification|price\s*list|selector|brochure|'
    r'catalog|catalogue|robotics|technical\s+reference|overview|portfolio|'
    r'system\s+guide|product\s+guide', re.I)


def _extract_abb_designation(results, model):
    """Find the ABB product/type designation (e.g. 'AO815', 'CI867K01') that
    belongs to an order number (e.g. '3BSE052605R1').
    `results` is any iterable of dicts carrying 'title' and/or 'url'/'href'.
    Strongest evidence: a designation that appears in the SAME listing as the
    order number. Failing that: a designation seen across several listings."""
    if not results:
        return None
    from collections import Counter
    pat = re.compile(r'\b([A-Z]{2,4}\d{3}[A-Z0-9]*)\b')
    model_clean = re.sub(r'[-_\s.]+', '', model).upper()
    counter = Counter()
    for r in results:
        blob = (f"{r.get('title','') or ''} "
                f"{r.get('snippet','') or r.get('body','') or ''} "
                f"{r.get('url','') or r.get('href','') or ''}").upper()
        blob_clean = re.sub(r'[-_\s.]+', '', blob)
        found = [f for f in pat.findall(blob)
                 if f != model_clean
                 and f not in _ABB_DESIG_STOP
                 and re.sub(r'[-_\s.]+', '', f) != model_clean]
        if not found:
            continue
        if model_clean in blob_clean:        # order # and designation together
            return found[0]
        counter.update(found)
    if counter:
        cand, n = counter.most_common(1)[0]
        if n >= 2:
            return cand
    return None


def _is_abb(mfr) -> bool:
    return bool(mfr) and "ABB" in mfr.upper()


def _norm_abb(pn) -> str:
    return re.sub(r'[-_\s.]+', '', (pn or '')).upper()


def _abb_alias_for(model, mfr):
    """Cached designation for this order number, or None."""
    if not _is_abb(mfr):
        return None
    return _ABB_ALIAS.get(_norm_abb(model))


def _abb_register_alias(model, mfr, results):
    """Resolve and cache the ABB type designation for `model`. Idempotent and
    cheap — safe to call repeatedly as candidates accumulate. Accepts either a
    {url: {title}} dict (the `direct` pool) or a list of raw search results."""
    if not _is_abb(mfr):
        return None
    key = _norm_abb(model)
    if key in _ABB_ALIAS:
        return _ABB_ALIAS[key]
    if isinstance(results, dict):
        items = [{"title": v.get("title", ""), "url": u} for u, v in results.items()]
    else:
        items = [{"title": r.get("title") or "",
                  "url": r.get("url") or r.get("href") or ""}
                 for r in (results or [])]
    desig = _extract_abb_designation(items, model)
    if desig:
        with _ABB_ALIAS_LOCK:
            _ABB_ALIAS.setdefault(key, desig.upper())
        tprint(f"    [ABB] {model} \u2194 type designation {desig.upper()} "
               f"\u2014 treating both as the same part")
    return _ABB_ALIAS.get(key)


def _match_model(model, mfr):
    """Identifier used for part-number MATCHING only (never for filenames or
    queries). For ABB this folds the resolved type designation into the model
    via the '|' alternation that _model_in_text already understands, so the
    order number OR the designation counts as a match."""
    alias = _abb_alias_for(model, mfr)
    return f"{model}|{alias}" if alias else model

def _is_own_mfr_domain(url, mfr):
    """True when the URL lives on THIS part's manufacturer's own website.
    Used to hard-prioritise official documentation over distributors/aggregators."""
    if not mfr:
        return False
    d  = _dom(url)
    db = _base_dom(d)
    mu = mfr.upper()
    ml = re.sub(r'[^a-z]', '', mfr.lower())
    if len(ml) >= 4 and ml[:6 if len(ml) >= 6 else len(ml)] in re.sub(r'[^a-z]', '', db):
        return True
    for key, doms in _MFR_DOM_ALIASES.items():
        if key in mu and any(_base_dom(x) == db for x in doms):
            return True
    site = _mfr_doc_site(mfr)
    if site and _base_dom(site) == db:
        return True
    return False

def _throttle(url):
    d = _dom(url)
    with _domain_locks[d]:
        e = time.monotonic() - _domain_times[d]
        if e < MIN_DOMAIN_GAP: time.sleep(MIN_DOMAIN_GAP - e)
        _domain_times[d] = time.monotonic()

# Lighter pacing for cheap HEAD probes — they don't fetch payloads, so they
# don't need the full courtesy gap that real downloads use. Without this, N
# guessed URLs on ONE host (e.g. several euautomation shards) serialise at
# MIN_DOMAIN_GAP each, turning a "parallel" probe into an N*1.5s stall.
PROBE_DOMAIN_GAP = 0.2
_probe_times = defaultdict(float)
_probe_locks = defaultdict(threading.Lock)
def _throttle_probe(url):
    d = _dom(url)
    with _probe_locks[d]:
        e = time.monotonic() - _probe_times[d]
        if e < PROBE_DOMAIN_GAP: time.sleep(PROBE_DOMAIN_GAP - e)
        _probe_times[d] = time.monotonic()

# Hard ceiling for ANY single sniff/scrape network op. requests' `timeout=`
# only bounds idle gaps between bytes — an endpoint that DRIPS a byte every
# few seconds (e.g. Rittal's pdf-creator, which renders the PDF on the fly)
# can otherwise hold a worker for minutes with NO log output, which looks
# exactly like a freeze. These helpers enforce a real wall-clock budget.
SNIFF_CONNECT_TIMEOUT = 8     # (s) to establish connection + first byte
SNIFF_TOTAL_BUDGET    = 18    # (s) absolute cap on the whole streamed read

def _get_first_bytes(url, session, headers, n=1024,
                     connect_timeout=SNIFF_CONNECT_TIMEOUT,
                     total_budget=SNIFF_TOTAL_BUDGET):
    """GET `url` and return up to `n` leading bytes, guaranteeing the call
    returns within ~total_budget seconds even if the server drip-feeds.
    Returns b'' on any failure/timeout. Always closes the connection."""
    deadline = time.monotonic() + total_budget
    r = None
    try:
        r = session.get(url, timeout=connect_timeout, stream=True,
                        allow_redirects=True, headers=headers)
        if r.status_code not in (200, 206):
            return b""
        buf = b""
        for chunk in r.iter_content(512):
            if chunk:
                buf += chunk
                if len(buf) >= n:
                    break
            if time.monotonic() > deadline:    # drip-feed / stall — give up
                break
        return buf[:n]
    except Exception:
        return b""
    finally:
        if r is not None:
            try: r.close()
            except Exception: pass

def _get_text_capped(url, session, headers,
                     connect_timeout=SNIFF_CONNECT_TIMEOUT,
                     total_budget=SNIFF_TOTAL_BUDGET, max_bytes=3_000_000):
    """GET an HTML page as text within a hard wall-clock budget. b''→ "" on fail."""
    deadline = time.monotonic() + total_budget
    r = None
    try:
        r = session.get(url, timeout=connect_timeout, stream=True,
                        allow_redirects=True, headers=headers)
        if r.status_code != 200:
            return ""
        ct = (r.headers.get("Content-Type") or "").lower()
        if ct and not any(k in ct for k in ("html", "xml", "text")):
            return ""
        chunks, size = [], 0
        for chunk in r.iter_content(65536):
            if chunk:
                chunks.append(chunk); size += len(chunk)
                if size > max_bytes:
                    break
            if time.monotonic() > deadline:
                break
        return b"".join(chunks).decode(r.encoding or "utf-8", "replace")
    except Exception:
        return ""
    finally:
        if r is not None:
            try: r.close()
            except Exception: pass

def _is_pdf_url(url):
    u = url.lower(); return u.endswith(".pdf") or ".pdf?" in u or ".pdf#" in u

def _variants(m):
    v = {m.lower(), re.sub(r'[-_\s.]+','',m).lower(), re.sub(r'[-_]+',' ',m).lower()}
    v |= {t.lower() for t in re.findall(r'[A-Za-z]{2,}|\d{2,}', m)}
    return list(v)

# =============================================================================
#  DOCUMENT-TYPE CLASSIFICATION  —  "is this actually a manual?"
#  Layer 1: classify by URL filename + link title (pre-download, costs nothing)
#  Layer 2: classify by PDF content after download (see _verify_pdf below)
# =============================================================================

def _url_path(url):
    """Path+query of a URL, lowercased — domain excluded so that e.g.
    catalog.weidmueller.com isn't penalised for the word 'catalog'."""
    m = re.match(r'https?://[^/]+(/.*)?$', (url or '').strip())
    return ((m.group(1) or '') if m else (url or '')).lower()

_RX_NAME_MANUAL = re.compile(
    r'manual|handbuch|anleitung|instruction|operat(?:ing|ion)|installation|'
    r'montage|commissioning|user[-_ ]?guide|getting[-_ ]?started|quick[-_ ]?start|'
    r'\bqsg\b|programming|hardware[-_ ]?guide|[-_](?:um|in|rm|sg|qs)\d{3}', re.I)

_RX_NAME_DS = re.compile(
    r'datasheet|data[-_ ]sheet|technical[-_ ]data|specification|\bspecs?\b|'
    r'[-_]td\d{3}|product[-_ ]data', re.I)

_RX_NAME_REJECT = [
    ("license",      re.compile(r'licen[cs]e|\beula\b|lizenz', re.I)),
    ("certificate",  re.compile(
        r'certificat|declaration|conformit|konformit|zertifikat|attestation|'
        r'\brohs\b|\breach\b|\bpcn\b|product[-_ ]change|\bsdoc\b|'
        r'\bmsds\b|safety[-_ ]data[-_ ]sheet', re.I)),
    ("cad_drawing",  re.compile(
        r'\bcad\b|\bdwg\b|\bdxf\b|\bstp\b|\bstep\b(?![-_ ]?by)|\bigs\b|\biges\b|'
        r'\beplan\b|3[-_ ]?d[-_ ]?(model|file)?\b|outline[-_ ]?drawing|'
        r'dimension(?:al)?[-_ ]?drawing|\bdrawing\b|\bdrw\b|wiring[-_ ]macro|\bmacros?\b', re.I)),
    ("brochure",     re.compile(
        r'brochure|flyer|prospekt|katalog|\bcatalog(?:ue)?\b|'
        r'product[-_ ]overview|press[-_ ]release|\bposter\b', re.I)),
    ("other_reject", re.compile(
        r'warranty|garantie|terms[-_ ]and[-_ ]conditions|release[-_ ]notes?', re.I)),
]

_REJECT_TYPES = {"license", "certificate", "cad_drawing", "brochure",
                 "other_reject", "wrong_part", "unreadable"}

_HIMA_FAMILIES = {
    "F-BASE": ["himatrix", "hiquad", "baserack", "subrack", "base"],
    "F-PWR":  ["himatrix", "hiquad", "power supply", "powermodule", "power module", "power"],
    "F-CPU":  ["himatrix", "hiquad", "cpu", "processor", "hicore", "control"],
    "F-COM":  ["himatrix", "hiquad", "communication", "commodule", "com module", "ethernet", "profibus"],
    "F-IOP":  ["himatrix", "hiquad", "iop", "processor", "io", "input", "output"],
    "F-BLK":  ["himatrix", "hiquad", "blk", "blocking", "terminal"],
    "F3":     ["himatrix", "hiquad", "f3", "digital output", "output module", "output"],
    "F6":     ["himatrix", "hiquad", "f6", "analog input", "input module", "input"],
    "F7":     ["himatrix", "hiquad", "f7", "coupling", "input", "output", "safety"],
    "H7":     ["himatrix", "hiquad", "h7", "rack", "mounting", "system"],
    "894":    ["himatrix", "hiquad", "order", "catalog"],
    "892":    ["himatrix", "hiquad", "connector", "connection", "cable"],
    "Z 71":   ["himatrix", "hiquad", "cable", "connection", "z71"],
    "K 9":    ["himatrix", "hiquad", "cable", "connection", "k9", "connector"],
    "K9":     ["himatrix", "hiquad", "cable", "connection", "k9", "connector"],
}

_REJECT_LABEL = {
    "license":      "software license agreement",
    "certificate":  "certificate / declaration of conformity",
    "cad_drawing":  "CAD / dimensional drawing",
    "brochure":     "marketing brochure / catalogue",
    "other_reject": "non-manual document (warranty/terms/notes)",
    "wrong_part":   "different part — part number absent from document",
    "unreadable":   "corrupt or unreadable PDF",
}

def _name_doc_type(url, title):
    """Best-guess doc type from URL filename + title alone. Returns one of
    'manual' / 'datasheet' / a reject type / None (no signal)."""
    hay = _url_path(url) + " " + (title or "").lower()
    is_man = bool(_RX_NAME_MANUAL.search(hay))
    is_ds  = bool(_RX_NAME_DS.search(hay))
    rej = next((name for name, rx in _RX_NAME_REJECT if rx.search(hay)), None)
    if rej and not (is_man or is_ds):
        return rej
    if rej:
        return None          # conflicting signals — let content verification decide
    if is_man: return "manual"
    if is_ds:  return "datasheet"
    return None


def _relevant(url, title, model, mfr):
    haystack = url + " " + title
    mm = _match_model(model, mfr)          # ABB: also matches the type designation
    # Require the model as a whole token — EN2T must NOT match EN2TSC
    if _model_in_text(haystack, mm):
        return True
    
    # Suffix fallback for conformal coating
    if model.upper().endswith("-CC") and _model_in_text(haystack, model[:-3]):
        return True
    
    # MTL/Eaton-specific relevance fallback
    if mfr and ("MTL" in mfr.upper() or "EATON" in mfr.upper()):
        if _is_own_mfr_domain(url, mfr) or "mtl" in url.lower() or "eaton" in url.lower():
            digits_m = re.findall(r'\d{3,4}', model)
            if digits_m and any(d in haystack for d in digits_m):
                return True
    
    # HIMA-specific family mapping fallback
    if mfr and "HIMA" in mfr.upper():
        mu_model = model.upper().strip()
        for prefix, families in _HIMA_FAMILIES.items():
            if mu_model.startswith(prefix):
                hl = haystack.lower()
                if any(fam in hl for fam in families):
                    return True
                break

    # Manufacturer CDN domain is a strong enough signal on its own,
    # but still require at least one significant chunk of the model
    if mfr and len(mfr) >= 3:
        ml = re.sub(r'[-_\s]+', '', mfr).lower()
        if ml[:6] in re.sub(r'[-_\s.]+', '', haystack).lower() or _is_own_mfr_domain(url, mfr):
            chunks = [t for t in re.findall(r'[A-Za-z]{2,}|\d{3,}', model)]
            if any(_model_in_text(haystack, c) for c in chunks[:2]):
                return True
    return False

# Language codes manufacturers append to document filenames (8108235_Instructions_DE_EN)
_LANG_CODES = {"DE","FR","IT","ES","NL","SV","PL","RU","ZH","JA","CS","BG","HU","RO",
               "TR","DA","FI","NO","PT","KO","SK","HR","SL","ET","LV","LT","EL","UK"}
_EN_CODES   = {"EN","US","GB"}

def _lang_signal(url):
    """+1 = English available, -1 = foreign-language-only, 0 = no language info.
    Only reads UPPERCASE two-letter codes delimited in the URL filename
    (…_Instructions_DE_EN → DE,EN) so ordinary words never false-match."""
    codes = set(re.findall(r'[_-]([A-Z]{2})(?=[_\-.]|$)', url))
    if codes & _EN_CODES:               return 1
    if codes & _LANG_CODES:             return -1
    return 0

def _score(url, title, model, mfr, desc="", snippet=""):
    s=0; ul=url.lower(); tl=title.lower(); sl=(snippet or "").lower(); d=_dom(url)
    ns=re.sub(r'[-_\s.]+','',model).lower(); ml=(mfr or "").lower()
    mm=_match_model(model, mfr)                 # ABB: order number OR designation
    # The search-result snippet often spells out the part/designation even when
    # the URL and link-title don't — fold it into the matching haystack.
    part_signal = _model_in_text(f"{url} {title} {snippet}", mm)
    # Model match — whole-token only, strong penalty for substring-only matches
    if part_signal: s += 20
    elif ns in re.sub(r'[-_\s.]+','',ul+tl+sl): s += 6   # compact form found but not as whole token — weak
    # Language — prefer English variants; foreign-only ones stay as fallback
    lang = _lang_signal(url)
    if   lang > 0: s += 6
    elif lang < 0: s -= 10
    # Manufacturer / domain signals — the manufacturer's OWN site beats everything
    if _is_own_mfr_domain(url, mfr):
        # ABB's official library also hosts many GENERIC family/umbrella PDFs on
        # the same domain. Only reward the domain in full when THIS listing
        # actually references the part (order #, designation) or its
        # description — otherwise a catch-all doc would outrank the real one.
        if _is_abb(mfr) and not part_signal \
           and not any(k in ul or k in tl or k in sl for k in _desc_keywords(desc)[:4]):
            s += 12
        else:
            s += 25
    elif any(x in d for x in MFR_DOMS): s+=10
    elif any(x in d for x in DIST_DOMS): s+=5
    # Document-type signal — the single most important factor for accuracy
    nt = _name_doc_type(url, title)
    if   nt == "manual":    s += 14 if PREFER_MANUALS else 8
    elif nt == "datasheet": s += 6
    elif nt is not None:    s -= 40   # license / certificate / CAD / brochure …
    if _model_in_text(title, mm): s+=6
    elif model.lower() in tl: s+=3
    # Description keyword matches — each keyword found in title/url/snippet is extra confidence
    for kw in _desc_keywords(desc)[:4]:
        if kw in ul or kw in tl or kw in sl: s += 3
    # Freshness
    cur = datetime.now().year
    for offset in range(0, 4):
        if str(cur - offset) in ul or str(cur - offset) in tl:
            s += max(1, 4 - offset); break
    for old in range(cur - 5, cur - 15, -1):
        if str(old) in ul or str(old) in tl:
            s -= 4; break
    # Penalise if this is neither the manufacturer's site nor a known doc source
    if mfr and len(ml) >= 4 and not _is_own_mfr_domain(url, mfr) and not any(x in d for x in MFR_DOMS+DIST_DOMS):
        s -= 5
    # ABB: a generic family / umbrella / non-product PDF that doesn't name this
    # specific part must never become the auto-suggested pick over a real
    # part-specific datasheet.
    if _is_abb(mfr) and not part_signal:
        if "[family doc]" in tl or _ABB_FAMILY_RX.search(ul + " " + tl):
            s -= 18
    return s

def _scrape_page(page_url, model, mfr, session):
    if _is_cancelled(): return []
    try:
        _throttle(page_url)
        r = session.get(page_url, timeout=12, stream=True,
                        headers={**BROWSER_HDR,"Referer":page_url})
        if r.status_code != 200:
            r.close(); return []
        ct = (r.headers.get("Content-Type") or "").lower()
        if ct and "html" not in ct and "xml" not in ct and "text" not in ct:
            r.close(); return []          # don't download/parse binaries
        chunks, size = [], 0
        for ch in r.iter_content(65536):
            chunks.append(ch); size += len(ch)
            if size > 3_000_000: break    # 3 MB of HTML is plenty for link scraping
        r.close()
        html = b"".join(chunks).decode(r.encoding or "utf-8", "replace")
        soup = BeautifulSoup(html,"html.parser")
        base = page_url.split('/')[0]+'//'+page_url.split('/')[2]; found={}
        for a in soup.find_all("a",href=True):
            href=urljoin(base,a["href"]); txt=a.get_text(strip=True) or href
            if _is_pdf_url(href) and _relevant(href,txt,model,mfr): found[href]={"title":txt,"referer":page_url}
        for raw in re.findall(r'https?://[^\s"\'<>]+\.pdf',html):
            if _relevant(raw,raw,model,mfr) and raw not in found: found[raw]={"title":raw.split('/')[-1],"referer":page_url}
        return [{"url":u,"title":v["title"],"referer":v["referer"]} for u,v in found.items()]
    except Exception: return []

_MOUSER_EMPTY = [0]   # consecutive zero-hit searches — Mouser bot-challenge detector

def _mouser_scrape(model, mfr, session):
    """Scrape Mouser search results for datasheet links — no API key needed.
    Mouser product pages reliably carry the manufacturer datasheet PDF.
    """
    if _cb.is_open("Mouser-scrape"): return []
    try:
        model_q = model.split('|')[0]
        q = f"{mfr} {model_q}".strip() if mfr else model_q
        search_url = (f"https://www.mouser.com/Search/Refine"
                      f"?Keyword={requests.utils.quote(q)}&FS=True&Ns=Pricing|0")
        _throttle(search_url)
        r = session.get(search_url, timeout=15,
                        headers={**BROWSER_HDR, "Referer": "https://www.mouser.com/"})
        if r.status_code != 200:
            _cb.record_failure("Mouser-scrape", f"HTTP {r.status_code}"); return []
        soup = BeautifulSoup(r.text, "html.parser"); found = []

        # Pass 1: direct PDF links on the search results page
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if not href: continue
            full = urljoin("https://www.mouser.com", href)
            txt  = a.get_text(strip=True) or ""
            if _is_pdf_url(full) and _relevant(full, txt, model, mfr):
                found.append({"url": full, "title": txt or f"{model} datasheet",
                              "referer": search_url})

        # Pass 2: follow up to 3 product-detail pages and scrape their datasheet links
        if not found:
            prod_links = []
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if "/ProductDetail/" in href:
                    txt = a.get_text(strip=True)
                    if _model_in_text(href + " " + txt, model):
                        prod_links.append(urljoin("https://www.mouser.com", href))
            for prod_url in prod_links[:3]:
                for item in _scrape_page(prod_url, model, mfr, session):
                    found.append(item)

        if found:
            _MOUSER_EMPTY[0] = 0
            _cb.record_success("Mouser-scrape")
        else:
            # Mouser serves a JS bot-challenge page (HTTP 200, zero parseable
            # results) to plain-requests clients. After 4 consecutive empties,
            # stop spending ~15s per part on a backend that can't see results.
            _MOUSER_EMPTY[0] += 1
            if _MOUSER_EMPTY[0] >= 4:
                _cb.trip("Mouser-scrape",
                         "0 hits in 4 consecutive searches (likely bot-challenge "
                         "page) — disabling for this run")
        tprint(f"    [Mouser-scrape] {len(found)} hit(s) for {mfr} {model}".strip())
        return found[:6]
    except Exception as e:
        _cb.record_failure("Mouser-scrape", str(e)[:80]); return []


def _rs_scrape(model, mfr, session):
    """Scrape RS Online / RS Components for datasheet links — no API key needed."""
    if _cb.is_open("RS-scrape"): return []   # silently skip — already failed this run
    try:
        model_q = model.split('|')[0]
        q = f"{mfr} {model_q}".strip() if mfr else model_q
        search_url = f"https://uk.rs-online.com/web/c/?searchTerm={requests.utils.quote(q)}"
        _throttle(search_url)
        r = session.get(search_url, timeout=15,
                        headers={**BROWSER_HDR, "Referer": "https://uk.rs-online.com/"})
        if r.status_code != 200:
            if r.status_code in (403, 429, 503):   # bot-blocked — decisive, no retry value
                _cb.trip("RS-scrape", f"HTTP {r.status_code} — disabling for this run")
            else:
                _cb.record_failure("RS-scrape", f"HTTP {r.status_code}")
            return []
        soup = BeautifulSoup(r.text, "html.parser"); found = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if not href: continue
            full = urljoin("https://uk.rs-online.com", href)
            txt  = a.get_text(strip=True) or ""
            if _is_pdf_url(full) and _relevant(full, txt, model, mfr):
                found.append({"url": full, "title": txt or f"{model} datasheet",
                              "referer": search_url})
        _cb.record_success("RS-scrape")
        if found: tprint(f"    [RS-scrape] {len(found)} hit(s) for {mfr} {model}".strip())
        return found[:4]
    except Exception as e:
        _cb.record_failure("RS-scrape", str(e)[:80]); return []


def _hima_queries(model: str) -> list[str]:
    """
    HIMA uses descriptive module names ("F-BASE RACK 01", "F-CPU 01") rather than
    traditional part numbers. These don't appear verbatim in third-party PDFs.
    Generate broader queries that find the right product family documentation.

    Module → product family / alternate names
    """
    m = model.upper().strip()
    short = m.split()[0]  # e.g. "F-BASE" from "F-BASE RACK 01"

    # Map module prefixes to HIMA product family context
    _FAMILIES = {
        "F-BASE": ["HIMatrix FX", "HIMatrix SX", "HIMatrix MX", "HIQuad X", "base rack installation"],
        "F-PWR":  ["HIMatrix power supply", "HIQuad power supply", "power module"],
        "F-CPU":  ["HIMatrix CPU module", "HIQuad CPU", "HICore processor"],
        "F-COM":  ["HIMatrix communication", "HICore COM module", "HIMatrix FX COM"],
        "F-IOP":  ["HIMatrix I/O processor", "IOP module HIMatrix"],
        "F-BLK":  ["HIMatrix F60 BLK", "blocking module HIMatrix"],
        "F3":     ["HIMA F3 series", "digital output module HIMA"],
        "F6":     ["HIMA F6 series", "analog input module HIMA"],
        "F7":     ["HIMA F7 series", "HIMatrix module", "HIQuad module"],
        "H7":     ["HIMA H7 series", "HIQuad rack", "mounting rack"],
        "894":    ["HIMA HIMatrix", "HIQuad X order"],
        "892":    ["HIMA connector", "HIMA cable", "connection module"],
        "Z 71":   ["HIMA cable assembly", "Z 71 connection"],
        "K 9":    ["HIMA system cable", "HIMA connection"],
        "K9":     ["HIMA system cable", "HIMA connection"],
    }

    queries = []

    # 1. Broad family system manual queries (highest success rate for generic/family docs)
    queries += [
        f"HIMatrix system manual pdf",
        f"HIQuad system manual pdf",
        f"HIMA HIMatrix safety manual pdf",
    ]

    # 2. Specific family-specific alias query (unquoted, prefix-free to prevent confusing search engine)
    for prefix, aliases in _FAMILIES.items():
        if m.startswith(prefix):
            for alias in aliases[:2]:
                queries.append(f'HIMA {alias} manual pdf')
            break

    # 3. Model queries (unquoted for fuzzy matching)
    queries.append(f'HIMA {model} datasheet')
    queries.append(f'HIMA {model} manual pdf')

    # 4. Model + family alias queries
    for prefix, aliases in _FAMILIES.items():
        if m.startswith(prefix):
            for alias in aliases[:2]:
                queries.append(f'HIMA {model} {alias}')
            if short != m and len(short) >= 4:
                queries.append(f'HIMA {short} datasheet pdf')
                queries.append(f'{short} HIMA safety module datasheet')
            break

    # 5. Also try known third-party HIMA documentation sites directly
    queries += [
        f'site:sds-automatyka.pl HIMA {short}',
        f'site:rms-dcs.com HIMA {short}',
        f'HIMA {short} "data sheet" OR "technical data" filetype:pdf',
    ]

    return queries


def _rank_candidates(direct, model, mfr, desc):
    """Score, filter and order candidate URLs.
    Hard rule: the manufacturer's OWN website always outranks distributors and
    third-party archives — third-party docs are only tried when the official
    ones fail download or content verification.
    Known-junk document types (licenses, certificates, CAD drawings, brochures)
    never survive — not even in the low-confidence fallback."""
    def _tier(u, v):
        nt = _name_doc_type(u, v["title"])
        return 1 if (_is_own_mfr_domain(u, mfr) and nt not in _REJECT_TYPES) else 0
    scored = [(u, v, _score(u, v["title"], model, mfr, desc, v.get("snippet","")), _tier(u, v))
              for u, v in direct.items()
              # never spend a download on a name-identified license/cert/CAD/brochure
              if _name_doc_type(u, v["title"]) not in _REJECT_TYPES]
    good = [t for t in scored if t[2] >= 10]
    if not good:   # nothing passed threshold — fall back to best 3 (junk already gone)
        good = sorted(scored, key=lambda t: t[2], reverse=True)[:3]
    good.sort(key=lambda t: (t[3], t[2]), reverse=True)   # (own-mfr-domain, score)
    return [{"url": u, "title": v["title"], "referer": v["referer"]}
            for u, v, _s, _t in good[:10]]


def _find_pdfs(model, mfr, session, desc="", force=False):
    if _is_cancelled(): return []
    mfr = _fix_mfr_spelling(mfr)   # 'PHEONIX CONTACT' etc. → canonical spelling
    pf = f"{mfr} {model}".strip() if mfr else model
    tprint(f"    [Search] {pf}: checking manufacturer sources…")
    direct = {}
    abb_alias = None        # ABB type designation, resolved lazily below

    def _add(hits):
        for h in hits:
            direct.setdefault(h["url"], {"title": h["title"], "referer": h.get("referer")})
        if hits and _is_abb(mfr):
            _abb_register_alias(model, mfr, direct)

    # Per-part wall-clock budget. Only applies when no interactive (GUI
    # re-search) deadline is active — the GUI sets its own, usually longer one.
    part_deadline = (time.monotonic() + PART_SEARCH_BUDGET
                     if (PART_SEARCH_BUDGET and _search_deadline[0] is None)
                     else None)
    # Publish to this worker thread so _ddg / _should_stop_search can enforce
    # the budget even while a part is WAITING IN LINE for the global DDG lock.
    _part_deadline_tl.t = part_deadline
    _ddg_part_calls_tl.n = 0

    def _out_of_budget():
        """True when the interactive or per-part search-time budget is spent."""
        t_dl = getattr(_part_deadline_tl, "t", None)
        if _deadline_passed() or (t_dl is not None and time.monotonic() > t_dl):
            tprint(f"    [Search] {pf}: time budget reached — "
                   f"ranking the {len(direct)} candidate(s) found so far")
            return True
        return False

    # ── Tier 1: DirectProbe (hardcoded CDN URL patterns) ─────────────────────
    hits = _direct_probe(model, mfr, session)
    _add(hits)
    if hits: tprint(f"    [DirectProbe] {len(hits)} hit(s) for {pf}")
    if _is_cancelled(): return []
    if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)

    # ── Tier 1.5: Manufacturer portal — go straight to the source ────────────
    # Scrapes the manufacturer's own product/documentation page directly.
    # Far more accurate than any search engine for known manufacturers.
    portal_hits = _mfr_portal_search(model, mfr, session, force=force)
    _add(portal_hits)
    if _is_cancelled(): return []

    # If the manufacturer's own site already gave us usable (non-junk) docs,
    # skip the slow search-engine cascade. If those official docs include a
    # MANUAL we are completely done; if they are only datasheets, spend exactly
    # ONE quick search looking for a manual to pair with them, then rank.
    own = [(u, v) for u, v in direct.items()
           if _is_own_mfr_domain(u, mfr)
           and _name_doc_type(u, v["title"]) not in _REJECT_TYPES]
    # HEAD-verified DirectProbe hits are deterministic mirrors of OFFICIAL docs
    # (DigiKey media CDN, Rockwell literature, Rittal pdf-creator, …). They are
    # just as good as manufacturer-domain docs for the purpose of skipping the
    # engine cascade — content verification after download still has the final
    # say. Without this, a part with a live mirror datasheet in hand would still
    # burn the full DDG budget "looking for official docs" it can't reach.
    probe_urls = {h["url"] for h in hits}
    mirror = [] if own else [(u, v) for u, v in direct.items()
                             if u in probe_urls
                             and _name_doc_type(u, v["title"]) not in _REJECT_TYPES]
    trusted = own or mirror
    if trusted:
        has_official_manual = any(_name_doc_type(u, v["title"]) == "manual"
                                  for u, v in trusted)
        tprint(f"    [{'MfrPortal' if own else 'DirectProbe'}] {pf}: "
               f"{len(trusted)} {'official' if own else 'verified mirror'} doc(s) "
               f"— skipping search engines")
        _add(_component_db_docs(model, mfr))
        if (not has_official_manual and PREFER_MANUALS and not _is_cancelled()
                and _ddg_budget_left()):
            try:
                for r in _ddg(f'"{model}" {mfr} manual', n=6, force=force):
                    url = (r.get("href") or "").strip()
                    title = (r.get("title") or "").strip()
                    if url and _is_pdf_url(url) and _relevant(url, title, model, mfr):
                        direct.setdefault(url, {"title": title, "referer": None})
            except Exception:
                pass
        return _rank_candidates(direct, model, mfr, desc)

    # ── Tier 2: Nexar component database (if key configured) ──────────────────
    _add(_component_db_docs(model, mfr))
    if _is_cancelled(): return []
    if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)

    # FAIL FAST: if every search engine is currently rate-limited/paused and we
    # have no API-key backends configured, the entire search-engine cascade
    # below can only return empty — so don't waste 15-20s per part walking it.
    # Mark the part not-found immediately and let the run move on. (Affected
    # parts surface cleanly in Search-again, where the cooldown has expired.)
    if (not direct and not force
            and time.monotonic() < _DDG_PAUSED_UNTIL[0]
            and not _any_search_backend_ready()):
        cooldown_left = int(_DDG_PAUSED_UNTIL[0] - time.monotonic())
        tprint(f"    [Search] {pf}: search engines cooling down "
               f"({cooldown_left}s left) — skipping, retry later")
        return []

    tprint(f"    [Search] {pf}: no official docs yet — querying search engines…")

    # ── Tier 3: Manufacturer-portal site-specific DDG search ─────────────────
    # DDG honours site: — searching literature.rockwellautomation.com directly
    # returns exact documents rather than random distributor pages.
    doc_site = _mfr_doc_site(mfr)
    if doc_site:
        ms = re.sub(r'[-_\s.]+', '', model)
        site_qs = [
            f'site:{doc_site} "{model}"',
        ]
        if ms.lower() != model.lower():
            site_qs.append(f'site:{doc_site} "{ms}"')
        pages = []
        for q in site_qs:
            if _is_cancelled(): return []
            if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)
            for r in _ddg(q, n=8, force=force, critical=(not direct)):          # bypass _search so DDG always runs here
                url = r.get("href", "").strip(); title = r.get("title", "").strip()
                if not url: continue
                if _is_pdf_url(url) and _relevant(url, title, model, mfr):
                    direct.setdefault(url, {"title": title, "referer": None})
                elif url not in pages:
                    pages.append(url)
        if pages and not _is_cancelled():
            with ThreadPoolExecutor(max_workers=SCRAPE_WORKERS) as ex:
                for fut in as_completed(
                    {ex.submit(_scrape_page, pg, model, mfr, session): pg
                     for pg in pages[:6]}
                ):
                    _add(fut.result())
        if direct:
            tprint(f"    [SiteSearch:{doc_site}] {len(direct)} candidate(s)")
            own_site_docs = [u for u in direct if _is_own_mfr_domain(u, mfr) or (doc_site and doc_site in _dom(u))]
            if own_site_docs:
                has_official_manual = any(_name_doc_type(u, direct[u]["title"]) == "manual"
                                          for u in own_site_docs)
                tprint(f"    [SiteSearch] {pf}: {len(own_site_docs)} official doc(s) found — skipping further search")
                if (not has_official_manual and PREFER_MANUALS and not _is_cancelled()
                        and _ddg_budget_left()):
                    try:
                        for r in _ddg(f'"{model}" {mfr} manual', n=6, force=force):
                            url = (r.get("href") or "").strip()
                            title = (r.get("title") or "").strip()
                            if url and _is_pdf_url(url) and _relevant(url, title, model, mfr):
                                direct.setdefault(url, {"title": title, "referer": None})
                    except Exception:
                        pass
                return _rank_candidates(direct, model, mfr, desc)

    # ── Tier 3.5: HIMA-specific multi-query search ───────────────────────────
    # HIMA module names ("F-BASE RACK 01") don't appear verbatim in docs.
    # We generate product-family-aware queries to find the right manuals.
    gen_pages = []   # shared page list used by Tier 3.5 and Tier 5
    if mfr and "HIMA" in mfr.upper() and not direct:
        for q in _hima_queries(model):
            if _is_cancelled(): break
            for r in _ddg(q, n=6, force=force, critical=(not direct)):
                url = r.get("href", "").strip(); title = r.get("title", "").strip()
                if not url: continue
                if _is_pdf_url(url) and _relevant(url, title, model, mfr):
                    direct.setdefault(url, {"title": title, "referer": None})
                elif not _is_pdf_url(url) and url not in gen_pages:
                    gen_pages.append(url)
            if direct:
                valid_count = sum(1 for u in direct if "dex.cz" not in u)
                if valid_count >= 1 or len(direct) >= 3:
                    tprint(f"    [HIMA-search] {len(direct)} candidate(s) found after query: {q[:60]}")
                    break

    # ── Tier 4: Distributor scrapers (no API key required) ────────────────────
    if _is_cancelled(): return []
    if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)
    _add(_mouser_scrape(model, mfr, session))
    _add(_rs_scrape(model, mfr, session))
    if _is_cancelled(): return []
    if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)
    tprint(f"    [Search] {pf}: broad web search…")

    # ── ABB designation augmentation ─────────────────────────────────────────
    # ABB documents a part under its TYPE designation (AO815) far more reliably
    # than under the order number (3BSE052605R1). Resolve the designation, then
    # search for it directly: this rescues parts the order-number search can't
    # find, and (via the module-level alias cache) lets content-verification
    # accept the right document afterwards.
    if _is_abb(mfr) and not _is_cancelled():
        abb_alias = _abb_register_alias(model, mfr, direct)
        if not abb_alias and not _out_of_budget():
            try:
                res = _search(f'ABB "{model}"', n=8, force=force, critical=(not direct))
                abb_alias = _abb_register_alias(model, mfr, res)
            except Exception:
                abb_alias = None
        if abb_alias:
            for q in (f'ABB {abb_alias} "{model}" datasheet',
                      f'ABB {abb_alias} datasheet',
                      f'ABB {abb_alias} manual'):
                if _is_cancelled() or _out_of_budget():
                    break
                for r in _search(q, n=8, force=force, critical=(not direct)):
                    url = r.get("href", "").strip(); title = r.get("title", "").strip()
                    body = (r.get("body") or "").strip()
                    if not url:
                        continue
                    if _is_pdf_url(url) and _relevant(url, f"{title} {body}", model, mfr):
                        direct.setdefault(url, {"title": title, "snippet": body, "referer": None})
                    elif not _is_pdf_url(url) and _relevant(url, f"{title} {body}", model, mfr) \
                            and url not in gen_pages:
                        gen_pages.append(url)
                if direct:
                    break

    # ── Tier 5: General DDG search + page scraping (broad fallback) ───────────
    ms     = re.sub(r'[-_\s.]+', '', model)   # dotless too: 8108.235 → 8108235
    kw     = _desc_keywords(desc)
    kw_str = " ".join(kw[:2]) if kw else ""
    term   = f'{mfr} "{model}"' if mfr else f'"{model}"'
    gen_qs = [
        f'{term} {kw_str} manual'.strip()     if kw_str else f'{term} manual',
        f'{term} {kw_str} datasheet'.strip() if kw_str else f'{term} datasheet',
    ]
    if model.upper().endswith("-CC"):
        model_alt = model[:-3]
        term_alt = f'{mfr} "{model_alt}"' if mfr else f'"{model_alt}"'
        gen_qs.append(f'{term_alt} manual')
        gen_qs.append(f'{term_alt} datasheet')
    if ms.lower() != model.lower():
        mfr_ms_term = f'{mfr} "{ms}"' if mfr else f'"{ms}"'
        gen_qs.append(f'{mfr_ms_term} datasheet')
    # If we already hold candidates, the broad search is only a top-up — run 2
    # queries instead of ~7. Cuts global DDG pressure massively on big BOMs,
    # which keeps the shared query queue short for the parts that have nothing.
    for q in (gen_qs if not direct else gen_qs[:2]):
        if _is_cancelled(): return []
        if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)
        for r in _search(q, n=10, force=force, critical=(not direct)):
            url = r.get("href", "").strip(); title = r.get("title", "").strip()
            body = (r.get("body") or "").strip()
            if not url: continue
            if _is_pdf_url(url):
                if _relevant(url, f"{title} {body}", model, mfr):
                    direct.setdefault(url, {"title": title, "snippet": body, "referer": None})
            else:
                mh = mfr and len(mfr) >= 4 and mfr.lower()[:5] in _dom(url)
                if (_relevant(url, f"{title} {body}", model, mfr) or mh) and url not in gen_pages:
                    gen_pages.append(url)
    if gen_pages and not _is_cancelled():
        with ThreadPoolExecutor(max_workers=SCRAPE_WORKERS) as ex:
            for fut in as_completed(
                {ex.submit(_scrape_page, pg, model, mfr, session): pg
                 for pg in gen_pages[:6]}
            ):
                _add(fut.result())

    # ── Tier 6: Last-ditch ────────────────────────────────────────────────────
    if _is_cancelled(): return []
    if _out_of_budget(): return _rank_candidates(direct, model, mfr, desc)
    if not direct:
        pdf_term = f'{mfr} "{model}"' if mfr else f'"{model}"'
        for r in _search(f'{pdf_term} {kw_str} pdf'.strip(), n=8, force=force, critical=True):
            url = r.get("href", "").strip()
            if url and _is_pdf_url(url):
                direct.setdefault(url, {"title": r.get("title", ""), "referer": None})

    # ── Tier 7: FAMILY-level fallback (no part number in the query) ──────────
    # Some parts (enclosure systems, configured articles) only have family
    # documentation whose name/content never mentions the article number —
    # e.g. "Rittal TS 8 Specification Guide". Searched by manufacturer +
    # description keywords. These rank low and usually fail strict
    # verification, but they surface in the Search-again modal so the USER
    # can choose them deliberately.
    if not direct and mfr and not _is_cancelled() and not _deadline_passed():
        fam_kw = " ".join(_desc_keywords(desc)[:3])
        fam_qs = []
        # ABB: a designation-targeted query usually returns the ACTUAL product
        # datasheet, not a random family doc — try it first and keep it as a
        # normal (un-prefixed) candidate so it can be auto-suggested.
        if _is_abb(mfr) and abb_alias:
            for q in (f'ABB {abb_alias} datasheet pdf', f'ABB {abb_alias} manual pdf'):
                if _is_cancelled() or _deadline_passed(): break
                for r in _search(q, n=8, force=force, critical=True):
                    url = r.get("href", "").strip(); title = r.get("title", "").strip()
                    if url and _is_pdf_url(url) and _relevant(url, title, model, mfr):
                        direct.setdefault(url, {"title": title, "referer": None})
                if direct:
                    break
        fam_qs = [q for q in (
            f'{mfr} {fam_kw} manual pdf'.strip(),
            f'{mfr} {fam_kw} specification guide'.strip() if fam_kw else "",
        ) if q]
        for q in fam_qs:
            if direct: break
            if _is_cancelled() or _deadline_passed(): break
            for r in _search(q, n=8, force=force, critical=True):
                url = r.get("href", "").strip(); title = r.get("title", "").strip()
                if not url: continue
                ml2 = mfr.lower()[:6]
                if _is_pdf_url(url):
                    if ml2 in url.lower() or ml2 in title.lower():
                        direct.setdefault(url, {"title": f"[family doc] {title}",
                                                "referer": None})
                elif _is_own_mfr_domain(url, mfr) and \
                        any(k in (url + title).lower()
                            for k in ("manual", "guide", "instruction", "imf/",
                                      "handbook", "specification")):
                    # manufacturers often serve PDFs without a .pdf extension
                    # (e.g. rittal.com/imf/…) — magic-byte check happens at download
                    direct.setdefault(url, {"title": f"[family doc] {title}",
                                            "referer": None})
            if direct:
                tprint(f"    [FamilySearch] {len(direct)} family-level doc(s) for {pf} "
                       f"— pick manually, these are not part-specific")
                break

    # ── Rank and filter ───────────────────────────────────────────────────────
    # Final chance to resolve the ABB designation from whatever titles we ended
    # up with (e.g. a family-doc title that names the part) — the cached alias is
    # then used by content-verification to accept the right document.
    if _is_abb(mfr):
        _abb_register_alias(model, mfr, direct)
    return _rank_candidates(direct, model, mfr, desc)

def _sanitize(s,n=50): return re.sub(r'[^\w\-_.]','_',s).strip('_')[:n]

def _expected_filename(idx, mfr, model):
    label = f"{mfr} {model}".strip() if mfr else model
    return f"{idx:03d}_{_sanitize(label.replace(' ','_'))}.pdf"


# Domains/paths that block anonymous PDF downloads — try alternative URL forms first
_URL_ALT_TRANSFORMS = [
    # HIMA serves PDFs via SharePoint-sync (auth required) but also via fileadmin (public)
    ("www.hima.com/sharepoint-sync/PDFs/", "www.hima.com/fileadmin/PDFs/"),
    ("www.hima.com/sharepoint-sync/PDFs/", "www.hima.com/media/PDFs/"),
]

def _try_alt_url(url: str) -> list[str]:
    """Return alternative URL forms when the original is behind auth."""
    alts = []
    for blocked, alt in _URL_ALT_TRANSFORMS:
        if blocked in url:
            alts.append(url.replace(blocked, alt, 1))
    return alts

_DOWNLOADED_URLS = {}
_DOWNLOAD_IN_PROGRESS = {}
_DOWNLOAD_CACHE_LOCK = threading.Lock()

def _get_cached_or_download(url, dest, session, referer=None):
    """
    Checks if `url` has already been downloaded.
    If yes, links/copies the cached file to `dest` and returns True.
    If no, downloads it, caches it on success, and returns True.
    Uses an Event to make other threads wait if a download is in progress.
    """
    with _DOWNLOAD_CACHE_LOCK:
        # 1. Check if already downloaded
        if url in _DOWNLOADED_URLS:
            src = _DOWNLOADED_URLS[url]
            if src.exists():
                try:
                    if dest.exists():
                        dest.unlink()
                    os.link(str(src), str(dest))
                except Exception:
                    try:
                        shutil.copy2(src, dest)
                    except Exception:
                        return False
                try:
                    kb = dest.stat().st_size // 1024
                    tprint(f"         {_green('Saved:')} {dest.name}  ({kb} KB, cached)")
                except Exception:
                    pass
                return True
            else:
                del _DOWNLOADED_URLS[url]

        # 2. Check if download in progress
        if url in _DOWNLOAD_IN_PROGRESS:
            event = _DOWNLOAD_IN_PROGRESS[url]
        else:
            event = threading.Event()
            _DOWNLOAD_IN_PROGRESS[url] = event
            event = None

    if event is not None:
        event.wait()
        with _DOWNLOAD_CACHE_LOCK:
            if url in _DOWNLOADED_URLS:
                src = _DOWNLOADED_URLS[url]
                if src.exists():
                    try:
                        if dest.exists():
                            dest.unlink()
                        os.link(str(src), str(dest))
                    except Exception:
                        try:
                            shutil.copy2(src, dest)
                        except Exception:
                            return False
                    try:
                        kb = dest.stat().st_size // 1024
                        tprint(f"         {_green('Saved:')} {dest.name}  ({kb} KB, cached)")
                    except Exception:
                        pass
                    return True
        with _DOWNLOAD_CACHE_LOCK:
            if url not in _DOWNLOAD_IN_PROGRESS:
                event = threading.Event()
                _DOWNLOAD_IN_PROGRESS[url] = event
                event = None
            else:
                event = _DOWNLOAD_IN_PROGRESS[url]
        if event is not None:
            event.wait()
            with _DOWNLOAD_CACHE_LOCK:
                if url in _DOWNLOADED_URLS:
                    src = _DOWNLOADED_URLS[url]
                    if src.exists():
                        try:
                            if dest.exists():
                                dest.unlink()
                            os.link(str(src), str(dest))
                        except Exception:
                            try:
                                shutil.copy2(src, dest)
                            except Exception:
                                return False
                        try:
                            kb = dest.stat().st_size // 1024
                            tprint(f"         {_green('Saved:')} {dest.name}  ({kb} KB, cached)")
                        except Exception:
                            pass
                        return True
            return False

    success = _write_pdf(url, dest, session, referer)

    with _DOWNLOAD_CACHE_LOCK:
        ev = _DOWNLOAD_IN_PROGRESS.pop(url, None)
        if success:
            _DOWNLOADED_URLS[url] = dest
        if ev:
            ev.set()

    return success

def _write_pdf(url, dest, session, referer=None):
    # For known blocked CDN paths, probe alternative URLs first
    alts = _try_alt_url(url)
    urls_to_try = alts + [url] if alts else [url]

    for attempt_url in urls_to_try:
        for i, hdrs in enumerate([
            {**BROWSER_HDR, **({"Referer": referer} if referer else {})},
            {"Accept": "*/*", "User-Agent": BROWSER_UA},
            {"Accept": "application/pdf,*/*", "Referer": referer or attempt_url, "User-Agent": BROWSER_UA},
        ], 1):
            if _is_cancelled():
                if dest.exists():
                    try: dest.unlink()
                    except Exception: pass
                return False
            try:
                _throttle(attempt_url)
                r = session.get(attempt_url, timeout=30, stream=True, headers=hdrs)
                if r.status_code in (403, 401, 406, 429) and i < 3:
                    time.sleep(2 ** i); continue
                if r.status_code in (403, 401) and attempt_url != url:
                    break  # this alt URL also blocked — try next
                r.raise_for_status()
                cancelled_mid = False
                with open(dest, "wb") as f:
                    for chunk in r.iter_content(16384):
                        if _is_cancelled():
                            cancelled_mid = True
                            break
                        f.write(chunk)
                if cancelled_mid:
                    try: r.close()
                    except Exception: pass
                    try: dest.unlink()
                    except Exception: pass
                    return False
                kb = dest.stat().st_size // 1024
                if kb < 10: dest.unlink(); break
                with open(dest, "rb") as f:
                    head = f.read(5)
                    if not head.startswith(b"%PDF"): dest.unlink(); break
                    # Completeness check: a valid PDF ends with %%EOF. Server-
                    # rendered endpoints (e.g. Rittal pdf-creator) sometimes drop
                    # the connection mid-stream, leaving a file that STARTS with
                    # %PDF but is truncated — which later fails pypdf/pdfplumber
                    # parsing and gets wrongly rejected as "unreadable". If the
                    # trailer is missing and we have another header/URL to try,
                    # retry rather than keep the corrupt file.
                    f.seek(max(0, dest.stat().st_size - 2048))
                    tail = f.read()
                    if b"%%EOF" not in tail:
                        dest.unlink()
                        if i < 3:
                            time.sleep(1); continue
                        break
                if attempt_url != url:
                    tprint(f"         [AltURL] used {_dom(attempt_url)} instead of blocked CDN")
                tprint(f"         {_green('Saved:')} {dest.name}  ({kb} KB)")
                return True
            except requests.exceptions.HTTPError as e:
                if e.response.status_code in (403, 401, 406, 429) and i < 3:
                    time.sleep(2 ** i); continue
                break
            except Exception:
                break

    if dest.exists(): dest.unlink()
    return False

# =============================================================================
#  CONTENT VERIFICATION  —  Layer 2: open the PDF and check what it really is
#  Uses pdfplumber (already a dependency). Completely free, runs offline.
# =============================================================================

_W_MANUAL = [
    ("table of contents", 3), ("user manual", 3), ("instruction manual", 3),
    ("operating instructions", 3), ("operating manual", 3), ("this manual", 3),
    ("installation manual", 3), ("installation instruction", 2),
    ("original instructions", 2), ("user guide", 3), ("reference manual", 3),
    ("betriebsanleitung", 3), ("bedienungsanleitung", 3), ("montageanleitung", 2),
    ("assembly instructions", 3), ("assembly and operating instructions", 3),
    ("montage- und bedienungsanleitung", 3),
    ("safety instructions", 2), ("intended use", 2), ("commissioning", 2),
    ("maintenance", 2), ("troubleshooting", 2), ("important user information", 2),
    ("revision history", 1), ("wiring", 1), ("mounting", 1),
    ("warning", 1), ("caution", 1),
]
_W_DS = [
    ("datasheet", 3), ("data sheet", 3), ("technical data", 2),
    ("technical specifications", 2), ("specifications", 2),
    ("ordering information", 2), ("ordering data", 2),
    ("electrical characteristics", 2), ("absolute maximum ratings", 2),
    ("approvals", 1), ("accessories", 1), ("characteristics", 1), ("ratings", 1),
    # Spec-table vocabulary used by Rittal / Phoenix / Weidmüller etc.
    ("article no", 2), ("order no", 2), ("model no", 2), ("packs of", 2),
    ("net weight", 2), ("gross weight", 1), ("customs tariff", 2),
    ("supply includes", 2), ("scope of delivery", 2), ("surface finish", 1),
    ("basic material", 1), ("ip protection category", 2), ("to fit", 1),
    ("eclass", 1), ("ean", 1), ("dimensions", 1), ("colour", 1),
]
_W_LIC = [
    ("license agreement", 3), ("end user license", 3), ("eula", 2),
    ("licensee", 2), ("licensor", 2), ("grant of license", 2),
    ("gnu general public license", 3), ("open source software", 2),
    ("lizenzvereinbarung", 3), ("terms of use", 1),
]
_W_CERT = [
    ("declaration of conformity", 4), ("eu declaration", 3), ("ec declaration", 3),
    ("declaration of incorporation", 3), ("certificate of compliance", 3),
    ("certificate of conformity", 3), ("attestation", 2), ("hereby declare", 2),
    ("notified body", 2), ("konformitätserklärung", 4), ("certificate number", 2),
    ("rohs", 1), ("reach regulation", 1),
]
_W_CAD = [
    ("do not scale", 3), ("drawing number", 3), ("third angle projection", 3),
    ("first angle projection", 3), ("tolerances unless otherwise", 3),
    ("all dimensions", 2), ("dimensions in mm", 2), ("drawn by", 2),
    ("sheet 1 of", 2), ("scale 1:", 2), ("this drawing", 2),
]
_W_BROCH = [
    ("brochure", 3), ("your benefits", 2), ("product overview", 1), ("highlights", 1),
]

def _pypdf_pagecount(path):
    """Page count via pypdf (fallback when pdfplumber fails). 0 on failure."""
    try:
        return len(PdfReader(str(path)).pages)
    except Exception:
        return 0

def _pdf_is_complete(path):
    """True if the file looks like a structurally complete PDF (has %%EOF near
    the end). Cheap guard against truncated server-rendered downloads."""
    try:
        size = path.stat().st_size
        with open(path, "rb") as f:
            if not f.read(5).startswith(b"%PDF"):
                return False
            f.seek(max(0, size - 2048))
            return b"%%EOF" in f.read()
    except Exception:
        return False

def _pdf_sample(path, max_pages=6):
    """(page_count, lowercased text of first pages, avg chars/page) — never raises."""
    try:
        with pdfplumber.open(str(path)) as pdf:
            n = len(pdf.pages)
            chunks, chars = [], 0
            for pg in pdf.pages[:max_pages]:
                try: t = pg.extract_text() or ""
                except Exception: t = ""
                chunks.append(t); chars += len(t)
            avg = chars / max(1, min(n, max_pages))
            return n, "\n".join(chunks).lower(), avg
    except Exception:
        return 0, "", 0.0

def _kw_score(text, table):
    return sum(w for kw, w in table if kw in text)

def _verify_pdf(path, model, mfr, name_hint=None, src_url="", src_title="", desc=""):
    """Classify a downloaded PDF by its actual content.
    Returns (doc_type, quality, note):
      doc_type — 'manual' / 'datasheet' / 'scanned' / 'unknown' or a reject type
      quality  — higher = better; used to pick the best file per part
      note     — human-readable reason for the log
    src_url/src_title: where the file came from. Manufacturer manuals are often
    image-heavy and never mention the part number as extractable TEXT, but the
    official download link literally contains it (…/8108235_Instructions_EN).
    An exact whole-token match there counts as part-number confirmation.
    """
    n, text, density = _pdf_sample(path)
    if n == 0:
        # pdfplumber couldn't parse it. Before rejecting, fall back to pypdf —
        # some valid PDFs (vector-only, or with quirks pdfplumber dislikes) are
        # readable by pypdf. And if the file is a structurally-complete PDF that
        # came from the manufacturer's OWN domain with a doc-type filename hint,
        # trust it: the old build downloaded these official datasheets fine, and
        # the content-verifier should not be stricter than necessary for a
        # source we already know is authoritative.
        pages2 = _pypdf_pagecount(path)
        if pages2 > 0:
            n = pages2
        else:
            if (_is_own_mfr_domain(src_url, mfr) and name_hint in ("manual", "datasheet")
                    and _pdf_is_complete(path)):
                return name_hint, 3, "official manufacturer PDF (parser couldn't extract text — accepted on source)"
            return "unreadable", 0, "could not open / parse PDF"
        text, density = "", 0   # no extractable text, but we have a page count

    # Part-number presence: exact whole token, else significant chunks of it
    mm = _match_model(model, mfr)   # ABB: the type designation counts as identity
    url_match = bool((src_url or src_title) and
                     _model_in_text(f"{src_url} {src_title}", mm))
    if _model_in_text(text, mm):
        model_ok = 2
    else:
        chunks = re.findall(r'[A-Za-z]{2,}|\d{3,}', model)
        model_ok = 1 if any(_model_in_text(text, c) for c in chunks[:3]) else 0
        if model_ok == 0 and mfr and "HIMA" in mfr.upper():
            mu_model = model.upper().strip()
            for prefix, families in _HIMA_FAMILIES.items():
                if mu_model.startswith(prefix):
                    tl_full = text.lower()
                    if any(fam in tl_full for fam in families):
                        model_ok = 1
                        break

    # Image-only scan (no extractable text) — judge by filename + page count
    if len(text.strip()) < 60:
        if name_hint in _REJECT_TYPES:
            return name_hint, 0, "image-only scan, filename indicates non-manual"
        if not ACCEPT_SCANNED:
            return "unreadable", 0, f"image-only scan ({n}p) and ACCEPT_SCANNED is off"
        if name_hint == "manual":   # e.g. 'Instruction/Installation Manual' scans
            return "manual", 4, f"image-only scan, {n}p — filename indicates a manual"
        if name_hint == "datasheet":
            return "datasheet", 2, f"image-only scan, {n}p — filename indicates a datasheet"
        if n >= 5:
            return "scanned", 2, f"image-only scan, {n} pages — kept on page-count evidence"
        return "cad_drawing", 0, f"image-only, {n} page(s) — looks like a drawing/certificate"

    sc = {
        "manual":      _kw_score(text, _W_MANUAL),
        "datasheet":   _kw_score(text, _W_DS),
        "license":     _kw_score(text, _W_LIC),
        "certificate": _kw_score(text, _W_CERT),
        "cad_drawing": _kw_score(text, _W_CAD),
        "brochure":    _kw_score(text, _W_BROCH),
    }
    # Structural signals
    if n <= 3 and density < 220 and sc["manual"] < 3:
        sc["cad_drawing"] += 3              # short + sparse = drawing/cert layout
    if n >= 8:
        sc["manual"] += 2                   # licenses/certs/drawings are rarely 8+ pages
    if name_hint == "manual":    sc["manual"]    += 2
    if name_hint == "datasheet": sc["datasheet"] += 2

    dtype = max(sc, key=lambda k: sc[k])
    if sc[dtype] < 3:
        dtype = "unknown"

    # A document that never mentions the part (not even a family chunk) is wrong
    # — UNLESS the source link itself carries the exact part number (official
    # portals serve picture-heavy manuals whose text never spells it out).
    if model_ok == 0 and not url_match \
            and dtype not in ("license", "certificate", "cad_drawing", "brochure"):
        return "wrong_part", 0, f"{n}p — part number not found anywhere in document"

    # For ALL-NUMERIC part numbers (Rittal 8601020 etc.) a partial/fragment
    # match means a DIFFERENT variant (e.g. 8601.000 found instead of 8601.020)
    # — require the exact number (in the text OR in the source link).
    alnum = re.sub(r'[^A-Za-z0-9]', '', model)
    if (alnum.isdigit() and len(alnum) >= 5 and model_ok == 1 and not url_match
            and dtype not in ("license", "certificate", "cad_drawing", "brochure")):
        return "wrong_part", 0, (f"{n}p — only a fragment of the part number found; "
                                 f"this looks like a different variant")

    base = {"manual": 5, "datasheet": 3, "scanned": 2, "unknown": 1}
    q = base.get(dtype, 0)
    if model_ok == 2: q += 1
    if dtype == "manual" and n >= 10: q += 1
    # Description keywords found in the document add confidence (bonus only —
    # never used to reject, so a sparse desc can't cause false rejections)
    dkw = 0
    if desc:
        tl_full = text.lower()
        dkw = sum(1 for k in _desc_keywords(desc) if k in tl_full)
        if dkw >= 2: q += 1

    hits = {k: v for k, v in sc.items() if v}
    pm = ("exact" if model_ok == 2
          else "via source filename" if (model_ok < 2 and url_match)
          else "partial" if model_ok == 1 else "none")
    note = f"{n}p, signals={hits or 'none'}, part-match={pm}"
    if dkw >= 2:
        note += f", desc-match={dkw} keyword(s)"
    return dtype, q, note

def _ensure_unprotected(path):
    """Deal with encrypted PDFs right after download.
    Manufacturer PDFs are often 'permission-protected' (encrypted with an EMPTY
    user password) — these open fine but break pypdf merging and show as
    (SECURED). We silently rewrite those without encryption.
    Truly password-locked PDFs are rejected.
    Returns (ok, note)."""
    try:
        from pypdf import PdfReader, PdfWriter
        r = PdfReader(str(path))
        if not r.is_encrypted:
            return True, ""
        try:
            res = r.decrypt("")
            if int(res) == 0:                       # NOT_DECRYPTED
                return False, "password-protected PDF (locked)"
        except Exception as e:
            if "cryptography" in str(e).lower() or "AES" in str(e):
                return False, "AES-protected PDF — install the 'cryptography' package to handle these"
            return False, "password-protected PDF (locked)"
        w = PdfWriter()
        for pg in r.pages:
            w.add_page(pg)
        tmp = path.with_suffix(".unlocked.tmp")
        with open(tmp, "wb") as f:
            w.write(f)
        tmp.replace(path)
        return True, "protection removed (was permission-protected)"
    except Exception as e:
        return False, f"unreadable PDF ({type(e).__name__})"


def candidate_dest(folder, bom_idx, mfr, model, title):
    """Canonical output filename for a candidate — shared by the download loop
    and the GUI's re-search 'Keep this PDF' action so names always match."""
    label = f"{mfr} {model}".strip() if mfr else model
    safe = _sanitize(label.replace(" ", "_"))
    clean_title = re.sub(r'\.pdf\s*$', '', (title or "document").strip(), flags=re.I)
    return Path(folder) / f"{bom_idx:03d}_{safe}__{_sanitize(clean_title.replace(' ', '_'))}.pdf"


# =============================================================================
#  CATALOG  -  your own library of manuals, local OR online
#  Sources accepted (mix freely, one per line in the GUI field / repeated -C):
#    - a local folder            D:\\Manuals   (incl. OneDrive/GDrive sync folders)
#    - a OneDrive share link     https://1drv.ms/f/...  or  https://...sharepoint.com/...
#                                (folder shared as "Anyone with the link can view")
#    - a Google Drive folder     https://drive.google.com/drive/folders/<id>?...
#                                (shared as "Anyone with the link - Viewer")
#  Files are matched by PART NUMBER IN THE FILENAME (dots/dashes/case ignored),
#  e.g. "Rittal 8205.521 manual.pdf" or "2866789_QUINT-PS.pdf". Online sources
#  are listed once per run; only MATCHING files are downloaded.
# =============================================================================

import base64 as _b64

_CATALOG_SOURCES : list = []   # raw strings: paths and/or share links
_CATALOG_INDEX   = [None]      # [(compact_name, raw_name, ref)]; ref: Path | dict
_CATALOG_LOCK    = threading.Lock()

def set_catalog_dirs(sources):
    """Configure catalog sources (local folders and/or OneDrive / Google Drive
    share links). Pass [] / None to disable. Kept under its old name for
    compatibility with the GUI."""
    global _CATALOG_SOURCES
    with _CATALOG_LOCK:
        _CATALOG_SOURCES = [str(s).strip() for s in (sources or [])
                            if s and str(s).strip()]
        _CATALOG_INDEX[0] = None          # force re-index on next lookup
        _CATALOG_MAP[0]   = None

# Optional explicit mapping file placed in the catalog root:  catalog_map.txt
#   one entry per line:   <part number> = <filename or fragment>
#   e.g.    8205.521 = Rittal Modular Enclosure Handbook
#           2153.000 ; TS8 accessories guide.pdf       (= ; , all accepted)
_MAP_NAMES = ("catalog_map.txt", "catalog_map.csv")
_CATALOG_MAP  = [None]    # {compact_part: [compact_filename_fragment, ...]}
_CATALOG_TEXT = [None]    # deep-scan cache: {path: {"sig":..., "txt": compact}}
CATALOG_DEEP_SCAN   = True    # read INSIDE local catalog PDFs when names don't match
CATALOG_DEEP_SCAN_ONLINE = True   # also read inside ONLINE catalog PDFs: each file is
                                  # downloaded ONCE, its text cached forever after —
                                  # essential when catalog filenames are meaningless
_CATALOG_CACHE_FILE = Path.home() / ".bom_catalog_text_cache.json"

def _parse_catalog_map(text):
    mp = {}
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        for sep in ("=", ";", ","):
            if sep in line:
                k, v = line.split(sep, 1)
                break
        else:
            continue
        kc = re.sub(r'[^a-z0-9]', '', k.lower())
        vc = re.sub(r'[^a-z0-9]', '',
                    re.sub(r'\.pdf\s*$', '', v.strip(), flags=re.I).lower())
        if kc and vc:
            mp.setdefault(kc, []).append(vc)
    return mp

def _fetch_map_text(ref, session):
    try:
        if isinstance(ref, Path):
            return ref.read_text(encoding="utf-8", errors="ignore")
        if ref.get("kind") == "gdrive":
            r = session.get("https://drive.usercontent.google.com/download",
                            params={"id": ref["id"], "export": "download",
                                    "confirm": "t"},
                            timeout=30, headers=BROWSER_HDR)
            return r.text
        r = session.get(ref["url"], timeout=30)
        return r.text
    except Exception:
        return ""

# ---- OneDrive (anonymous share link, no sign-in required) -------------------

def _onedrive_share_id(link):
    """Encode a share URL the way Microsoft's anonymous /shares API expects."""
    b = _b64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + b

def _list_onedrive(link, session):
    """List every PDF in an 'Anyone with the link' OneDrive/SharePoint folder,
    recursing into subfolders. Returns [(name, download_ref_dict)]."""
    sid  = _onedrive_share_id(link)
    api  = "https://api.onedrive.com/v1.0"
    out  = []
    def _walk(children_url, depth=0):
        if depth > 6: return
        url = children_url
        while url:
            r = session.get(url, timeout=25,
                            headers={"Accept": "application/json"})
            r.raise_for_status()
            data = r.json()
            for item in data.get("value", []):
                name = item.get("name", "")
                if "folder" in item:
                    _walk(f"{api}/shares/{sid}/items/{item['id']}/children",
                          depth + 1)
                elif name.lower().endswith(".pdf") or name.lower() in _MAP_NAMES:
                    dl = (item.get("@content.downloadUrl")
                          or f"{api}/shares/{sid}/items/{item['id']}/content")
                    out.append((name, {"kind": "url", "url": dl,
                                       "key": f"od:{item.get('id','')}",
                                       "sig": f"{item.get('eTag','')}-{item.get('size','')}"}))
            url = data.get("@odata.nextLink")
    _walk(f"{api}/shares/{sid}/root/children")
    return out

# ---- Google Drive (public folder; API key optional) --------------------------

def _gdrive_folder_id(link):
    m = (re.search(r'/folders/([\w-]{10,})', link)
         or re.search(r'[?&]id=([\w-]{10,})', link))
    return m.group(1) if m else None

def _list_gdrive_api(folder_id, session, depth=0):
    """Listing via the official Drive v3 API (needs GDRIVE_API_KEY; folder must
    be shared 'Anyone with the link'). Handles pagination + subfolders."""
    if depth > 6: return []
    out, token = [], None
    while True:
        params = {"q": f"'{folder_id}' in parents and trashed=false",
                  "fields": "nextPageToken,files(id,name,mimeType,size,md5Checksum)",
                  "pageSize": 1000, "key": GDRIVE_API_KEY}
        if token: params["pageToken"] = token
        r = session.get("https://www.googleapis.com/drive/v3/files",
                        params=params, timeout=25)
        r.raise_for_status()
        data = r.json()
        for f in data.get("files", []):
            if f.get("mimeType") == "application/vnd.google-apps.folder":
                out += _list_gdrive_api(f["id"], session, depth + 1)
            elif (f.get("name", "").lower().endswith(".pdf")
                  or f.get("name", "").lower() in _MAP_NAMES):
                out.append((f["name"], {"kind": "gdrive", "id": f["id"],
                                        "key": f"gd:{f['id']}",
                                        "sig": f"{f.get('md5Checksum','')}-{f.get('size','')}"}))
        token = data.get("nextPageToken")
        if not token: return out

def _list_gdrive_scrape(folder_id, session, depth=0):
    """No-API-key fallback: parse the public embedded folder view. Fine for
    typical catalogs; for thousands of files set GDRIVE_API_KEY instead."""
    if depth > 6: return []
    r = session.get("https://drive.google.com/embeddedfolderview",
                    params={"id": folder_id}, timeout=25, headers=BROWSER_HDR)
    r.raise_for_status()
    out = []
    for m in re.finditer(
            r'<div class="flip-entry"[^>]*id="entry-([\w-]+)".*?'
            r'<div class="flip-entry-title">([^<]+)</div>',
            r.text, re.S):
        fid, name = m.group(1), m.group(2).strip()
        blk = m.group(0)
        if "/folders/" in blk:             # subfolder entry → recurse into it
            out += _list_gdrive_scrape(fid, session, depth + 1)
        elif name.lower().endswith(".pdf") or name.lower() in _MAP_NAMES:
            out.append((name, {"kind": "gdrive", "id": fid,
                               "key": f"gd:{fid}", "sig": ""}))
    return out

def _gdrive_download(file_id, dest, session):
    """Download a public Drive file, transparently handling the 'can't scan
    for viruses' interstitial that large files trigger."""
    url = "https://drive.usercontent.google.com/download"
    try:
        r = session.get(url, params={"id": file_id, "export": "download",
                                     "confirm": "t"},
                        timeout=60, stream=True, headers=BROWSER_HDR)
        head = next(r.iter_content(1024), b"")
        if head[:4] != b"%PDF" and b"<html" in head[:512].lower():
            # interstitial page -> parse the hidden form params and retry
            body = head + r.content
            r.close()
            params = dict(re.findall(
                r'<input type="hidden" name="(\w+)" value="([^"]*)"',
                body.decode("utf-8", "ignore")))
            params.setdefault("id", file_id)
            params.setdefault("export", "download")
            params.setdefault("confirm", "t")
            r = session.get(url, params=params, timeout=60, stream=True,
                            headers=BROWSER_HDR)
            head = next(r.iter_content(1024), b"")
        if head[:4] != b"%PDF":
            r.close(); return False
        with open(dest, "wb") as f:
            f.write(head)
            for chunk in r.iter_content(16384):
                if _is_cancelled(): r.close(); dest.unlink(missing_ok=True); return False
                f.write(chunk)
        r.close()
        return True
    except Exception:
        dest.unlink(missing_ok=True)
        return False

# ---- Unified index / lookup / fallback ---------------------------------------

def _catalog_index():
    """Build (once per run) a filename index of every catalog PDF — local
    folders are walked, online sources are listed over HTTPS. Thread-safe."""
    with _CATALOG_LOCK:
        if _CATALOG_INDEX[0] is None:
            idx, s = [], make_pooled_session()
            for src in _CATALOG_SOURCES:
                try:
                    if src.lower().startswith(("http://", "https://")):
                        sl = src.lower()
                        if "drive.google" in sl or "docs.google" in sl:
                            fid = _gdrive_folder_id(src)
                            if not fid:
                                tprint(f"    [Catalog] not a Drive FOLDER link: {src[:60]}")
                                continue
                            entries = (_list_gdrive_api(fid, s) if _has(GDRIVE_API_KEY)
                                       else _list_gdrive_scrape(fid, s))
                        else:                       # OneDrive / SharePoint share
                            entries = _list_onedrive(src, s)
                        for name, ref in entries:
                            idx.append((re.sub(r'[^a-z0-9]', '', name.lower()),
                                        name.lower(), ref))
                        tprint(f"    [Catalog] {len(entries)} PDF(s) listed from "
                               f"online catalog ({_dom(src)})")
                    else:
                        root = Path(src)
                        if not root.is_dir():
                            tprint(f"    [Catalog] folder not found: {src}")
                            continue
                        n0 = len(idx)
                        for p in root.rglob("*.pdf"):
                            name = p.name.lower()
                            idx.append((re.sub(r'[^a-z0-9]', '', name), name, p))
                        tprint(f"    [Catalog] indexed {len(idx)-n0} PDF(s) from {root}")
                except Exception as e:
                    tprint(f"    [Catalog] {_yellow('cannot read source')} "
                           f"{src[:60]}: {str(e)[:80]}")
            # split out catalog_map files, parse them, drop them from the index
            mp, pdfs = {}, []
            for compact, name, ref in idx:
                if name in _MAP_NAMES:
                    for k, v in _parse_catalog_map(_fetch_map_text(ref, s)).items():
                        mp.setdefault(k, []).extend(v)
                else:
                    pdfs.append((compact, name, ref))
            # local roots: map files aren't caught by the *.pdf walk — check now
            for srcdir in _CATALOG_SOURCES:
                if srcdir.lower().startswith(("http://", "https://")):
                    continue
                root = Path(srcdir)
                if root.is_dir():
                    for mn in _MAP_NAMES:
                        for mpf in root.rglob(mn):
                            for k, v in _parse_catalog_map(
                                    _fetch_map_text(mpf, s)).items():
                                mp.setdefault(k, []).extend(v)
            if mp:
                tprint(f"    [Catalog] mapping file: {len(mp)} part(s) pinned")
            _CATALOG_MAP[0]   = mp
            _CATALOG_INDEX[0] = pdfs
        return _CATALOG_INDEX[0]

def _desig_tokens(desc):
    """Distinctive type-designation tokens from a BOM description, compacted.
    'QUINT-PS/1AC/24DC/40' -> ['quintps1ac24dc40', 'quintps', 'quint'] so a
    file named 'QUINT-PS user guide.pdf' still matches."""
    toks = []
    for w in re.findall(r'\S+', desc or ""):
        if not re.search(r'[A-Za-z]', w):
            continue
        if not (re.search(r'\d', w) or "-" in w or "/" in w):
            continue
        segs = [sg for sg in re.split(r'[-/().,]+', w) if sg]
        for k in range(len(segs), 0, -1):           # longest first
            t = re.sub(r'[^a-z0-9]', '', "".join(segs[:k]).lower())
            if len(t) >= 5 and t not in toks:
                toks.append(t)
    return toks[:12]

def _extract_compact_text(path, pages=5):
    """First-pages text of a PDF, compacted for part-number search."""
    txt = ""
    try:
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages[:pages]:
                txt += (pg.extract_text() or "") + " "
    except Exception:
        pass
    return re.sub(r'[^a-z0-9]', '', txt.lower())[:200000]

def _fetch_catalog_pdf(ref, dest, session):
    """Plain download for catalog files (OneDrive tempauth / Drive). Unlike
    _write_pdf there is no minimum-size floor — your own catalog may contain
    legitimately tiny one-page PDFs — only the %PDF magic is enforced."""
    if ref.get("kind") == "gdrive":
        return _gdrive_download(ref["id"], dest, session)
    try:
        r = session.get(ref["url"], timeout=60, stream=True, headers=BROWSER_HDR)
        r.raise_for_status()
        head = next(r.iter_content(1024), b"")
        if head[:4] != b"%PDF":
            r.close(); return False
        with open(dest, "wb") as f:
            f.write(head)
            for chunk in r.iter_content(16384):
                if _is_cancelled():
                    r.close(); dest.unlink(missing_ok=True); return False
                f.write(chunk)
        r.close()
        return True
    except Exception:
        dest.unlink(missing_ok=True)
        return False

def _catalog_deep_scan(lond, session=None):
    """Search the TEXT of catalog PDFs for the part number (first pages only).

    Local files: read in place. Online files: downloaded ONCE to a temp file,
    their text cached on disk forever after (keyed by the file's cloud id +
    eTag/md5), so randomly-named cloud catalogs work from the second lookup
    at zero network cost. The first scan of a big online catalog downloads
    every file — it logs that clearly and never repeats it."""
    if not CATALOG_DEEP_SCAN or "pdfplumber" not in globals():
        return []
    import json, hashlib
    hits = []
    s = session or make_pooled_session()
    with _CATALOG_LOCK:
        if _CATALOG_TEXT[0] is None:
            try:
                _CATALOG_TEXT[0] = json.loads(_CATALOG_CACHE_FILE.read_text())
            except Exception:
                _CATALOG_TEXT[0] = {}
        cache, dirty = _CATALOG_TEXT[0], False

        # announce a first-time online scan before it starts
        if CATALOG_DEEP_SCAN_ONLINE:
            pending = [name for _c, name, ref in (_CATALOG_INDEX[0] or [])
                       if isinstance(ref, dict)
                       and (ref.get("key") or "u:" + hashlib.sha1(
                            str(ref).encode()).hexdigest()) not in cache]
            if pending:
                tprint(f"    [Catalog] one-time scan of {len(pending)} online "
                       f"PDF(s) to read part numbers inside them — cached "
                       f"after this, later runs are instant")

        for compact, name, ref in (_CATALOG_INDEX[0] or []):
            if _is_cancelled():
                break
            if isinstance(ref, Path):
                try:
                    st  = ref.stat()
                    key, sig = str(ref), f"{st.st_mtime_ns}-{st.st_size}"
                except Exception:
                    continue
                ent = cache.get(key)
                if not ent or ent.get("sig") != sig:
                    ent = {"sig": sig, "name": name,
                           "txt": _extract_compact_text(ref)}
                    cache[key] = ent
                    dirty = True
            else:
                if not CATALOG_DEEP_SCAN_ONLINE:
                    continue
                key = ref.get("key") or "u:" + hashlib.sha1(
                    str(ref).encode()).hexdigest()
                sig = ref.get("sig", "")
                ent = cache.get(key)
                if not ent or (sig and ent.get("sig") != sig):
                    tmp = Path(tempfile.gettempdir()) / f"bomcat_{hashlib.sha1(key.encode()).hexdigest()}.pdf"
                    ok = _fetch_catalog_pdf(ref, tmp, s)
                    txt = _extract_compact_text(tmp) if ok else ""
                    try: tmp.unlink()
                    except Exception: pass
                    if not ok:
                        tprint(f"    [Catalog] could not fetch {name} for scanning — will retry next run")
                        continue          # don't cache failures
                    ent = {"sig": sig, "name": name, "txt": txt}
                    cache[key] = ent
                    dirty = True
            if lond in ent["txt"]:
                hits.append((name, ref))
        if dirty:
            try:
                _CATALOG_CACHE_FILE.write_text(json.dumps(cache))
            except Exception:
                pass
    return hits

# Canonical manufacturer tokens (compacted) used to spot cross-manufacturer
# catalog mismatches in the fuzzy tier. Short/ambiguous tokens (abb, mtl, gmi…)
# are deliberately excluded from FOREIGN detection so they can't match inside
# unrelated words; a part's OWN manufacturer is still checked separately.
_CATALOG_MFR_TOKENS = [t for t in sorted(
    {re.sub(r'[^a-z0-9]', '', c.lower()) for _n, c in _KNOWN_MFRS}
    | {"allenbradley", "rockwell", "schneider", "phoenixcontact", "weidmuller",
       "honeywell", "siemens", "eaton", "prosoft", "tracopower", "rosemount"},
    key=len, reverse=True) if len(t) >= 4]

def _catalog_foreign_mfr(compact_name, own_full):
    """True when a catalog filename clearly belongs to a DIFFERENT known
    manufacturer than this part's (and not the part's own). Used to stop a
    fuzzy keyword match from returning, say, an ABB 'digital input' datasheet
    for an MTL part."""
    if own_full and own_full in compact_name:
        return False
    return any(t in compact_name and t != own_full for t in _CATALOG_MFR_TOKENS)

def _catalog_lookup(model, mfr, desc="", session=None, confident_only=False):
    """Find catalog entries for a part, in confidence order. Each tier runs
    only if the previous found nothing:
      0. catalog_map.txt pin            (explicit — you said so)          CONFIDENT
      1. part number in the FILENAME    (dots/dashes/case ignored)        CONFIDENT
      2. part number INSIDE the PDF     (local files, cached text scan)   CONFIDENT
      3. type designation / description keywords vs the filename          FUZZY
    Tiers 0-2 pin the part with certainty; tier 3 is only a description guess.
    With confident_only=True the fuzzy tier is skipped — callers use that to
    decide whether a catalog hit is trustworthy enough to SKIP the web search.
    Returns [(display_name, ref)]."""
    if not _CATALOG_SOURCES:
        return []
    idx = _catalog_index()
    c, nd, dot, lo, lond = _norm(model)
    if len(lond) < 4:                      # too short — would match everything
        return []
    mtok = (mfr or "").lower().split()[0] if mfr else ""
    own_full = re.sub(r'[^a-z0-9]', '', (mfr or "").lower())

    # Tier 0 — explicit mapping file
    for frag in (_CATALOG_MAP[0] or {}).get(lond, []):
        hits = [(name, ref) for compact, name, ref in idx if frag in compact]
        if hits:
            tprint(f"    [Catalog] {model}: matched via catalog_map entry")
            return hits[:6]

    # Tier 1 — part number in the filename
    hits = []
    for compact, name, ref in idx:
        if lond in compact:
            mfr_bonus    = 1 if (mtok and len(mtok) >= 3 and mtok in name) else 0
            manual_bonus = 1 if _name_doc_type(name, name) == "manual" else 0
            hits.append((mfr_bonus + manual_bonus, name, ref))
    if hits:
        hits.sort(key=lambda t: t[0], reverse=True)
        return [(name, ref) for _s, name, ref in hits[:6]]

    # Tier 2 — part number inside the PDF text (local files only)
    hits = _catalog_deep_scan(lond, session)
    if hits:
        tprint(f"    [Catalog] {model}: found inside PDF text "
               f"({len(hits)} file(s), filename didn't contain the number)")
        return hits[:6]

    # Everything below is a FUZZY guess (no part-number evidence). When the
    # caller only trusts confident hits, stop here so the web search can run.
    if confident_only:
        return []

    # Tier 3 — type designation / description keywords vs filename
    toks = _desig_tokens(desc)
    kws  = [k.lower() for k in _desc_keywords(desc) if len(k) >= 4]
    hits = []
    for compact, name, ref in idx:
        # Never let a fuzzy keyword match return a file that clearly belongs to
        # a DIFFERENT manufacturer (the MTL-5544 → ABB-DI818 bug).
        if _catalog_foreign_mfr(compact, own_full):
            continue
        strong = next((t for t in toks if t in compact), None)
        kwhits = sum(1 for k in kws if k in name)
        if strong or kwhits >= 2:
            mfr_bonus = 1 if (mtok and len(mtok) >= 3 and mtok in name) else 0
            hits.append(((2 if strong else 0) + kwhits + mfr_bonus, name, ref))
    if hits:
        hits.sort(key=lambda t: t[0], reverse=True)
        tprint(f"    [Catalog] {model}: matched by type designation/description "
               f"({_yellow('verify the result')} — part number not in name or text)")
        return [(name, ref) for _s, name, ref in hits[:3]]
    return []

def _catalog_fallback(bom_idx, mfr, model, folder, desc, label, max_dl,
                      session=None, confident_only=False):
    """Fetch up to max_dl matching catalog PDFs into the output folder and run
    them through the normal verification. Because the catalog is user-curated,
    a file that FAILS strict verification is still kept (clearly labelled) —
    family-level docs whose text never mentions the article number are exactly
    what people put in their own catalogs.
    confident_only=True restricts matching to exact part-number / map hits, so
    a fuzzy description guess never short-circuits the online search."""
    saved = []
    s = session or make_pooled_session()
    for name, ref in _catalog_lookup(model, mfr, desc, s, confident_only=confident_only):
        if len(saved) >= max_dl or _is_cancelled():
            break
        stem = re.sub(r'\.pdf$', '', name, flags=re.I)
        dest = candidate_dest(folder, bom_idx, mfr, model, "CATALOG_" + stem)
        ok = False
        try:
            if dest.exists():
                ok = True
            elif isinstance(ref, Path):
                shutil.copy2(ref, dest)    # OneDrive placeholder auto-hydrates
                ok = True
            else:                          # online ref (OneDrive / Drive)
                ok = _fetch_catalog_pdf(ref, dest, s)
                if ok:
                    kb = dest.stat().st_size // 1024
                    tprint(f"         {_green('Saved:')} {dest.name}  ({kb} KB, from catalog)")
        except Exception as e:
            tprint(f"    [{label}] [Catalog] fetch failed for {name}: {str(e)[:80]}")
        if not ok:
            tprint(f"    [{label}] [Catalog] could not fetch {name}")
            continue
        ok_prot, prot_note = _ensure_unprotected(dest)
        if not ok_prot:
            tprint(f"    [{label}] [Catalog] {name}: {prot_note} — skipped")
            try: dest.unlink()
            except Exception: pass
            continue
        dtype, q, vnote = "unverified", 1, "content verification off"
        if VERIFY_CONTENT:
            dtype, q, vnote = _verify_pdf(dest, model, mfr,
                                          _name_doc_type(name, stem),
                                          src_url=name, src_title=stem,
                                          desc=desc)
        if dtype in _REJECT_TYPES:
            note = f"kept despite verifying as {_REJECT_LABEL.get(dtype, dtype)} (user catalog)"
            dtype = "catalog"
        else:
            note = vnote
        tprint(f"    [{label}] {_green('Catalog hit:')} {name}  →  "
               f"{dtype.upper()}  [{note}]")
        saved.append(dest)
    return saved


def _download_part(bom_idx, mfr, model, folder, max_dl, session, desc="", force=False):
    """Wrapper that tags this worker thread with the part it handles, so the
    GUI's per-part skip can abort just this part via _is_cancelled()."""
    _current_part_tl.idx = bom_idx
    try:
        return _download_part_impl(bom_idx, mfr, model, folder, max_dl,
                                   session, desc, force)
    finally:
        _current_part_tl.idx = None

def _download_part_impl(bom_idx, mfr, model, folder, max_dl, session, desc="", force=False):
    """Returns {"saved":[Path,...], "candidates":[{url,title,referer},...]}
    Every saved file is content-verified (when VERIFY_CONTENT is on).
    Speed-bounded two-phase strategy:
      Phase 1 — walk candidates in rank order, stop as soon as max_dl docs
                are accepted (or a manual is found).
      Phase 2 — if we only got datasheets, spend at most MANUAL_HUNT_EXTRA
                additional downloads, and ONLY on candidates whose filename/
                title explicitly says 'manual'. Everything else is skipped."""
    if _is_cancelled():
        return {"saved": [], "candidates": []}
    label = f"{mfr} {model}".strip() if mfr else model
    prefix = f"{bom_idx:03d}"
    _emit("part_start", idx=bom_idx, label=label)

    # ── 1) CONFIDENT catalog hit first — exact part number in the filename, the
    #       part number found INSIDE a catalog PDF, or an explicit catalog_map
    #       pin. These identify the part with certainty, so use them and skip
    #       the web entirely (this is the catalog-first behaviour you want).
    cat = _catalog_fallback(bom_idx, mfr, model, folder, desc, label, max_dl,
                            session, confident_only=True)
    if cat:
        tprint(f"    [{label}] {_green('Found in catalog')} — skipping web search")
        _emit("part_done", idx=bom_idx, label=label, status="found",
              files=[str(f) for f in cat])
        return {"saved": cat, "candidates": []}

    # ── 2) Online search ──────────────────────────────────────────────────────
    candidates = _find_pdfs(model, mfr, session, desc, force=force)
    if not candidates:
        # 3) Last resort ONLY: a fuzzy catalog match (description/keyword guess,
        #    same-manufacturer only). Clearly flagged, and never allowed to
        #    pre-empt the online search — that is what made MTL parts grab ABB
        #    datasheets.
        cat = _catalog_fallback(bom_idx, mfr, model, folder, desc, label, max_dl, session)
        if cat:
            tprint(f"    [{label}] {_yellow('Using closest catalog match')} "
                   f"(no online result — please verify)")
            _emit("part_done", idx=bom_idx, label=label, status="found",
                  files=[str(f) for f in cat])
            return {"saved": cat, "candidates": []}
        tprint(f"    [{label}] {_red('No URLs found')}")
        _emit("part_done", idx=bom_idx, label=label, status="not_found", files=[])
        return {"saved": [], "candidates": []}

    safe = _sanitize(label.replace(" ", "_"))
    accepted = []          # list of (quality, Path, doc_type)
    attempts = 0           # network downloads actually performed
    have_manual = False

    def _try_candidate(item, pos):
        """Download + verify one candidate. Returns (q, path, dtype) or None."""
        nonlocal attempts
        if _is_cancelled():
            return None
        url, title, ref = item["url"], item["title"], item.get("referer")
        dest = candidate_dest(folder, bom_idx, mfr, model, title)
        name_hint = _name_doc_type(url, title)
        if any(p == dest for _q, p, _t in accepted):
            return None                      # same filename already accepted
        if dest.exists():
            tprint(f"    [{label}] Re-checking existing: {dest.name}")
        else:
            tprint(f"    [{label}] [{pos}/{len(candidates)}] {title[:60]}")
            tprint(f"         {url[:80]}")
            attempts += 1
            if not _get_cached_or_download(url, dest, session, ref):
                return None
        # Encrypted PDFs: strip harmless permission-protection, reject locked ones
        ok_prot, prot_note = _ensure_unprotected(dest)
        if not ok_prot:
            tprint(f"         {_red('Rejected:')} {prot_note}")
            try: dest.unlink()
            except Exception: pass
            return None
        if prot_note:
            tprint(f"         {prot_note}")
        if not VERIFY_CONTENT:
            return (1, dest, "unverified")
        dtype, q, note = _verify_pdf(dest, model, mfr, name_hint,
                                     src_url=url, src_title=title, desc=desc)
        if dtype in _REJECT_TYPES or (dtype == "datasheet" and not ACCEPT_DATASHEETS):
            tprint(f"         {_red('Rejected:')} {_REJECT_LABEL.get(dtype, dtype)}  [{note}]")
            try: dest.unlink()
            except Exception: pass
            return None
        tprint(f"         {_green('Verified:')} {dtype.upper()}  [{note}]")
        return (q, dest, dtype)

    # ── Phase 1: fill the quota with SOLID docs, stop early ──────────────────
    # 'unknown' verifications are provisional: kept as a last-resort fallback,
    # but they do NOT fill the quota — the walk continues looking for a real
    # manual/datasheet (still bounded by VERIFY_MAX_ATTEMPTS).
    SOLID = ("manual", "datasheet", "scanned", "unverified")
    def _solid_count():
        return sum(1 for _q, _p, t in accepted if t in SOLID)

    manual_named_later = []      # (pos, item) — saved for the upgrade hunt
    for pos, item in enumerate(candidates, 1):
        if _is_cancelled() or attempts >= VERIFY_MAX_ATTEMPTS:
            break
        if _solid_count() >= max_dl:
            # Quota full — don't download; just remember manual-named candidates
            if PREFER_MANUALS and not have_manual and \
               _name_doc_type(item["url"], item["title"]) == "manual":
                manual_named_later.append((pos, item))
            continue
        res = _try_candidate(item, pos)
        if res:
            accepted.append(res)
            if res[2] == "manual":
                have_manual = True
                if _solid_count() >= max_dl:
                    break                    # got the manual(s) — done immediately

    # ── Phase 2: bounded upgrade hunt (datasheet → manual) ───────────────────
    if PREFER_MANUALS and accepted and not have_manual:
        for pos, item in manual_named_later[:MANUAL_HUNT_EXTRA]:
            if _is_cancelled() or attempts >= VERIFY_MAX_ATTEMPTS:
                break
            res = _try_candidate(item, pos)
            if res:
                accepted.append(res)
                if res[2] == "manual":
                    break                    # upgrade succeeded

    # Keep the best max_dl files (manual > solid datasheet/scan > unknown)
    accepted.sort(key=lambda t: (t[2] == "manual", t[2] in SOLID, t[0]), reverse=True)
    keep, surplus = accepted[:max_dl], accepted[max_dl:]
    for _q, p, _t in surplus:
        try: p.unlink()
        except Exception: pass
    saved = [p for _q, p, _t in keep]

    if saved and VERIFY_CONTENT:
        kinds = ", ".join(t.upper() for _q, _p, t in keep)
        tprint(f"    [{label}] {_green('Kept:')} {kinds}  ({attempts} download(s))")
    elif not saved:
        # Catalog was already tried first, so there's nothing left to fall back
        # to here — just report the web result.
        tprint(f"    [{label}] {_yellow('No verified manual/datasheet')} "
               f"— {len(candidates)} URL(s) checked")
    _emit("part_done", idx=bom_idx, label=label,
          status="found" if saved else "not_found",
          files=[str(f) for f in saved])
    return {"saved": saved, "candidates": candidates}

# =============================================================================
#  PDF PAGE BUILDERS
# =============================================================================

def _make_placeholder(idx, mfr, model, tmp):
    p=Path(tmp)/f"ph_{idx:03d}.pdf"; label=f"{mfr} {model}".strip() if mfr else model
    W,H=A4; c=rl_canvas.Canvas(str(p),pagesize=A4)
    c.setFillColorRGB(0.93,0.93,0.93); c.rect(0,H-80,W,80,fill=1,stroke=0)
    c.setFillColorRGB(0.7,0.0,0.0); c.setFont("Helvetica-Bold",11); c.drawCentredString(W/2,H-28,"DOCUMENT NOT FOUND")
    c.setFillColorRGB(0.2,0.2,0.2); c.roundRect(W/2-30,H-66,60,28,5,fill=1,stroke=0)
    c.setFillColorRGB(1,1,1); c.setFont("Helvetica-Bold",14); c.drawCentredString(W/2,H-56,f"#{idx:03d}")
    c.setFillColorRGB(0.1,0.1,0.1); c.setFont("Helvetica-Bold",20); c.drawCentredString(W/2,H/2+30,label)
    c.setStrokeColorRGB(0.7,0.7,0.7); c.setLineWidth(0.8); c.line(60,H/2+10,W-60,H/2+10)
    c.setFillColorRGB(0.45,0.45,0.45); c.setFont("Helvetica",11)
    c.drawCentredString(W/2,H/2-15,"No datasheet or manual could be located.")
    c.drawCentredString(W/2,H/2-35,"Please source this document manually.")
    c.setFont("Helvetica",8); c.setFillColorRGB(0.6,0.6,0.6); c.drawCentredString(W/2,30,"BOM Downloader")
    c.save(); return p

# =============================================================================
#  INTECH MANUFACTURER'S RECORD BOOK  —  cover (Table of Contents) + page breakers
#  Reproduces the exact reference format: black "TABLE OF CONTENTS" bar, gray
#  column header, ruled item grid with clickable "Go to Datasheet" links, and a
#  per-manual breaker page (Chapter # block + Manufacturer/Part # table +
#  "Back to Index" link). Footers read "Page X of Y" across the whole book.
# =============================================================================
import base64 as _base64

# -- Embedded logo assets (decoded to temp PNGs at merge time; PyInstaller-safe).
_LOGO_TOC_B64 = (
    "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"
    "YII="
)
_LOGO_BREAKER_B64 = (
    "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"
    "SAAAAABJRU5ErkJggg=="
)


def _write_logo(tmp, tag, b64):
    p = Path(tmp) / f"_logo_{tag}.png"
    try:
        p.write_bytes(_base64.b64decode(b64))
        return str(p)
    except Exception:
        return None


# -- Page geometry (A4, top-origin measurements converted to reportlab baseline).
_A4_W, _A4_H = A4                         # 595.32 x 841.92
_RB_DESCENT = 0.21


def _yb(bottom, size):
    """reportlab baseline-y from a pdfplumber 'bottom' (top-origin) coordinate."""
    return _A4_H - bottom + _RB_DESCENT * size


def _rect_td(c, x0, top, x1, bottom, fill_rgb=None, stroke_rgb=None, lw=0.6):
    """Draw a rectangle given top-origin top/bottom coordinates."""
    if fill_rgb is not None:
        c.setFillColorRGB(*fill_rgb)
    if stroke_rgb is not None:
        c.setStrokeColorRGB(*stroke_rgb)
        c.setLineWidth(lw)
    c.rect(x0, _A4_H - bottom, x1 - x0, bottom - top,
           fill=1 if fill_rgb is not None else 0,
           stroke=1 if stroke_rgb is not None else 0)


_BLUE = (0.0, 0.0, 1.0)
_BLACK = (0.0, 0.0, 0.0)
_GRAY_CHAP = (0.749, 0.749, 0.749)
_GRAY_HDR = (0.753, 0.753, 0.753)

# Clickable "Back to Index" target box on every breaker page (PDF / bottom-origin).
_BACK_TO_INDEX_RECT = (481.3, _A4_H - 150.5, 521.6, _A4_H - 142.5)


def _draw_footer(c, page_x, page_y):
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica", 8.0)
    c.drawRightString(523.3, _yb(815.6, 8.0), f"Page {page_x} of {page_y}")


# ---------------------------------------------------------------------------
#  BREAKER PAGE  (one before every manual)
# ---------------------------------------------------------------------------
def _draw_breaker_page(c, chapter_no, mfr, part, desc, page_x, page_y, logo_path):
    desc = " ".join((desc or "").split()).upper()
    mfr = (mfr or "").strip()
    part = (part or "").strip()

    # Header title (bold) + logo top-right + horizontal rule
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica-Bold", 9.2)
    c.drawString(61.0, _yb(93.6, 9.2), "MANUFACTURER'S RECORD BOOK")
    if logo_path:
        try:
            c.drawImage(logo_path, 402.1, _A4_H - 103.2, width=115.2, height=37.8,
                        mask='auto', preserveAspectRatio=True, anchor='ne')
        except Exception:
            pass
    _rect_td(c, 50.9, 107.2, 522.7, 107.8, fill_rgb=_BLACK)

    # "Back to Index" (blue, right-aligned, underlined)
    c.setFillColorRGB(*_BLUE)
    c.setFont("Helvetica", 6.6)
    c.drawRightString(521.6, _yb(149.8, 6.6), "Back to Index")
    _rect_td(c, 481.3, 149.2, 521.5, 149.6, fill_rgb=_BLUE)

    # Chapter block (gray fill + black border)
    _rect_td(c, 109.2, 152.6, 482.3, 194.9, fill_rgb=_GRAY_CHAP)
    _rect_td(c, 108.6, 152.0, 482.9, 195.5, stroke_rgb=_BLACK, lw=1.3)
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica", 11.9)
    c.drawString(120.0, _yb(179.9, 11.9), "Chapter #")
    c.drawString(184.2, _yb(180.3, 11.9), str(chapter_no))

    # Description: centered in the right sub-region, up to 3 lines @10pt
    lines = simpleSplit(desc, "Helvetica", 10.0, 268) or [""]
    if len(lines) > 3:
        lines = lines[:3]
        lines[-1] = lines[-1].rstrip() + " ..."
    c.setFont("Helvetica", 10.0)
    bottoms = [166.6, 178.6, 190.6]
    start = 3 - len(lines)
    for i, ln in enumerate(lines):
        c.drawCentredString(345.7, _yb(bottoms[start + i], 10.0), ln)

    # Manufacturer / Part # table
    t_l, t_r, t_div = 109.0, 482.5, 268.0
    t_top, t_mid, t_bot = 219.5, 232.0, 245.0
    for (a, b) in [(t_top, t_top + 0.6), (t_mid, t_mid + 0.6), (t_bot - 0.6, t_bot)]:
        _rect_td(c, t_l, a, t_r, b, fill_rgb=_BLACK)
    for x in [t_l, t_div, t_r]:
        _rect_td(c, x, t_top, x + 0.6, t_bot, fill_rgb=_BLACK)
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica-Bold", 7.9)
    c.drawString(110.9, _yb(230.0, 7.9), "Manufacturer")
    c.drawString(110.9, _yb(242.5, 7.9), "Part #")
    c.setFont("Helvetica", 7.9)
    c.drawString(269.6, _yb(230.0, 7.9), mfr)
    c.drawString(269.6, _yb(242.5, 7.9), part)

    _draw_footer(c, page_x, page_y)


# ---------------------------------------------------------------------------
#  COVER / TABLE OF CONTENTS
# ---------------------------------------------------------------------------
_COLS = [76.1, 120.2, 293.2, 362.5, 441.2, 519.9]          # column x-boundaries
_COL_CENTERS = [(_COLS[i] + _COLS[i + 1]) / 2 for i in range(5)]
_HDR_LABELS = ["Chapter #", "Item Description", "Manufacturer", "Part No.", "Attachement"]
_ROW_H = 7.05
_BODY_TOP = 116.4
_BODY_BOTTOM_LIMIT = 806.0


def _toc_header(c, page_x, page_y, logo_path):
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica-Bold", 9.4)
    c.drawString(77.5, _yb(82.1, 9.4), "MANUFACTURER'S RECORD BOOK")
    if logo_path:
        try:
            c.drawImage(logo_path, 423.0, _A4_H - 88.0, width=97.0, height=38.0,
                        mask='auto', preserveAspectRatio=True, anchor='ne')
        except Exception:
            pass
    # Black title bar + white centered title
    _rect_td(c, 76.6, 94.7, 520.0, 104.0, fill_rgb=_BLACK)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica", 8.0)
    c.drawCentredString((76.6 + 520.0) / 2, _yb(103.3, 8.0), "TABLE OF CONTENTS")
    # Gray column-header row + black bold centered labels
    _rect_td(c, 76.6, 108.0, 520.0, 116.4, fill_rgb=_GRAY_HDR)
    c.setFillColorRGB(*_BLACK)
    c.setFont("Helvetica-Bold", 6.8)
    for cx, lab in zip(_COL_CENTERS, _HDR_LABELS):
        c.drawCentredString(cx, _yb(116.4, 6.8), lab)
    _draw_footer(c, page_x, page_y)


def _toc_paginate(rows):
    """Greedily pack TOC rows into pages. Returns a list of pages; each page is a
    list of (row, row_top, row_height, desc_lines)."""
    pages = []
    cur = []
    y = _BODY_TOP
    for row in rows:
        desc = " ".join((row["desc"] or "").split()).upper()
        dlines = simpleSplit(desc, "Helvetica", 5.8, _COLS[2] - _COLS[1] - 3) or [""]
        rh = _ROW_H * len(dlines)
        if cur and (y + rh) > _BODY_BOTTOM_LIMIT:
            pages.append(cur)
            cur = []
            y = _BODY_TOP
        cur.append((row, y, rh, dlines))
        y += rh
    if cur:
        pages.append(cur)
    return pages


def _toc_draw_page(c, entries, page_x, page_y, logo_path):
    """Draw one cover page. Returns [(chapter_no, link_rect_pdf), ...] for the
    'Go to Datasheet' cells on this page."""
    _toc_header(c, page_x, page_y, logo_path)
    links = []
    for (row, row_top, rh, dlines) in entries:
        row_bot = row_top + rh
        # grid: column verticals + bottom rule
        c.setFillColorRGB(*_BLACK)
        for x in _COLS:
            _rect_td(c, x, row_top, x + 0.5, row_bot, fill_rgb=_BLACK)
        _rect_td(c, _COLS[0], row_bot - 0.4, _COLS[-1], row_bot, fill_rgb=_BLACK)
        # text
        c.setFillColorRGB(*_BLACK)
        c.setFont("Helvetica", 5.8)
        vcen = row_top + rh / 2.0
        c.drawCentredString(_COL_CENTERS[0], _yb(vcen + 2.0, 5.8), str(row["chapter_no"]))
        ly = row_top + _ROW_H
        for ln in dlines:
            c.drawString(_COLS[1] + 1.3, _yb(ly - 1.7, 5.8), ln)
            ly += _ROW_H
        c.drawCentredString(_COL_CENTERS[2], _yb(vcen + 2.0, 5.8), (row["mfr"] or "").upper())
        c.drawCentredString(_COL_CENTERS[3], _yb(vcen + 2.0, 5.8), row["part"] or "")
        # "Go to Datasheet" (blue, centered, underlined) + clickable rect
        c.setFillColorRGB(*_BLUE)
        link = "Go to Datasheet"
        c.drawCentredString(_COL_CENTERS[4], _yb(vcen + 2.0, 5.8), link)
        lw_ = stringWidth(link, "Helvetica", 5.8)
        ux0 = _COL_CENTERS[4] - lw_ / 2.0
        _rect_td(c, ux0, vcen + 3.0, ux0 + lw_, vcen + 3.4, fill_rgb=_BLUE)
        rect = (ux0 - 1.0, _A4_H - (vcen + 4.5), ux0 + lw_ + 1.0, _A4_H - (vcen - 4.0))
        links.append((row["chapter_no"], rect))
    # top border of the table on this page
    _rect_td(c, _COLS[0], _BODY_TOP, _COLS[-1], _BODY_TOP + 0.5, fill_rgb=_BLACK)
    return links


def _file_md5(path):
    import hashlib
    h = hashlib.md5()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


# =============================================================================
#  INTERACTIVE RESOLUTION
# =============================================================================

def _scan_folder(folder, idx):
    return sorted(folder.glob(f"{idx:03d}_*.pdf"))

def interactive_resolve(parts, part_results, folder):
    W = 64
    found_idx   = [i for i,r in part_results.items() if r["saved"]]
    missing_idx = [i for i,r in part_results.items() if not r["saved"]]

    print(f"\n{_bold('='*W)}")
    print(_bold("  DOWNLOAD SUMMARY"))
    print(_bold("='*W"))

    print(f"\n  {_green(_bold(f'FOUND  ({len(found_idx)} parts)'))}")
    print("  " + "-"*58)
    for idx in found_idx:
        p=parts[idx-1]; label=f"{p['manufacturer']} {p['part_number']}".strip()
        fname=part_results[idx]["saved"][0].name
        print(f"  {_dim(f'{idx:03d}')}  {label:<38}  {_dim(fname)}")

    print(f"\n  {_red(_bold(f'NOT FOUND  ({len(missing_idx)} parts)'))}")
    print("  " + "-"*58)
    for idx in missing_idx:
        p=parts[idx-1]; label=f"{p['manufacturer']} {p['part_number']}".strip()
        cands=part_results[idx]["candidates"]; efn=_expected_filename(idx,p["manufacturer"],p["part_number"])
        print(f"\n  {_yellow(_bold(f'  {idx:03d}  {label}'))}")
        print(f"       {_dim('Save as  :')}  {_cyan(efn)}")
        print(f"       {_dim('In folder:')}  {_dim(str(folder.resolve()))}")
        pdf_links=[c for c in cands if _is_pdf_url(c["url"])][:3]
        page_links=[c for c in cands if not _is_pdf_url(c["url"])][:2]
        if pdf_links:
            print(f"       {_green('Direct PDF URL(s):')}")
            for c in pdf_links: print(f"         -> {c['url']}")
        if page_links:
            print(f"       {_yellow('Page link(s):')}")
            for c in page_links: print(f"         -> {c['url']}")
        if not cands: print(f"       {_dim('No URL found — search manually.')}")

    if not missing_idx:
        print(f"\n  {_green('All parts downloaded!')}")
        return {i: part_results[i]["saved"] for i in range(1,len(parts)+1) if part_results[i]["saved"]}, set()

    print(f"\n{_bold('='*W)}")
    print(f"""
  {_bold('Commands:')}
    {_cyan('done')}              — files placed, proceed to merge
    {_cyan('skip <num>')}        — skip item #num (omit from merged PDF)
    {_cyan('skip <a> <b>')}      — skip multiple items
    {_cyan('skip all')}          — skip all remaining missing items
    {_cyan('list')}              — re-show missing items with links
    {_cyan('quit')}              — exit without merging
""")
    skipped=set()
    while True:
        still=[i for i in missing_idx if not _scan_folder(folder,i) and i not in skipped]
        if not still: print(f"\n  {_green('All items resolved!')} Proceeding …\n"); break
        try: raw=input(_bold(f"  [{len(still)} unresolved]  > ")).strip()
        except (EOFError,KeyboardInterrupt): print("\n  Cancelled."); sys.exit(0)
        if not raw: continue
        cmd=raw.lower()
        if cmd=="done": break
        elif cmd=="quit": print("  Exiting."); sys.exit(0)
        elif cmd=="list":
            for idx in still:
                p=parts[idx-1]; label=f"{p['manufacturer']} {p['part_number']}".strip()
                efn=_expected_filename(idx,p["manufacturer"],p["part_number"])
                cands=part_results[idx]["candidates"]
                print(f"\n  {_yellow(f'{idx:03d}')}  {_bold(label)}")
                print(f"       Save as: {_cyan(efn)}")
                if cands:
                    best=next((c for c in cands if _is_pdf_url(c["url"])),cands[0])
                    print(f"       URL:     {best['url']}")
        elif cmd.startswith("skip"):
            tokens=raw.split()[1:]
            if not tokens: print(f"  Usage: {_cyan('skip <num>')} or {_cyan('skip all')}"); continue
            if tokens[0].lower()=="all":
                new=set(still); skipped|=new; print(f"  {_yellow(f'Skipped {len(new)} item(s).')}")
            else:
                ok=[]
                for t in tokens:
                    try:
                        n=int(t)
                        if 1<=n<=len(parts): skipped.add(n); ok.append(n)
                        else: print(f"  {_red(f'No item #{n}')}")
                    except ValueError: print(f"  {_red(f'Not a number: {t!r}')}")
                if ok: print(f"  {_yellow('Skipped: '+', '.join(f'#{x:03d}' for x in sorted(ok)))}")
        else:
            print(f"  Unknown. Type {_cyan('done')}, {_cyan('skip')}, {_cyan('list')}, or {_cyan('quit')}")

    final={}
    for idx in range(1,len(parts)+1):
        files=_scan_folder(folder,idx)
        if files: final[idx]=files
    return final, skipped

# =============================================================================
#  MERGE
# =============================================================================

def merge_pdfs(bom_stem, parts, part_files, skipped, output_path, no_dividers=False, custom_logo_path=None):
    """Merge all found manuals into one Manufacturer's Record Book.

    Layout (default): a Table-of-Contents cover (one or more pages) listing every
    included manual with a clickable "Go to Datasheet" link, followed by each
    manual preceded by a "page breaker" page (Chapter # + Manufacturer/Part #).
    Every generated page carries a "Page X of Y" footer and the TOC<->breaker
    links are real internal PDF links. Manual pages themselves are left untouched.

    no_dividers=True  ->  plain merge: manuals only, no cover / breakers / footers.
    The call signature is unchanged so existing callers keep working.
    """
    writer = PdfWriter()
    tmp = tempfile.mkdtemp()
    missing = []
    print(f"\n  {_bold('Building merged PDF ...')}")

    # ---- Collect the included manuals (readable file, not skipped) -----------
    included = []   # each: dict(mfr, part, desc, manual_pages, mpc)
    for idx, part in enumerate(parts, 1):
        if idx in skipped:
            continue
        mfr = part.get("manufacturer", "")
        model = part.get("part_number", "")
        label = f"{mfr} {model}".strip() if mfr else model
        files = part_files.get(idx, [])
        if not files:
            missing.append(f"{idx:03d}. {label}")
            continue
        try:
            manual_pages = list(PdfReader(str(files[0])).pages)
        except Exception:
            missing.append(f"{idx:03d}. {label}")
            continue
        included.append({
            "mfr": mfr,
            "part": model,
            "desc": part.get("description", ""),
            "manual_pages": manual_pages,
            "mpc": len(manual_pages),
        })

    # Sequential chapter numbers (1..N) in BOM order.
    for n, it in enumerate(included, 1):
        it["chapter_no"] = n

    # ---- Plain merge (no record-book furniture) ------------------------------
    if no_dividers:
        for it in included:
            for pg in it["manual_pages"]:
                writer.add_page(pg)
        if len(writer.pages) == 0:
            writer.add_blank_page(width=_A4_W, height=_A4_H)
        out = Path(output_path)
        with open(out, "wb") as f:
            writer.write(f)
        _merge_report(out, len(writer.pages), missing)
        try:
            shutil.rmtree(tmp)
        except Exception:
            pass
        return out

    # ---- Record-book layout --------------------------------------------------
    if custom_logo_path and os.path.isfile(custom_logo_path):
        logo_toc = str(custom_logo_path)
        logo_brk = str(custom_logo_path)
    else:
        logo_toc = _write_logo(tmp, "toc", _LOGO_TOC_B64)
        logo_brk = _write_logo(tmp, "brk", _LOGO_BREAKER_B64)

    # Pass A: paginate the Table of Contents (measurement only).
    toc_rows = [{"chapter_no": it["chapter_no"], "desc": it["desc"],
                 "mfr": it["mfr"], "part": it["part"]} for it in included]
    pages_layout = _toc_paginate(toc_rows) if toc_rows else []
    cover_n = len(pages_layout)

    # Absolute (1-based) breaker page numbers and total page count.
    cursor = cover_n
    for it in included:
        it["breaker_index"] = cursor          # 0-based page index of this breaker
        cursor += 1 + it["mpc"]
    total_pages = cursor if included else cover_n

    # Pass B1: render the cover page(s), capturing link rects per row.
    cover_links = []   # (cover_page_index, rect, chapter_no)
    if pages_layout:
        cover_pdf = Path(tmp) / "cover.pdf"
        c = rl_canvas.Canvas(str(cover_pdf), pagesize=A4)
        for pidx, entries in enumerate(pages_layout):
            rects = _toc_draw_page(c, entries, pidx + 1, total_pages, logo_toc)
            for (chap_no, rect) in rects:
                cover_links.append((pidx, rect, chap_no))
            c.showPage()
        c.save()

    # Pass B2: render each breaker page.
    breaker_pdf = {}
    for it in included:
        bp = Path(tmp) / f"brk_{it['chapter_no']:04d}.pdf"
        c = rl_canvas.Canvas(str(bp), pagesize=A4)
        _draw_breaker_page(c, it["chapter_no"], it["mfr"], it["part"], it["desc"],
                           it["breaker_index"] + 1, total_pages, logo_brk)
        c.showPage()
        c.save()
        breaker_pdf[it["chapter_no"]] = bp

    # ---- Assemble: cover pages, then (breaker + manual) per chapter ----------
    if pages_layout:
        for pg in PdfReader(str(cover_pdf)).pages:
            writer.add_page(pg)

    breaker_actual = {}
    for it in included:
        breaker_actual[it["chapter_no"]] = len(writer.pages)
        for pg in PdfReader(str(breaker_pdf[it["chapter_no"]])).pages:
            writer.add_page(pg)
        for pg in it["manual_pages"]:
            writer.add_page(pg)

    if len(writer.pages) == 0:
        writer.add_blank_page(width=_A4_W, height=_A4_H)

    # ---- Internal links (degrade gracefully on older pypdf) ------------------
    try:
        from pypdf.annotations import Link
        from pypdf.generic import Fit, ArrayObject, NameObject

        def _link(page_idx, rect, target_idx):
            ann = writer.add_annotation(
                page_number=page_idx,
                annotation=Link(rect=rect, target_page_index=target_idx, fit=Fit.fit()),
            )
            # Replace the literal page-index destination with a real indirect page
            # reference so the link navigates in Acrobat / Chrome / Edge / etc.
            try:
                ref = writer.pages[target_idx].indirect_reference
                ann[NameObject("/Dest")] = ArrayObject([ref, NameObject("/Fit")])
            except Exception:
                pass

        for (cover_pi, rect, chap_no) in cover_links:
            tgt = breaker_actual.get(chap_no)
            if tgt is not None:
                _link(cover_pi, rect, tgt)
        for it in included:
            _link(breaker_actual[it["chapter_no"]], _BACK_TO_INDEX_RECT, 0)
    except Exception as exc:
        print(f"  {_dim(f'(internal links not added: {exc})')}")

    out = Path(output_path)
    with open(out, "wb") as f:
        writer.write(f)
    _merge_report(out, len(writer.pages), missing)
    try:
        shutil.rmtree(tmp)
    except Exception:
        pass
    return out


def _merge_report(out, pages, missing):
    mb = out.stat().st_size / (1024 * 1024)
    print(f"  {_green('OK')}  {_bold(out.name)}  ({pages} pages, {mb:.1f} MB)")
    if missing:
        print(f"\n  {len(missing)} item(s) skipped (no doc found):")
        for m in missing:
            print(f"    {_red('x')} {m}")

# =============================================================================
#  STARTUP MENU
# =============================================================================

def startup_menu():
    print(f"""
{_bold('  BOM Manual Downloader')}
  {"─"*40}
  {_bold('[1]')}  Load a BOM file  {_dim('(PDF / Excel .xlsx / Markdown .md)')}
  {_bold('[2]')}  Enter parts manually  {_dim('(no BOM file needed)')}
  {_bold('[3]')}  Quit
""")
    while True:
        try: choice=input("  Choice > ").strip()
        except (EOFError,KeyboardInterrupt): sys.exit(0)
        if choice=="1":
            try: raw=input("\n  BOM file path > ").strip().strip('"').strip("'")
            except (EOFError,KeyboardInterrupt): sys.exit(0)
            p=Path(raw)
            if not p.exists(): print(f"  {_red(f'File not found: {raw}')}"); continue
            return p, None
        elif choice=="2":
            parts=_enter_parts_manually()
            if not parts: print(f"  {_red('No parts entered.')}"); continue
            return None, parts
        elif choice=="3": sys.exit(0)
        else: print(f"  {_red('Enter 1, 2, or 3.')}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    ap=argparse.ArgumentParser(
        description="Download BOM manuals and merge into one ordered PDF.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples\n--------\n"
            "  python bom_downloader.py BOM.pdf\n"
            "  python bom_downloader.py BOM.xlsx\n"
            "  python bom_downloader.py BOM.md\n"
            "  python bom_downloader.py -P RITTAL 8808000 -P 'ALLEN BRADLEY' 1756-A4K\n"
            "  python bom_downloader.py        (interactive menu)\n"
        )
    )
    ap.add_argument("bom",         nargs="?",  default=None,
                    help="BOM file (PDF / .xlsx / .md). Omit for interactive menu.")
    ap.add_argument("--part","-P", nargs=2, action="append",
                    metavar=("MANUFACTURER","PART_NUMBER"),
                    help="Add a part directly. Repeatable. Skips BOM file.\n"
                         "  -P RITTAL 8808000\n  -P 'ALLEN BRADLEY' 1756-A4K")
    ap.add_argument("--folder", "-o", default=DOWNLOAD_FOLDER, help=f"Save folder (default: {DOWNLOAD_FOLDER})")
    ap.add_argument("--max",    "-n", type=int, default=MAX_PER_PART, help=f"Max PDFs per part (default {MAX_PER_PART})")
    ap.add_argument("--workers","-w", type=int, default=DEFAULT_WORKERS, help=f"Parallel workers (default {DEFAULT_WORKERS})")
    ap.add_argument("--preview","-p", action="store_true", help="Show parts list only, do not download")
    ap.add_argument("--no-merge",     action="store_true", help="Download only, skip merge step")
    ap.add_argument("--no-dividers",  action="store_true", help="Omit divider pages from merged PDF")
    ap.add_argument("--logo",         default=None,        help="Path to custom logo image (PNG/JPEG)")
    ap.add_argument("--output", "-O", default=None, help="Output path for merged PDF")
    ap.add_argument("--catalog", "-C", action="append", default=None, metavar="FOLDER",
                    help="Local catalog folder of your own PDFs (e.g. a OneDrive-synced "
                         "folder) used as a fallback when nothing is found online. "
                         "Repeatable. Files are matched by part number in the filename.")
    args=ap.parse_args()

    if args.catalog:
        set_catalog_dirs(args.catalog)

    bom_stem="manual_entry"

    if args.part:
        parts=[{"manufacturer":m.upper(),"part_number":p} for m,p in args.part]
        print(f"\n  {len(parts)} part(s) specified directly.")

    elif args.bom:
        bom_path=Path(args.bom)
        if not bom_path.exists(): print(f"  {_red(f'File not found: {args.bom}')}"); sys.exit(1)
        bom_stem=bom_path.stem; parts=extract_parts_from_bom(str(bom_path))

    else:
        bom_path_or_none,manual_parts=startup_menu()
        if bom_path_or_none: bom_stem=bom_path_or_none.stem; parts=extract_parts_from_bom(str(bom_path_or_none))
        else: parts=manual_parts

    if not parts: print(f"\n  {_red('No parts found.')}"); sys.exit(1)

    print(f"\n{'='*64}")
    print(f"  {_bold(f'{len(parts)} part(s) queued')}\n")
    for i,p in enumerate(parts,1):
        print(f"  {_dim(f'{i:>4}.')}  {(p['manufacturer'] or '(no mfr)'):<25}  {p['part_number']}")

    if args.preview: print(f"\n  Preview only."); return

    folder=Path(args.folder); folder.mkdir(parents=True,exist_ok=True)
    active_optional=[n for n,e,_ in _BACKENDS[:-1] if e() and not _cb.is_open(n)]
    chain="DirectProbe -> DDG" + (("  +  " + "  ".join(active_optional)) if active_optional else "")
    print(f"\n  Folder  : {folder.resolve()}")
    print(f"  Engines : {chain}")
    print(f"  Workers : {args.workers}  |  Max per part : {args.max}")
    print(f"{'='*64}\n")

    session=make_pooled_session()

    part_results={}; total_saved=0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        fut_map={ex.submit(_download_part,idx,p["manufacturer"],p["part_number"],folder,args.max,session,p.get("description","")):(idx,p)
                 for idx,p in enumerate(parts,1)}
        done=0
        for fut in as_completed(fut_map):
            idx,p=fut_map[fut]; done+=1; label=f"{p['manufacturer']} {p['part_number']}".strip()
            try:
                result=fut.result(); part_results[idx]=result; total_saved+=len(result["saved"])
                if result["saved"]: status=_green(f"{len(result['saved'])} file(s)")
                elif result["candidates"]: status=_yellow(f"0 saved  ({len(result['candidates'])} URL(s) found, download failed)")
                else: status=_red("0 files  (no URLs found)")
                tprint(f"[{done}/{len(parts)}]  {label}  -> {status}\n")
            except Exception as exc:
                part_results[idx]={"saved":[],"candidates":[]}
                tprint(f"[{done}/{len(parts)}]  {label}  {_red(f'ERROR: {exc}')}\n")

    for idx in range(1,len(parts)+1): part_results.setdefault(idx,{"saved":[],"candidates":[]})

    if args.no_merge:
        print(f"{'='*64}\n  Done. {total_saved} file(s). Merge skipped.\n{'='*64}"); return

    final_files,skipped=interactive_resolve(parts,part_results,folder)
    output=args.output or str(folder/f"{bom_stem}_merged.pdf")
    merged=merge_pdfs(bom_stem,parts,final_files,skipped,output,args.no_dividers,custom_logo_path=args.logo)

    print(f"\n{'='*64}")
    print(f"  {_green(_bold('Done!'))}")
    print(f"  Files  : {folder.resolve()}")
    print(f"  Merged : {_bold(str(merged.resolve()))}")
    print(f"{'='*64}")

if __name__=="__main__":
    main()