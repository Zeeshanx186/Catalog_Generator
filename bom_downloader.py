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

import re, sys, time, argparse, threading, tempfile, shutil
from pathlib import Path
from urllib.parse import urljoin
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

NEXAR_CLIENT_ID     = ""   # nexar.com/api  (free 1k/month)
NEXAR_CLIENT_SECRET = ""

GOOGLE_CSE_KEY = ""        # console.cloud.google.com  (100 free/day)
GOOGLE_CSE_CX  = ""

SERPER_API_KEY = ""        # serper.dev  ($50/month 50k)
SERPAPI_KEY    = ""        # serpapi.com  ($50/month 5k)
BING_API_KEY   = ""        # Azure  (1k free/month)
TAVILY_API_KEY = ""        # app.tavily.com  (1k free/month)
EXA_API_KEY    = ""        # dashboard.exa.ai  (1k free/month)
BRAVE_API_KEY  = ""        # brave.com/search/api  ($3/1k)

MOUSER_API_KEY  = ""       # mouser.com/api-hub  (free)
FARNELL_API_KEY = ""       # partner.element14.com  (free)
FARNELL_STORE   = "uk.farnell.com"

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

    def record_success(self, name):
        with self._lock: self._fails[name] = 0; self._tripped.discard(name)

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

def _is_cancelled() -> bool:
    return _cancel_ev is not None and _cancel_ev.is_set()

def tprint(*a, **k):
    with _print_lock: print(*a, **k)
    _emit("log", message=" ".join(str(x) for x in a))

def _has(v):
    return bool(v and v.strip() and not v.strip().startswith("#"))

# =============================================================================
#  KNOWN MANUFACTURERS
# =============================================================================

_KNOWN_MFRS = sorted([
    ("ALLEN BRADLEY","ALLEN BRADLEY"),("PHOENIX CONTACT","PHOENIX CONTACT"),
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
    ("SWAGELOK","SWAGELOK"),("APOLLO","APOLLO"),
], key=lambda x: len(x[0]), reverse=True)

_MFR_PAT = [
    (re.compile(r'(?<![A-Za-z])' + re.escape(n) + r'(?![A-Za-z])', re.I), c)
    for n, c in _KNOWN_MFRS
]

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

def _dedup(parts):
    seen, out = set(), []
    for p in parts:
        k = (p["manufacturer"].upper(), p["part_number"].upper())
        if k not in seen and _looks_like_part(p["part_number"]):
            seen.add(k); out.append(p)
    return out

_HDR_MFR  = re.compile(r'(manufacturer|make|mfr|vendor|supplier)', re.I)
_HDR_PART = re.compile(r'(part[\s_\-]?no|part[\s_\-]?num|mpn|mfr[\s_\-]?pn|model[\s_\-]?no|p/?n\b)', re.I)
_HDR_SKIP = re.compile(r'(qty|quantity|ref|unit|price|tag|rev)', re.I)
_HDR_DESC = re.compile(r'(description|desc|item[\s_]?desc|material|component|service|function|notes?)', re.I)

def _find_cols(header):
    mc = next((i for i,h in enumerate(header) if _HDR_MFR.search(str(h)) and not _HDR_SKIP.search(str(h))), None)
    pc = [i for i,h in enumerate(header) if _HDR_PART.search(str(h)) and not _HDR_SKIP.search(str(h))]
    dc = next((i for i,h in enumerate(header) if _HDR_DESC.search(str(h)) and not _HDR_PART.search(str(h)) and not _HDR_MFR.search(str(h))), None)
    return mc, pc, dc

def _rows_to_parts(rows):
    if not rows or len(rows) < 2: return []
    cl = lambda v: re.sub(r'\s+', ' ', str(v or '')).strip()
    hdr = [cl(c) for c in rows[0]]; mc, pcs, dc = _find_cols(hdr)
    if not pcs: return []
    parts = []
    for row in rows[1:]:
        cells = [cl(c) for c in row]
        for pc in pcs:
            pn   = cells[pc] if pc < len(cells) else ""
            mfr  = cells[mc] if mc is not None and mc < len(cells) else ""
            desc = cells[dc] if dc is not None and dc < len(cells) else ""
            if _looks_like_part(pn): parts.append({"manufacturer": mfr, "part_number": pn, "description": desc})
    return parts

def _desc_keywords(desc):
    """Extract meaningful search keywords from a BOM item description."""
    if not desc: return []
    _STOP = {'the','a','an','and','or','for','to','of','in','with','by','at','from',
             'is','are','be','as','on','no','not','type','unit','module','assembly',
             'system','device','series','standard','general','purpose'}
    words = re.findall(r'[A-Za-z]{3,}', desc)
    return [w.lower() for w in words if w.lower() not in _STOP][:5]

def _model_in_text(text, model):
    """True only if model appears as a WHOLE token — not as a prefix of a longer part number.
    e.g. 'EN2T' must NOT match 'EN2TSC'.
    Handles: "1756-EN2T-nd" (separators), "1756EN2T.pdf" (numeric prefix), "EN2T.pdf".
    """
    ml = model.lower()
    tl = text.lower()
    mn = re.sub(r'[-_\s]+', '', ml)
    # Check 1: original text with full word boundaries (handles hyphenated/spaced forms)
    if re.search(r'(?<![A-Za-z0-9])' + re.escape(ml) + r'(?![A-Za-z0-9])', tl):
        return True
    # Check 2: split the text on URL/path separators and match each segment
    # This handles compact forms like "1756EN2T" where the digit prefix has no separator
    for token in re.split(r'[-_\s./+,;:()\[\]{}]', tl):
        tn = re.sub(r'[-_\s]+', '', token)
        if tn == mn:
            return True
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
                    model, desc = _split_model_desc(line[m.end():])
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
            for row in rows:
                text = " ".join(str(c or '') for c in row)
                if _is_cyr(text): continue
                for pat, can in _MFR_PAT:
                    m = pat.search(text)
                    if not m: continue
                    model, desc = _split_model_desc(text[m.end():])
                    if model and 2 <= len(model) <= 50:
                        parts.append({"manufacturer": can, "part_number": model, "description": desc})
                    break
    wb.close(); return _dedup(parts)

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
                model, desc = _split_model_desc(line[m.end():])
                if model and 2 <= len(model) <= 50: parts.append({"manufacturer": can, "part_number": model, "description": desc})
                break
    return _dedup(parts)

# ── dispatcher ────────────────────────────────────────────────────────────────
def extract_parts_from_bom(bom_path):
    path = Path(bom_path); ext = path.suffix.lower()
    print(f"\n  Reading BOM: {path.name}")
    if ext == ".pdf":               return _extract_pdf(path)
    elif ext in (".xlsx", ".xls"):  return _extract_excel(path)
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
    c = pn.strip(); nd = re.sub(r'[-_\s]','',c); dot = re.sub(r'[-_\s]','.',c)
    return c, nd, dot, c.lower(), nd.lower()

def _direct_candidates(model, mfr):
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
            (f"https://www.routeco.com/ResourceViewer?resource=RI/{nd}.pdf",          f"{model} (routeco)"),
            (f"https://www.rittal.com/img/products/DATA/{nd}.pdf",                    f"{model} (Rittal)"),
            (f"https://www.rittal.com/img/products/DATA/{c}.pdf",                     f"{model} (Rittal)"),
            (f"https://static.rittal.com/dokumente/{nd}.pdf",                         f"{model} (Rittal)"),
        ]
    if "allen" in ml or "rockwell" in ml:
        for dt in ("td","um","in","rm","pp","qr","sg"):
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
    """HEAD-probe a list of (url, title) pairs in parallel; return live PDF hits."""
    found = []
    def _p(url, title):
        try:
            _throttle(url)
            r = session.head(url, timeout=6, allow_redirects=True, headers={"User-Agent": BROWSER_UA})
            if r.status_code != 200: return None
            ct = r.headers.get("Content-Type", "")
            if "pdf" in ct.lower() or _is_pdf_url(url):
                return {"url": url, "title": title, "referer": None}
        except Exception: pass
        return None
    with ThreadPoolExecutor(max_workers=12) as ex:
        for fut in as_completed({ex.submit(_p, u, t): (u, t) for u, t in url_title_pairs}):
            r = fut.result()
            if r: found.append(r)
    return found

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
    "RITTAL": [
        "https://www.rittal.com/int_en/product/{model}/",
        "https://www.rittal.com/com_en/content/en/PRODUCTS/search.jsp?q={model}",
    ],
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


def _mfr_portal_search(model, mfr, session):
    """
    Scrape manufacturer product pages directly — no search engine required.
    For each known manufacturer we try their product detail page URL first,
    then their on-site search, returning any PDF links found.
    """
    if not mfr or _is_cancelled():
        return []

    mu = mfr.upper()
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

# -----------------------------------------------------------------------------
# Active search backends.
# DDG is always the last fallback and is NEVER circuit-broken.
# Uncomment the others once you fill in their API keys above.
# -----------------------------------------------------------------------------
_BACKENDS = [
    # ("Google CSE", lambda: _has(GOOGLE_CSE_KEY) and _has(GOOGLE_CSE_CX), _gcse),
    # ("Bing",       lambda: _has(BING_API_KEY),   _bing),
    # ("Tavily",     lambda: _has(TAVILY_API_KEY),  _tavily),
    # ("Exa",        lambda: _has(EXA_API_KEY),     _exa),
    # ("Serper",     lambda: _has(SERPER_API_KEY),  _serper),
    # ("SerpApi",    lambda: _has(SERPAPI_KEY),     _serpapi),
    # ("Brave",      lambda: _has(BRAVE_API_KEY),   _brave),
    ("DDG", lambda: True, None),   # handler assigned below
]

def _ddg(q, n=10, retries=3):
    """DuckDuckGo - always active, never circuit-broken."""
    for attempt in range(1, retries + 1):
        try:
            with DDGS() as d:
                r = list(d.text(q, max_results=n))
            if r: return r
            time.sleep(attempt * 2)
        except Exception:
            time.sleep(attempt * 4)
    return []

_BACKENDS[-1] = ("DDG", lambda: True, _ddg)   # wire in now that fn is defined

def _search(q, n=10):
    for name, enabled, fn in _BACKENDS:
        if not enabled(): continue
        if name != "DDG" and _cb.is_open(name): continue   # DDG never blocked
        r = fn(q, n)
        if r: return r
    return []

# =============================================================================
#  DOWNLOAD UTILITIES
# =============================================================================

MFR_DOMS  = ["rockwellautomation.com","rittal.com","schneider-electric.com",
             "phoenixcontact.com","weidmuller.com","hima.com","cisco.com",
             "moxa.com","honeywell.com","mtl-inst.com","prosoft-technology.com"]
DIST_DOMS = ["farnell.com","mouser.com","digikey.com","rs-online.com",
             "element14.com","avnet.com","arrow.com","tme.eu",
             "sminor.is","bolenscontrol.com","rspsupply.com","routeco.com"]

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

def _throttle(url):
    d = _dom(url)
    with _domain_locks[d]:
        e = time.monotonic() - _domain_times[d]
        if e < MIN_DOMAIN_GAP: time.sleep(MIN_DOMAIN_GAP - e)
        _domain_times[d] = time.monotonic()

def _is_pdf_url(url):
    u = url.lower(); return u.endswith(".pdf") or ".pdf?" in u or ".pdf#" in u

def _variants(m):
    v = {m.lower(), re.sub(r'[-_\s]+','',m).lower(), re.sub(r'[-_]+',' ',m).lower()}
    v |= {t.lower() for t in re.findall(r'[A-Za-z]{2,}|\d{2,}', m)}
    return list(v)

def _relevant(url, title, model, mfr):
    haystack = url + " " + title
    # Require the model as a whole token — EN2T must NOT match EN2TSC
    if _model_in_text(haystack, model):
        return True
    # Manufacturer CDN domain is a strong enough signal on its own,
    # but still require at least one significant chunk of the model
    if mfr and len(mfr) >= 4:
        ml = re.sub(r'[-_\s]+', '', mfr).lower()
        if ml[:6] in re.sub(r'[-_\s]+', '', haystack).lower():
            chunks = [t for t in re.findall(r'[A-Za-z]{2,}|\d{3,}', model)]
            if any(_model_in_text(haystack, c) for c in chunks[:2]):
                return True
    return False

def _score(url, title, model, mfr, desc=""):
    s=0; ul=url.lower(); tl=title.lower(); d=_dom(url)
    ns=re.sub(r'[-_\s]+','',model).lower(); ml=(mfr or "").lower()
    # Model match — whole-token only, strong penalty for substring-only matches
    if _model_in_text(url + " " + title, model): s += 20
    elif ns in re.sub(r'[-_\s]+','',ul+tl): s += 6   # compact form found but not as whole token — weak
    # Manufacturer / domain signals
    if ml and len(ml)>=4 and ml[:5] in d: s+=15
    elif any(x in d for x in MFR_DOMS): s+=10
    elif any(x in d for x in DIST_DOMS): s+=5
    if "datasheet" in ul: s+=8
    if "manual" in ul: s+=5
    if _model_in_text(title, model): s+=6
    elif model.lower() in tl: s+=3
    # Description keyword matches — each keyword found in title/url is extra confidence
    for kw in _desc_keywords(desc)[:4]:
        if kw in ul or kw in tl: s += 3
    # Freshness
    cur = datetime.now().year
    for offset in range(0, 4):
        if str(cur - offset) in ul or str(cur - offset) in tl:
            s += max(1, 4 - offset); break
    for old in range(cur - 5, cur - 15, -1):
        if str(old) in ul or str(old) in tl:
            s -= 4; break
    # Penalise if manufacturer name is absent from domain entirely
    if mfr and len(ml) >= 4 and ml[:5] not in d and not any(x in d for x in MFR_DOMS+DIST_DOMS):
        s -= 5
    return s

def _scrape_page(page_url, model, mfr, session):
    try:
        _throttle(page_url)
        r = session.get(page_url, timeout=12, headers={**BROWSER_HDR,"Referer":page_url})
        if r.status_code != 200: return []
        soup = BeautifulSoup(r.text,"html.parser")
        base = page_url.split('/')[0]+'//'+page_url.split('/')[2]; found={}
        for a in soup.find_all("a",href=True):
            href=urljoin(base,a["href"]); txt=a.get_text(strip=True) or href
            if _is_pdf_url(href) and _relevant(href,txt,model,mfr): found[href]={"title":txt,"referer":page_url}
        for raw in re.findall(r'https?://[^\s"\'<>]+\.pdf',r.text):
            if _relevant(raw,raw,model,mfr) and raw not in found: found[raw]={"title":raw.split('/')[-1],"referer":page_url}
        return [{"url":u,"title":v["title"],"referer":v["referer"]} for u,v in found.items()]
    except Exception: return []

def _mouser_scrape(model, mfr, session):
    """Scrape Mouser search results for datasheet links — no API key needed.
    Mouser product pages reliably carry the manufacturer datasheet PDF.
    """
    if _cb.is_open("Mouser-scrape"): return []
    try:
        q = f"{mfr} {model}".strip() if mfr else model
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

        _cb.record_success("Mouser-scrape")
        tprint(f"    [Mouser-scrape] {len(found)} hit(s) for {mfr} {model}".strip())
        return found[:6]
    except Exception as e:
        _cb.record_failure("Mouser-scrape", str(e)[:80]); return []


def _rs_scrape(model, mfr, session):
    """Scrape RS Online / RS Components for datasheet links — no API key needed."""
    if _cb.is_open("RS-scrape"): return []   # silently skip — already failed this run
    try:
        q = f"{mfr} {model}".strip() if mfr else model
        search_url = f"https://uk.rs-online.com/web/c/?searchTerm={requests.utils.quote(q)}"
        _throttle(search_url)
        r = session.get(search_url, timeout=15,
                        headers={**BROWSER_HDR, "Referer": "https://uk.rs-online.com/"})
        if r.status_code != 200:
            _cb.record_failure("RS-scrape", f"HTTP {r.status_code}")
            tprint(f"    [RS-scrape] HTTP {r.status_code} — disabling for this run")
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
        "894":    ["HIMA HIMatrix", "HIQuad X order"],
        "Z 71":   ["HIMA cable assembly", "Z 71 connection"],
    }

    queries = [
        f'HIMA "{model}" datasheet',
        f'HIMA "{model}" manual filetype:pdf',
    ]

    for prefix, aliases in _FAMILIES.items():
        if m.startswith(prefix):
            for alias in aliases[:2]:   # top 2 aliases only to limit DDG calls
                queries.append(f'HIMA "{model}" {alias}')
                queries.append(f'HIMA {short} {alias} datasheet')
            # Short-form query (e.g. "HIMA F-BASE" instead of "HIMA F-BASE RACK 01")
            if short != m and len(short) >= 4:
                queries.append(f'HIMA "{short}" datasheet filetype:pdf')
                queries.append(f'"{short}" HIMA safety module datasheet')
            break

    # Also try known third-party HIMA documentation sites directly
    queries += [
        f'site:sds-automatyka.pl HIMA {short}',
        f'site:rms-dcs.com HIMA {short}',
        f'HIMA {short} "data sheet" OR "technical data" filetype:pdf',
    ]

    return queries


def _find_pdfs(model, mfr, session, desc=""):
    pf = f"{mfr} {model}".strip() if mfr else model
    direct = {}

    def _add(hits):
        for h in hits:
            direct.setdefault(h["url"], {"title": h["title"], "referer": h.get("referer")})

    # ── Tier 1: DirectProbe (hardcoded CDN URL patterns) ─────────────────────
    hits = _direct_probe(model, mfr, session)
    _add(hits)
    if hits: tprint(f"    [DirectProbe] {len(hits)} hit(s) for {pf}")

    # ── Tier 1.5: Manufacturer portal — go straight to the source ────────────
    # Scrapes the manufacturer's own product/documentation page directly.
    # Far more accurate than any search engine for known manufacturers.
    portal_hits = _mfr_portal_search(model, mfr, session)
    _add(portal_hits)
    if _is_cancelled(): return []

    # If we already have good results from the manufacturer's own portal,
    # skip the search-engine tiers entirely (faster + more accurate).
    if len(direct) >= 1 and any(
        any(mfr_dom in u for mfr_dom in MFR_DOMS)
        for u in direct
    ):
        tprint(f"    [MfrPortal] Found {len(direct)} doc(s) from manufacturer directly — skipping search engines")
        # Still run Nexar in case it has more/better options
        _add(_nexar_find(model, mfr))
        # Skip straight to ranking
        ranked = sorted(
            [(u, v) for u, v in direct.items()
             if _score(u, v["title"], model, mfr, desc) >= 10],
            key=lambda kv: _score(kv[0], kv[1]["title"], model, mfr, desc),
            reverse=True,
        )
        if not ranked:
            ranked = sorted(direct.items(), key=lambda kv: _score(kv[0], kv[1]["title"], model, mfr, desc), reverse=True)[:3]
        return [{"url": u, "title": v["title"], "referer": v["referer"]} for u, v in ranked[:10]]

    # ── Tier 2: Nexar component database (if key configured) ──────────────────
    _add(_nexar_find(model, mfr))
    if _is_cancelled(): return []

    # ── Tier 3: Manufacturer-portal site-specific DDG search ─────────────────
    # DDG honours site: — searching literature.rockwellautomation.com directly
    # returns exact documents rather than random distributor pages.
    doc_site = _mfr_doc_site(mfr)
    if doc_site:
        site_qs = [
            f'site:{doc_site} "{model}"',
            f'site:{doc_site} {model} datasheet',
            f'site:{doc_site} {model} manual',
        ]
        pages = []
        for q in site_qs:
            for r in _ddg(q, n=8):          # bypass _search so DDG always runs here
                url = r.get("href", "").strip(); title = r.get("title", "").strip()
                if not url: continue
                if _is_pdf_url(url) and _relevant(url, title, model, mfr):
                    direct.setdefault(url, {"title": title, "referer": None})
                elif url not in pages:
                    pages.append(url)
        if pages:
            with ThreadPoolExecutor(max_workers=SCRAPE_WORKERS) as ex:
                for fut in as_completed(
                    {ex.submit(_scrape_page, pg, model, mfr, session): pg
                     for pg in pages[:6]}
                ):
                    _add(fut.result())
        if direct: tprint(f"    [SiteSearch:{doc_site}] {len(direct)} candidate(s)")

    # ── Tier 3.5: HIMA-specific multi-query search ───────────────────────────
    # HIMA module names ("F-BASE RACK 01") don't appear verbatim in docs.
    # We generate product-family-aware queries to find the right manuals.
    gen_pages = []   # shared page list used by Tier 3.5 and Tier 5
    if mfr and "HIMA" in mfr.upper() and not direct:
        for q in _hima_queries(model):
            if _is_cancelled(): break
            for r in _ddg(q, n=6):
                url = r.get("href", "").strip(); title = r.get("title", "").strip()
                if not url: continue
                if _is_pdf_url(url) and _relevant(url, title, model, mfr):
                    direct.setdefault(url, {"title": title, "referer": None})
                elif not _is_pdf_url(url) and url not in gen_pages:
                    gen_pages.append(url)
            if direct:
                tprint(f"    [HIMA-search] {len(direct)} candidate(s) found after query: {q[:60]}")
                break  # found something — stop trying more queries

    # ── Tier 4: Distributor scrapers (no API key required) ────────────────────
    if _is_cancelled(): return []
    _add(_mouser_scrape(model, mfr, session))
    _add(_rs_scrape(model, mfr, session))
    if _is_cancelled(): return []

    # ── Tier 5: General DDG search + page scraping (broad fallback) ───────────
    ms     = re.sub(r'[-_\s]+', '', model)
    kw     = _desc_keywords(desc)
    kw_str = " ".join(kw[:2]) if kw else ""
    gen_qs = [
        f'"{pf}" {kw_str} datasheet'.strip() if kw_str else f'"{pf}" datasheet',
        f'"{pf}" {kw_str} manual'.strip()     if kw_str else f'"{pf}" manual',
        f'"{pf}" user guide',
        f'{pf} datasheet filetype:pdf',        # for Google CSE / Serper when configured
        f'{pf} manual filetype:pdf',
    ]
    if ms.lower() != model.lower():
        gen_qs.append(f'"{mfr} {ms}" {kw_str} datasheet'.strip() if mfr else f'"{ms}" datasheet')
    for q in gen_qs:
        for r in _search(q, n=10):
            url = r.get("href", "").strip(); title = r.get("title", "").strip()
            if not url: continue
            if _is_pdf_url(url):
                if _relevant(url, title, model, mfr):
                    direct.setdefault(url, {"title": title, "referer": None})
            else:
                mh = mfr and len(mfr) >= 4 and mfr.lower()[:5] in _dom(url)
                if (_relevant(url, title, model, mfr) or mh) and url not in gen_pages:
                    gen_pages.append(url)
    if gen_pages:
        with ThreadPoolExecutor(max_workers=SCRAPE_WORKERS) as ex:
            for fut in as_completed(
                {ex.submit(_scrape_page, pg, model, mfr, session): pg
                 for pg in gen_pages[:6]}
            ):
                _add(fut.result())

    # ── Tier 6: Last-ditch ────────────────────────────────────────────────────
    if not direct:
        for r in _search(f'{pf} {kw_str} pdf'.strip(), n=8):
            url = r.get("href", "").strip()
            if url and _is_pdf_url(url):
                direct.setdefault(url, {"title": r.get("title", ""), "referer": None})

    # ── Rank and filter ───────────────────────────────────────────────────────
    MIN_SCORE = 10
    ranked = sorted(
        [(u, v) for u, v in direct.items()
         if _score(u, v["title"], model, mfr, desc) >= MIN_SCORE],
        key=lambda kv: _score(kv[0], kv[1]["title"], model, mfr, desc),
        reverse=True,
    )
    if not ranked:  # nothing passed threshold — return top-3 unfiltered
        ranked = sorted(
            direct.items(),
            key=lambda kv: _score(kv[0], kv[1]["title"], model, mfr, desc),
            reverse=True,
        )[:3]
    return [{"url": u, "title": v["title"], "referer": v["referer"]}
            for u, v in ranked[:10]]

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
            try:
                _throttle(attempt_url)
                r = session.get(attempt_url, timeout=30, stream=True, headers=hdrs)
                if r.status_code in (403, 401, 406, 429) and i < 3:
                    time.sleep(2 ** i); continue
                if r.status_code in (403, 401) and attempt_url != url:
                    break  # this alt URL also blocked — try next
                r.raise_for_status()
                with open(dest, "wb") as f:
                    for chunk in r.iter_content(16384): f.write(chunk)
                kb = dest.stat().st_size // 1024
                if kb < 10: dest.unlink(); break
                with open(dest, "rb") as f:
                    if not f.read(5).startswith(b"%PDF"): dest.unlink(); break
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

def _download_part(bom_idx, mfr, model, folder, max_dl, session, desc=""):
    """Returns {"saved":[Path,...], "candidates":[{url,title,referer},...]}"""
    if _is_cancelled():
        return {"saved": [], "candidates": []}
    label = f"{mfr} {model}".strip() if mfr else model
    prefix = f"{bom_idx:03d}"
    _emit("part_start", idx=bom_idx, label=label)
    candidates = _find_pdfs(model, mfr, session, desc)
    if not candidates:
        tprint(f"    [{label}] {_red('No URLs found')}")
        _emit("part_done", idx=bom_idx, label=label, status="not_found", files=[])
        return {"saved": [], "candidates": []}
    safe = _sanitize(label.replace(" ", "_")); saved = []
    for idx, item in enumerate(candidates, 1):
        if _is_cancelled(): break
        if len(saved) >= max_dl: break
        url, title, ref = item["url"], item["title"], item.get("referer")
        dest = folder / f"{prefix}_{safe}__{_sanitize(title.replace(' ','_'))}.pdf"
        if dest.exists():
            tprint(f"    [{label}] Already exists: {dest.name}")
            saved.append(dest); continue
        tprint(f"    [{label}] [{idx}/{len(candidates)}] {title[:60]}")
        tprint(f"         {url[:80]}")
        if _write_pdf(url, dest, session, ref): saved.append(dest)
    if not saved:
        tprint(f"    [{label}] {_yellow('Could not save')} — {len(candidates)} URL(s) found")
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

def _make_divider(idx, mfr, model, tmp):
    p=Path(tmp)/f"div_{idx:03d}.pdf"; label=f"{mfr} {model}".strip() if mfr else model
    W,H=A4; c=rl_canvas.Canvas(str(p),pagesize=A4)
    c.setFillColorRGB(0.12,0.28,0.49); c.rect(0,0,18,H,fill=1,stroke=0)
    c.roundRect(50,H/2-10,56,34,6,fill=1,stroke=0)
    c.setFillColorRGB(1,1,1); c.setFont("Helvetica-Bold",18); c.drawCentredString(78,H/2+8,f"{idx:03d}")
    c.setFillColorRGB(0.1,0.1,0.1); c.setFont("Helvetica-Bold",22); c.drawString(120,H/2+8,label[:45])
    c.setStrokeColorRGB(0.12,0.28,0.49); c.setLineWidth(1.2); c.line(120,H/2+4,W-40,H/2+4)
    c.setFillColorRGB(0.5,0.5,0.5); c.setFont("Helvetica",10); c.drawString(120,H/2-12,f"BOM item #{idx:03d}")
    c.save(); return p

def _make_cover(bom_stem, parts, part_results, skipped, tmp):
    p=Path(tmp)/"cover.pdf"; W,H=A4; LH=14; TOP=H-120; MAX=int((TOP-60)/LH)
    c=rl_canvas.Canvas(str(p),pagesize=A4); pg=1
    def hdr():
        c.setFillColorRGB(0.12,0.28,0.49); c.rect(0,H-70,W,70,fill=1,stroke=0)
        c.setFillColorRGB(1,1,1); c.setFont("Helvetica-Bold",18); c.drawString(40,H-40,"BOM Manuals Package")
        c.setFont("Helvetica",10); c.drawString(40,H-58,f"Source: {bom_stem}")
        c.drawRightString(W-40,H-58,datetime.now().strftime("%Y-%m-%d %H:%M"))
        if pg>1: c.setFillColorRGB(0.5,0.5,0.5); c.setFont("Helvetica",9); c.drawRightString(W-40,20,f"Page {pg}")
    def col_hdr(y):
        c.setFillColorRGB(0.92,0.92,0.92); c.rect(30,y-3,W-60,LH+2,fill=1,stroke=0)
        c.setFillColorRGB(0.2,0.2,0.2); c.setFont("Helvetica-Bold",9)
        for x,t in ((36,"#"),(70,"Manufacturer"),(210,"Part Number"),(370,"Status"),(455,"Filename")): c.drawString(x,y,t)
    hdr()
    fnd=sum(1 for i,r in part_results.items() if r.get("saved") and i not in skipped)
    miss=sum(1 for i,r in part_results.items() if not r.get("saved") and i not in skipped)
    for x,clr,txt in ((40,(0.1,0.55,0.1),f"Found: {fnd}"),(130,(0.75,0.1,0.1),f"Not found: {miss}"),
                       (250,(0.5,0.5,0.5),f"Skipped: {len(skipped)}"),(340,(0.3,0.3,0.3),f"Total: {len(parts)}")):
        c.setFillColorRGB(*clr); c.setFont("Helvetica-Bold",11); c.drawString(x,H-92,txt)
    y=TOP; rows=0; col_hdr(c); y-=LH+4
    for idx,part in enumerate(parts,1):
        if rows>=MAX: c.showPage(); pg+=1; hdr(); y=TOP; rows=0; col_hdr(c); y-=LH+4
        r=part_results.get(idx,{"saved":[],"candidates":[]}); found=bool(r.get("saved")); skip=(idx in skipped)
        if rows%2==0: c.setFillColorRGB(0.97,0.97,0.97); c.rect(30,y-3,W-60,LH,fill=1,stroke=0)
        c.setFillColorRGB(0.4,0.4,0.4); c.setFont("Helvetica",8); c.drawString(36,y,f"{idx:03d}")
        c.setFillColorRGB(0.1,0.1,0.1); c.drawString(70,y,(part["manufacturer"] or "")[:22]); c.drawString(210,y,part["part_number"][:28])
        if skip:
            c.setFillColorRGB(0.5,0.5,0.5); c.setFont("Helvetica-Oblique",8); c.drawString(370,y,"Skipped")
        elif found:
            c.setFillColorRGB(0.1,0.6,0.1); c.circle(378,y+4,3,fill=1,stroke=0); c.setFont("Helvetica",8); c.drawString(386,y,"Found")
            fn=r["saved"][0].name; disp=fn if len(fn)<=30 else "…"+fn[-29:]
            c.setFillColorRGB(0.2,0.2,0.5); c.setFont("Helvetica-Oblique",7); c.drawString(455,y,disp)
        else:
            c.setFillColorRGB(0.8,0.1,0.1); c.circle(378,y+4,3,fill=1,stroke=0); c.setFont("Helvetica",8); c.drawString(386,y,"Not found")
        y-=LH; rows+=1
    c.setFillColorRGB(0.5,0.5,0.5); c.setFont("Helvetica",8); c.drawCentredString(W/2,20,"BOM Downloader")
    c.save(); return p

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

def merge_pdfs(bom_stem, parts, part_files, skipped, output_path, no_dividers=False):
    writer=PdfWriter(); tmp=tempfile.mkdtemp(); missing=[]; pages=0
    print(f"\n  {_bold('Building merged PDF ...')}")
    try:
        cover=_make_cover(bom_stem,parts,
            {i:part_files.get(i,{"saved":[],"candidates":[]}) if not isinstance(part_files.get(i),list) else {"saved":part_files[i],"candidates":[]} for i in range(1,len(parts)+1)},
            skipped,tmp)
        for pg in PdfReader(str(cover)).pages: writer.add_page(pg); pages+=1
    except Exception as e: print(f"  [Merge] cover error: {e}")
    for idx,part in enumerate(parts,1):
        if idx in skipped: continue
        mfr=part["manufacturer"]; model=part["part_number"]; label=f"{mfr} {model}".strip() if mfr else model
        files=part_files.get(idx,[])
        if not no_dividers:
            try:
                div=_make_divider(idx,mfr,model,tmp)
                for pg in PdfReader(str(div)).pages: writer.add_page(pg); pages+=1
            except Exception: pass
        if files:
            try:
                for pg in PdfReader(str(files[0])).pages: writer.add_page(pg); pages+=1
            except Exception:
                ph=_make_placeholder(idx,mfr,model,tmp)
                for pg in PdfReader(str(ph)).pages: writer.add_page(pg); pages+=1
        else:
            missing.append(f"{idx:03d}. {label}")
            try:
                ph=_make_placeholder(idx,mfr,model,tmp)
                for pg in PdfReader(str(ph)).pages: writer.add_page(pg); pages+=1
            except Exception: pass
    out=Path(output_path)
    with open(out,"wb") as f: writer.write(f)
    mb=out.stat().st_size/(1024*1024)
    print(f"  {_green('OK')}  {_bold(out.name)}  ({pages} pages, {mb:.1f} MB)")
    if missing:
        print(f"\n  {len(missing)} placeholder(s) inserted (no doc found):")
        for m in missing: print(f"    {_red('x')} {m}")
    try: shutil.rmtree(tmp)
    except Exception: pass
    return out

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
    ap.add_argument("--output", "-O", default=None, help="Output path for merged PDF")
    args=ap.parse_args()

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

    session=requests.Session(); session.headers.update({"User-Agent":BROWSER_UA})

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
    merged=merge_pdfs(bom_stem,parts,final_files,skipped,output,args.no_dividers)

    print(f"\n{'='*64}")
    print(f"  {_green(_bold('Done!'))}")
    print(f"  Files  : {folder.resolve()}")
    print(f"  Merged : {_bold(str(merged.resolve()))}")
    print(f"{'='*64}")

if __name__=="__main__":
    main()
