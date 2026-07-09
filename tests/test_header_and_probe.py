import sys
sys.path.insert(0, r"D:\APP\BomDownloader")
import bom_downloader as bd

# ── Bug 1: repeated header cells leak through as part numbers ─────────────────
# A BOM whose header row repeats mid-table (multi-page PDF extract, stacked
# Excel sections) produced parts literally named "Manufacturer Part No".
assert not bd._looks_like_part("Manufacturer Part No")
assert not bd._looks_like_part("Part Number")
assert not bd._looks_like_part("Model No")
assert not bd._looks_like_part("MPN")
# ...but real part numbers still pass, including digitless family codes
assert bd._looks_like_part("8108.245")
assert bd._looks_like_part("MB3270I-T")
assert bd._looks_like_part("A9F44201")
assert bd._looks_like_part("WTR")

rows = [
    ["Manufacturer", "Manufacturer Part No", "Description"],
    ["WEIDMULLER", "1758260000", "Terminal block"],
    ["Manufacturer", "Manufacturer Part No", "Description"],   # repeated header
    ["SCHNEIDER", "A9F44201", "MCB 1P"],
]
parts = bd._rows_to_parts(rows)
assert [p["part_number"] for p in parts] == ["1758260000", "A9F44201"], parts

# ── Bug 2: HEAD 200 + text/html on a .pdf URL counted as a live PDF ──────────
# catalog.weidmueller.com answers HEAD with 200 text/html for asset URLs it
# does NOT serve as PDFs; the probe must byte-sniff instead of trusting the
# ".pdf" in the URL, otherwise a phantom "official doc" suppresses the whole
# search-engine cascade and the part ends not_found.
class _Resp:
    def __init__(self, status, ct):
        self.status_code = status
        self.headers = {"Content-Type": ct}

class _FakeSession:
    def __init__(self, head_ct, body=b"<html>Access Denied</html>"):
        self.head_ct = head_ct
        self.body = body
    def head(self, url, **kw):
        return _Resp(200, self.head_ct)
    def get(self, url, **kw):
        class _G:
            status_code = 200
            headers = {"Content-Type": self.head_ct}
            def iter_content(_s, n): yield self.body[:n]
            def close(_s): pass
            def __enter__(_s): return _s
            def __exit__(_s, *a): return False
        return _G()

pairs = [("https://catalog.weidmueller.com/assets/546040000.pdf", "546040000 (Weidmuller)")]

# HTML masquerading behind a .pdf URL → NOT a hit
hits = bd._probe_url_list(pairs, _FakeSession("text/html"))
assert hits == [], hits

# Same URL, real PDF content-type → hit
hits = bd._probe_url_list(pairs, _FakeSession("application/pdf"))
assert len(hits) == 1, hits

# HTML content-type but the body really IS a PDF (mislabelling server) → hit
hits = bd._probe_url_list(pairs, _FakeSession("text/html", body=b"%PDF-1.7 rest"))
assert len(hits) == 1, hits

print("OK")

# ── Bug 3: family prefix alone counted as a partial part-match ────────────────
# '1756' appears in every ControlLogix doc, so 1756-RM2K (redundancy module)
# was verified against the 274-page I/O modules spec book.
io_spec_text = "1756 ControlLogix I/O Specifications 1756-IA8D 1756-IA16 1756-IB16 technical data"
rm_doc_text  = "1756 ControlLogix Redundancy Modules Catalog Numbers 1756-RM2 1756-RM2XT"

frags = bd._partial_fragments("1756-RM2K")
assert not any(bd._model_in_text(io_spec_text, c) for c in frags), frags
assert any(bd._model_in_text(rm_doc_text, c) for c in frags), frags

# chassis doc listing 1756-A10 still partial-matches the K variant
chassis_text = "ControlLogix Chassis Standard Catalog Numbers: 1756-A4, 1756-A7, 1756-A10"
assert any(bd._model_in_text(chassis_text, c) for c in bd._partial_fragments("1756-A10K"))

# 'MTL' alone must not identify MTL 5541; the number must
assert bd._partial_fragments("MTL 5541") == ["5541"]

# homoglyph normalization
assert bd._fix_homoglyphs("1756-А10K") == "1756-A10K"

print("OK2")

# ── Bug 4: a date in the manufacturer column leaks in as the maker name ───────
# openpyxl (data_only=True) yields real datetime objects for cells Excel auto-
# formatted as dates. Row 183 of a real BOM had 2026-03-26 where the maker
# belonged; the app then searched "2026-03-26 00:00:00 60LF4-40/20-SOG".
import datetime as _dt
assert bd._is_datelike(_dt.datetime(2026, 3, 26))
assert bd._is_datelike(_dt.date(2026, 3, 26))
assert bd._is_datelike("2026-03-26 00:00:00")
assert bd._is_datelike("26/03/2026")
assert not bd._is_datelike("AUTOCLAVE")
assert not bd._is_datelike("60LF4-40/20-SOG")     # part number, not a date
assert not bd._is_datelike("")

# datetime object in the manufacturer column → part kept, manufacturer blanked
rows = [
    ["Manufacturer", "Part Number", "Description"],
    ["AUTOCLAVE", "30VM4001-SOG", "Needle valve"],
    [_dt.datetime(2026, 3, 26), "60LF4-40/20-SOG", "Inline filter"],
]
parts = bd._rows_to_parts(rows)
assert [p["part_number"] for p in parts] == ["30VM4001-SOG", "60LF4-40/20-SOG"], parts
assert parts[1]["manufacturer"] == "", parts[1]

# string form (as it arrives from PDF/docx tables) blanked the same way
rows2 = [
    ["Manufacturer", "Part Number", "Description"],
    ["2026-03-26 00:00:00", "60LF4-40/20-SOG", "Inline filter"],
]
assert bd._rows_to_parts(rows2)[0]["manufacturer"] == ""

print("OK3")

# ── Bug 5: a generic brochure on the maker's own site short-circuits search ────
# hima.com hosts only marketing brochures ("A New Dimension of Performance…"),
# not the F-module manuals (those live on mirror sites). Treating a brochure as
# "the official doc found" made the app skip the mirror tier and report the part
# not-found. _own_doc_identifies must reject brochures but accept real docs.
brochure_url   = "https://www.hima.com/sharepoint-sync/PDFs/HIQuad+X/HIMA_Brochure.pdf"
brochure_title = "A New Dimension of Performance for Your Safety System"
assert not bd._own_doc_identifies(brochure_url, brochure_title, "F-CPU 01")

# part number present in the title → identifying
assert bd._own_doc_identifies("https://www.hima.com/x.pdf", "F-CPU 01 module manual", "F-CPU 01")
# explicit manual/datasheet filename → identifying even without the part number
assert bd._own_doc_identifies("https://www.moxa.com/eds-316-installation-manual.pdf",
                              "EDS-316 Series User Manual", "EDS-316-SS-SC-T")
# a bare datasheet-named file → identifying
assert bd._own_doc_identifies("https://site/1206421_datasheet.pdf", "", "1206421")

print("OK4")

# ── Bug 6: transient Zyte blips must not disable it for the whole run ──────────
# 3 transient 520/timeout errors used to trip Zyte (TRIP_THRESHOLD), and once
# the only API backend is down the cascade fast-skips every part during a DDG
# rate-limit pause (the Schneider A9F44xxx "cooling down — skipping" batch).
cb = bd.CircuitBreaker()
for _ in range(bd.TRIP_THRESHOLD + 1):        # more than the HARD threshold
    cb.record_soft("Zyte", "520 blip")
assert not cb.is_open("Zyte"), "transient blips tripped Zyte too early"

# a success clears the transient count
cb.record_success("Zyte")
for _ in range(bd.SOFT_TRIP_THRESHOLD - 1):
    cb.record_soft("Zyte", "520 blip")
assert not cb.is_open("Zyte")
cb.record_soft("Zyte", "520 blip")            # crosses SOFT_TRIP_THRESHOLD
assert cb.is_open("Zyte"), "sustained outage should still disable Zyte"

# decisive auth/credit error disables immediately
cb2 = bd.CircuitBreaker()
cb2.trip("Zyte", "HTTP 401")
assert cb2.is_open("Zyte")

print("OK5")
