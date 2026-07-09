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
