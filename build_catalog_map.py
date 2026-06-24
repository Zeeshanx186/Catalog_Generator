"""
build_catalog_map.py  —  Run this ONCE to auto-generate catalog_map.txt

How it works:
  1. Reads part numbers from your BOM file (xlsx, xls, or txt)
  2. Scans your catalog folder: checks filenames first (fast), then reads PDF text
  3. Writes catalog_map.txt so BOM Downloader finds the right PDF automatically

Usage:
  python build_catalog_map.py

Requires:  pip install pdfplumber openpyxl
"""

import re
import sys
import io
from pathlib import Path
from contextlib import redirect_stdout

# ─── CONFIGURE THESE TWO PATHS ────────────────────────────────────────────────
CATALOG_FOLDER = r"C:\Users\zeeshan.yaqoob\OneDrive - INTECH Process Automation\Documents\Catalog"

# Your BOM xlsx file (or plain text with one part number per line).
# Leave as "" to be prompted.
BOM_FILE = ""
# ──────────────────────────────────────────────────────────────────────────────


def compact(s):
    """Strip everything except letters and digits, lowercase."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def read_parts_from_bom(bom_path):
    """
    Extract part numbers from BOM.

    Primary path: uses bom_downloader.extract_parts_from_bom — the exact same
    logic the app uses, so part numbers always match what the app will search for.

    Fallback: reads raw cells from Excel, converting numeric cells (openpyxl
    returns 8108245 as int, not str — the old script missed every pure-number
    Rittal/Phoenix Contact part because of an isinstance(cell, str) check).
    """
    path = Path(bom_path)

    # ── Primary: reuse the app's own BOM parser ────────────────────────────
    try:
        script_dir = Path(__file__).parent
        if str(script_dir) not in sys.path:
            sys.path.insert(0, str(script_dir))
        import bom_downloader as bd
        buf = io.StringIO()
        with redirect_stdout(buf):
            parts_data = bd.extract_parts_from_bom(str(path))
        pns = list(dict.fromkeys(
            p["part_number"] for p in parts_data if p.get("part_number")
        ))
        if pns:
            print(f"  (Used bom_downloader parser — {len(pns)} distinct part numbers)")
            return pns
    except Exception as e:
        print(f"  (bom_downloader not available, using fallback: {e})")

    # ── Fallback: read cells directly from Excel ───────────────────────────
    parts = []
    if path.suffix.lower() in (".xlsx", ".xls"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue

                    # FIX: openpyxl returns numeric Excel cells as int/float, NOT str.
                    # The old code had isinstance(cell, str) which silently skipped
                    # every pure-number part like 8108245, 8660003, 3044102, etc.
                    if isinstance(cell, (int, float)):
                        n = int(cell)
                        if n == cell and n >= 10000:   # 5+ digit integer → likely a part number
                            parts.append(str(n))
                        continue

                    v = str(cell).strip()
                    if len(v) < 3:
                        continue

                    # Accept alphanumeric part numbers (letters + digits)
                    if re.search(r"[A-Za-z]", v) and re.search(r"[0-9]", v):
                        parts.append(v)
                    # Accept pure-digit part numbers ≥ 5 digits
                    elif re.sub(r"[\s.]", "", v).isdigit() and len(re.sub(r"[\s.]", "", v)) >= 5:
                        parts.append(re.sub(r"[\s.]", "", v))

            wb.close()
        except Exception as e:
            print(f"  Could not read Excel: {e}")

    elif path.exists():
        # Plain text — one part number per line
        with open(path, encoding="utf-8", errors="replace") as f:
            for line in f:
                v = line.strip()
                if v and not v.startswith("#"):
                    parts.append(v)

    return list(dict.fromkeys(parts))   # deduplicate, preserve order


def extract_text_first_pages(pdf_path, max_pages=5):
    """Compact text from first N pages (for Tier 2 text scan)."""
    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            chunks = []
            for pg in pdf.pages[:max_pages]:
                try:
                    t = pg.extract_text() or ""
                    chunks.append(t)
                except Exception:
                    pass
            return re.sub(r"[^a-z0-9]", "", " ".join(chunks).lower())
    except Exception:
        return ""


def main():
    catalog_dir = Path(CATALOG_FOLDER)
    if not catalog_dir.is_dir():
        print(f"\nERROR: Catalog folder not found:\n  {CATALOG_FOLDER}")
        print("Edit CATALOG_FOLDER at the top of this script.")
        sys.exit(1)

    pdfs = list(catalog_dir.rglob("*.pdf"))
    print(f"\nFound {len(pdfs)} PDF(s) in catalog folder.")

    global BOM_FILE
    if not BOM_FILE:
        BOM_FILE = input(
            "\nPath to your BOM file (xlsx / txt, or Enter to skip): "
        ).strip().strip('"').strip("'")

    parts = []
    if BOM_FILE and Path(BOM_FILE).exists():
        parts = read_parts_from_bom(BOM_FILE)
        print(f"  Read {len(parts)} part number(s) from BOM.")
    else:
        print("  No BOM file — filename scan only (no text scan).")

    if not parts:
        print("\nNo parts to match against. Provide a BOM file and rerun.")
        sys.exit(0)

    # Build a compact → original mapping for all parts
    # (handles dots: BOM '8601200' matches catalog file '8601.200' because
    # both compact to '8601200')
    part_compact_map = {}   # compact_pn → original_pn
    for pn in parts:
        c = compact(pn)
        if len(c) >= 4:   # skip very short tokens
            part_compact_map[c] = pn

    map_entries = {}    # original_pn → pdf_filename  (first confident hit wins)
    text_cache = {}     # pdf_path → compact_text (avoid re-reading)

    print(f"\nScanning PDFs…\n")
    for i, pdf in enumerate(pdfs, 1):
        name = pdf.name
        name_c = compact(name)
        label = name[:70] + ("…" if len(name) > 70 else "")
        print(f"  [{i:3d}/{len(pdfs)}] {label}", end="", flush=True)

        matched = []

        # ── Tier 1: part number in filename (fast, no PDF read) ───────────
        for pc, pn in part_compact_map.items():
            if pn not in map_entries and len(pc) >= 4 and pc in name_c:
                matched.append(pn)

        if matched:
            for pn in matched:
                map_entries[pn] = pdf.name
            print(f"  → filename: {matched}")
            continue

        # ── Tier 2: part number inside PDF text (slower, cached) ──────────
        text = extract_text_first_pages(pdf)
        if not text:
            print("  (no text)")
            continue

        text_matched = []
        for pc, pn in part_compact_map.items():
            if pn not in map_entries and len(pc) >= 5 and pc in text:
                text_matched.append(pn)

        if text_matched:
            for pn in text_matched:
                map_entries[pn] = pdf.name
            print(f"  → text: {text_matched}")
        else:
            print("  (no match)")

    # ── Write catalog_map.txt ─────────────────────────────────────────────
    out = catalog_dir / "catalog_map.txt"
    unmatched = [p for p in parts if p not in map_entries]

    with open(out, "w", encoding="utf-8") as f:
        f.write("# catalog_map.txt — generated by build_catalog_map.py\n")
        f.write("# Format:  <part_number> = <PDF filename stem>\n")
        f.write("# BOM Downloader reads this file automatically from the catalog folder.\n\n")
        f.write(f"# ── Matched: {len(map_entries)} of {len(parts)} parts ────────────────\n\n")
        for pn, fname in sorted(map_entries.items()):
            stem = re.sub(r"\.pdf$", "", fname, flags=re.I)
            f.write(f"{pn} = {stem}\n")
        if unmatched:
            f.write(f"\n\n# ── Not matched ({len(unmatched)} parts) — fill in manually ──────────────\n")
            for pn in unmatched:
                f.write(f"# {pn} = ???\n")

    print(f"\n{'='*60}")
    print(f"  catalog_map.txt written to:")
    print(f"  {out}")
    print(f"\n  Matched:   {len(map_entries)} / {len(parts)} parts")
    print(f"  Skipped:   {len(unmatched)} — edit the '# ???' lines manually")
    print(f"{'='*60}")
    print(f"\n  Run BOM Downloader again — it will use this map automatically.")


if __name__ == "__main__":
    main()
