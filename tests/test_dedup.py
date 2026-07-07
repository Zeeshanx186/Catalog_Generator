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
