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
    n_files = 2 if model == "8108.245" else 1
    saved = []
    for j in range(n_files):
        dest = Path(folder) / f"{idx:03d}_{mfr}_{model}_{j}.pdf"
        _write_pdf(dest)
        saved.append(dest)
    bd._emit("part_done", idx=idx, label=f"{mfr} {model}", status="found",
             files=[str(f) for f in saved])
    return {"saved": saved, "candidates": []}

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
                             "max_per_part": 2, "no_merge": True})
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
assert len(st[3]["files"]) == 2, st[3]["files"]
assert len(set(st[3]["files"])) == 2, "copied filenames must be distinct"
for f in st[3]["files"]:
    assert "COPY_OF_001" in f and Path(f).is_file(), f
assert st[3]["duplicate_of"] == 1
# Duplicate of a not-found part: not found, no files
assert st[4]["status"] == "not_found"
assert st[5]["status"] == "not_found" and st[5]["files"] == []
# Counters include duplicates
assert prog["done"] == 5 and prog["found"] == 3 and prog["not_found"] == 2, \
    (prog["done"], prog["found"], prog["not_found"])
shutil.rmtree(out_dir, ignore_errors=True)
print("OK")
