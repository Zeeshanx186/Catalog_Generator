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
