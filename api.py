"""
api.py — Python API exposed to the frontend via pywebview's JS bridge.

Every public method is callable from JavaScript as:
    const result = await window.pywebview.api.method_name(args)
All methods return plain dicts (JSON-serialisable).
"""

from __future__ import annotations

import base64
import collections
import copy
import json
import os
import shutil
import subprocess
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed, wait as futures_wait
from pathlib import Path

import requests
import webview

import bom_downloader as bd


# ── Settings helpers ──────────────────────────────────────────────────────────

_CONFIG_PATH = Path(os.path.expandvars(r"%APPDATA%")) / "BomDownloader" / "config.json"
_DEFAULTS = {
    "output_folder": str(Path.home() / "Documents" / "BOM Manuals"),
    "max_per_part": 1,
    "workers": 5,
    "no_merge": False,
    "catalog_folder": "",
    "custom_logo_path": "",
}


def _load_cfg() -> dict:
    try:
        if _CONFIG_PATH.exists():
            return {**_DEFAULTS, **json.loads(_CONFIG_PATH.read_text())}
    except Exception:
        pass
    return dict(_DEFAULTS)


def _save_cfg(cfg: dict) -> None:
    try:
        _CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        _CONFIG_PATH.write_text(json.dumps(cfg, indent=2))
    except Exception:
        pass


# ── Connection pool ───────────────────────────────────────────────────────────

def _make_session() -> requests.Session:
    # Delegate to the single pooled-session factory in bom_downloader so the
    # connection-pool sizing is identical no matter which entry point runs.
    return bd.make_pooled_session()


# ── Main API class ────────────────────────────────────────────────────────────

class API:
    """Exposed to JavaScript as window.pywebview.api"""

    def __init__(self) -> None:
        self.window: webview.Window | None = None

        self._lock      = threading.Lock()
        self._log_lock  = threading.Lock()   # tiny, dedicated to the activity log
                                             # so high-frequency log events never
                                             # contend with the deepcopy in
                                             # get_progress() (that contention is
                                             # what froze the UI under load).
        self._log: "collections.deque[str]" = collections.deque(maxlen=300)
        self._cancel_ev = threading.Event()
        self._job_thread: threading.Thread | None = None
        self._executor: ThreadPoolExecutor | None = None  # kept so cancel() can kill queued futures
        self._active_workers = 0                          # live _download_part threads
        self._session: requests.Session | None = None   # kept so cancel() can close it
        self._research_cache: dict[int, dict] = {}        # idx -> {url: candidate}
        self._output_folder: str = _DEFAULTS["output_folder"]

        # Per-part skip: idxs the user ✕-ed on the download grid. The engine
        # consults this via the skip-check below at every cancellation
        # checkpoint, so ONE part can be aborted without touching the rest.
        # NOTE: always mutate with .add/.discard/.clear — never rebind — the
        # lambda captures this exact set object.
        self._skipped: set[int] = set()
        self._extra_futs: list = []       # futures for re-queued (un-skipped) parts
        self._run_ctx: dict | None = None # folder/max_dl/tracked-fn of the live run
        bd.set_skip_check(lambda i: i in self._skipped)

        self._progress: dict = {"status": "idle"}

    # ── File I/O ──────────────────────────────────────────────────────────────

    def open_bom_dialog(self) -> dict:
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=(
                "BOM Files (*.pdf;*.xlsx;*.xls;*.docx;*.md;*.txt)",
                "All Files (*.*)",
            ),
        )
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    def open_parts_dialog(self) -> dict:
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=(
                "Parts Files (*.csv;*.xlsx;*.xls;*.txt)",
                "All Files (*.*)",
            ),
        )
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    def browse_output_folder(self) -> dict:
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    def browse_catalog_folder(self) -> dict:
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    def browse_logo_file(self) -> dict:
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=(
                "Image Files (*.png;*.jpg;*.jpeg)",
                "All Files (*.*)",
            ),
        )
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    # ── BOM / Parts parsing ───────────────────────────────────────────────────

    def parse_file_from_path(self, path: str) -> dict:
        try:
            parts = bd.extract_parts_from_bom(path)
            return {"ok": True, "parts": parts, "filename": Path(path).name}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def parse_file_from_data(self, filename: str, data_b64: str) -> dict:
        try:
            raw = base64.b64decode(data_b64)
            ext = Path(filename).suffix.lower() or ".tmp"
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as f:
                f.write(raw)
                tmp = f.name
            try:
                parts = bd.extract_parts_from_bom(tmp)
                return {"ok": True, "parts": parts, "filename": filename}
            finally:
                try: os.unlink(tmp)
                except OSError: pass
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── Download job ──────────────────────────────────────────────────────────

    def start_download(self, parts: list[dict], options: dict) -> dict:
        if self._job_thread and self._job_thread.is_alive():
            return {"ok": False, "error": "A job is already running."}
        # Make sure NO thread from a previous (cancelled) job is still alive
        # before installing a fresh cancel event — otherwise those threads
        # would see "not cancelled" again and resume downloading.
        if not self._settle_old_job():
            return {"ok": False, "error":
                    "The previous job is still stopping — wait a few seconds and try again."}

        self._output_folder = options.get("output_folder", _DEFAULTS["output_folder"])

        self._skipped.clear()
        self._extra_futs = []
        self._run_ctx = None

        with self._log_lock:
            self._log.clear()
        with self._lock:
            self._progress = {
                "status":       "running",
                "total":        len(parts),
                "found":        0,
                "not_found":    0,
                "done":         0,
                "parts":        [
                    {
                        "idx":          i + 1,
                        "manufacturer": p.get("manufacturer", ""),
                        "part_number":  p.get("part_number", ""),
                        "description":  p.get("description", ""),
                        "status":       "pending",
                        "files":        [],
                    }
                    for i, p in enumerate(parts)
                ],
                "log":          [],
                "output_path":  None,
                "output_folder": self._output_folder,
            }

        self._job_thread = threading.Thread(
            target=self._run, args=(parts, options), daemon=True
        )
        self._job_thread.start()
        return {"ok": True}

    def _run(self, parts: list[dict], options: dict) -> None:
        folder   = Path(options.get("output_folder", _DEFAULTS["output_folder"]))
        max_dl   = int(options.get("max_per_part", 1))
        workers  = int(options.get("workers", 5))
        no_merge = bool(options.get("no_merge", False))
        bd.set_catalog_dirs([p.strip() for p in
                             str(options.get("catalog_folder", "")).split(";")
                             if p.strip()])

        def _on_event(ev: dict) -> None:
            t = ev.get("type")
            # Log events are by far the most frequent. Keep them OFF the main
            # progress lock entirely — append to a dedicated deque under its own
            # tiny lock. This is the change that stops the UI freezing: worker
            # threads logging dozens of lines/sec no longer fight get_progress()
            # for self._lock.
            if t == "log":
                with self._log_lock:
                    self._log.append(ev.get("message", ""))
                return
            with self._lock:
                if t == "part_start":
                    if ev["idx"] in self._skipped:
                        return          # user skipped it — don't flip to searching
                    for p in self._progress["parts"]:
                        if p["idx"] == ev["idx"]:
                            p["status"] = "searching"
                            break
                elif t == "part_done":
                    if ev["idx"] in self._skipped:
                        # Skipped mid-flight: skip_part() already set the status
                        # and counted it as done — just make sure it stays that
                        # way and discard whatever the aborted search returned.
                        for p in self._progress["parts"]:
                            if p["idx"] == ev["idx"]:
                                p["status"] = "skipped"
                                p["files"]  = []
                                break
                        return
                    st = ev.get("status", "not_found")
                    for p in self._progress["parts"]:
                        if p["idx"] == ev["idx"]:
                            p["status"] = st
                            p["files"]  = ev.get("files", [])
                            break
                    self._progress["done"] += 1
                    if st == "found":
                        self._progress["found"] += 1
                    else:
                        self._progress["not_found"] += 1

        bd.set_progress_callback(_on_event)
        bd.set_cancel_event(self._cancel_ev)
        self._session = _make_session()

        def _tracked_download(*args):
            """Wrap _download_part so we can count live worker threads —
            _settle_old_job() waits on this count before re-arming anything."""
            if args[0] in self._skipped:
                # Skipped while still queued: skip_part() already set the
                # status and counters — don't search, don't emit anything.
                return {"saved": [], "candidates": []}
            with self._lock:
                self._active_workers += 1
            try:
                return bd._download_part(*args)
            finally:
                with self._lock:
                    self._active_workers -= 1

        # Kept for requeue_part(): re-queued parts submit through the same
        # tracked wrapper with the same folder/quota as the main pass.
        self._run_ctx = {"folder": folder, "max_dl": max_dl,
                         "tracked": _tracked_download}

        # Use a manually managed executor so we can shutdown(wait=False) on cancel,
        # which lets in-flight network calls die naturally without blocking the UI.
        ex = ThreadPoolExecutor(max_workers=workers)
        self._executor = ex
        try:
            folder.mkdir(parents=True, exist_ok=True)
            part_results: dict[int, dict] = {}

            fut_map = {
                ex.submit(
                    _tracked_download,
                    i + 1,
                    p.get("manufacturer", ""),
                    p.get("part_number", ""),
                    folder,
                    max_dl,
                    self._session,
                    p.get("description", ""),
                ): i + 1
                for i, p in enumerate(parts)
            }

            for fut in as_completed(fut_map):
                if self._cancel_ev.is_set():
                    # Stop queued futures; running ones will exit at next _is_cancelled() check
                    for f in fut_map:
                        f.cancel()
                    break
                idx = fut_map[fut]
                try:
                    part_results[idx] = fut.result()
                except Exception as e:
                    # A crashed worker never emits part_done — surface the error
                    # and close the part out so its card doesn't sit on
                    # "searching" forever and the done-counter still completes.
                    part_results[idx] = {"saved": [], "candidates": []}
                    with self._log_lock:
                        self._log.append(f"    [Error] part #{idx:03d} failed: {e!r}")
                    with self._lock:
                        for p in self._progress["parts"]:
                            if p["idx"] == idx and p["status"] in ("pending", "searching"):
                                p["status"] = "not_found"
                                self._progress["done"] += 1
                                self._progress["not_found"] += 1
                                break

            # Wait for any parts the user skipped and then re-queued — their
            # futures aren't in fut_map, so as_completed() above ignores them.
            while not self._cancel_ev.is_set():
                with self._lock:
                    pending = [f for f in self._extra_futs if not f.done()]
                if not pending:
                    break
                futures_wait(pending, timeout=0.5)

        except Exception as e:
            with self._log_lock:
                self._log.append(f"FATAL ERROR: {e}")
            with self._lock:
                self._progress["status"] = "error"
            return
        finally:
            # ← KEY FIX: don't block — let running network calls die on their own
            ex.shutdown(wait=False)
            # Close the HTTP session so any blocked socket reads get an error immediately
            try:
                self._session.close()
            except Exception:
                pass
            bd.set_progress_callback(None)
            # DON'T clear the cancel event here — lingering threads need to see
            # _is_cancelled()==True so they stop instead of continuing to download
            # PDFs for other parts.  The event will be replaced with a fresh one
            # when research_part / auto_pick is called.

        # ── Cancelled: surface whatever was downloaded so far ─────────────────
        if self._cancel_ev.is_set():
            final_files = self._collect_files(folder, len(parts))
            with self._lock:
                self._progress["status"]       = "cancelled"
                self._progress["output_folder"] = str(folder)
                # Show partial results: mark any pending/searching parts as not_found
                for p in self._progress["parts"]:
                    if p["status"] in ("pending", "searching"):
                        p["status"] = "not_found"
            return

        # ── Second pass: recover search-engine-only misses ───────────────────
        # The parallel pass rate-limits DuckDuckGo into a global pause and trips
        # the distributor scrapers, which starves parts with no DirectProbe/
        # portal mirror (e.g. ABB). Re-run those misses ONE AT A TIME with the
        # engines forced — the same thing the manual 'Search again' button does.
        self._retry_misses(parts, folder, max_dl)

        # Cancellation may have arrived mid-retry — handle it like the main pass
        # instead of falling through and mislabelling the job 'done'.
        if self._cancel_ev.is_set():
            with self._lock:
                self._progress["status"]        = "cancelled"
                self._progress["output_folder"] = str(folder)
                for p in self._progress["parts"]:
                    if p["status"] in ("pending", "searching"):
                        p["status"] = "not_found"
            return

        # ── Normal completion ─────────────────────────────────────────────────
        if no_merge:
            with self._lock:
                self._progress["status"]       = "done"
                self._progress["output_path"]  = None
                self._progress["output_folder"] = str(folder)
            return

        with self._lock:
            self._progress["status"] = "merging"

        final_files = self._collect_files(folder, len(parts))
        with self._lock:
            skipped_set = {p["idx"] for p in self._progress["parts"]
                           if p.get("status") == "skipped"}
        final_files = {i: f for i, f in final_files.items() if i not in skipped_set}
        output = str(folder / "BOM_Manuals_Merged.pdf")
        bd.merge_pdfs(
            bom_stem="bom_manuals",
            parts=parts,
            part_files=final_files,
            skipped=skipped_set,
            output_path=output,
            custom_logo_path=options.get("custom_logo_path", ""),
        )

        with self._lock:
            self._progress["status"]       = "done"
            self._progress["output_path"]  = output
            self._progress["output_folder"] = str(folder)

    def _retry_misses(self, parts: list[dict], folder: Path, max_dl: int) -> None:
        """Re-run the parts that found nothing, ONE AT A TIME, with the search
        engines forced on.

        Why this exists: the parallel main pass funnels every DuckDuckGo query
        through one global lock, which rate-limits DDG into an escalating global
        pause, and it trips the Mouser/RS scrapers on bot-blocks for the whole
        run. Parts that depend purely on search engines (e.g. ABB — no
        DirectProbe/portal mirror) then get ZERO candidates and finish
        'not found', even though their datasheet/manual is reachable. Running
        them sequentially with force=True bypasses the self-imposed DDG pause and
        removes the lock contention that caused the rate-limit — i.e. exactly
        what the manual 'Search again' button already does, just automatic."""
        if self._cancel_ev.is_set():
            return
        with self._lock:
            misses = [p for p in self._progress["parts"]
                      if p.get("status") == "not_found"]
        if not misses:
            return

        bd.reset_search_throttle()      # clear DDG cooldown + scraper trips
        # Log-only callback: surface retry activity in the UI WITHOUT re-emitting
        # part_start/part_done (those would double-count done/found).
        def _log_only(ev: dict) -> None:
            if ev.get("type") == "log":
                with self._log_lock:
                    self._log.append(ev.get("message", ""))
        bd.set_progress_callback(_log_only)

        s = _make_session()             # main-pass session was already closed
        bd.tprint(f"    [Retry] re-checking {len(misses)} part(s) that found "
                  f"nothing — one at a time, search engines forced…")
        try:
            for p in misses:
                if self._cancel_ev.is_set():
                    break
                idx   = p["idx"]
                mfr   = p.get("manufacturer", "")
                model = p.get("part_number", "")
                desc  = p.get("description", "")
                with self._lock:
                    p["status"] = "searching"
                    self._active_workers += 1
                # A generous single-part budget. Setting any interactive deadline
                # also disables the 75s batch per-part cap inside _find_pdfs.
                bd.set_search_deadline(90)
                try:
                    res = bd._download_part(idx, mfr, model, folder, max_dl,
                                            s, desc, force=True)
                except Exception:
                    res = {"saved": []}
                finally:
                    bd.set_search_deadline(None)
                    with self._lock:
                        self._active_workers -= 1
                saved = res.get("saved") or []
                with self._lock:
                    if saved and not self._cancel_ev.is_set():
                        p["status"] = "found"
                        p["files"]  = [str(f) for f in saved]
                        self._progress["found"]     = self._progress.get("found", 0) + 1
                        self._progress["not_found"] = max(
                            0, self._progress.get("not_found", 0) - 1)
                    else:
                        p["status"] = "not_found"
        finally:
            bd.set_progress_callback(None)
            try: s.close()
            except Exception: pass

    def _collect_files(self, folder: Path, n_parts: int) -> dict:
        result = {}
        for idx in range(1, n_parts + 1):
            files = bd._scan_folder(folder, idx)
            if files:
                result[idx] = files
        return result

    # ── Progress polling ──────────────────────────────────────────────────────

    def get_progress(self) -> dict:
        # Snapshot the log from its own lock (cheap), then copy the rest of the
        # progress under the main lock. The two locks are never held at once, so
        # logging and polling can't deadlock or serialise against each other.
        with self._log_lock:
            log_snapshot = list(self._log)
        with self._lock:
            snap = copy.deepcopy(self._progress)
        if isinstance(snap, dict) and snap.get("status") != "idle":
            snap["log"] = log_snapshot
        return snap

    def cancel(self) -> dict:
        """Signal cancel, kill all queued (not-yet-started) parts immediately,
        and close the HTTP session so blocked socket reads fail fast."""
        self._cancel_ev.set()
        # Without this, queued futures only get cancelled when the NEXT running
        # future happens to complete — meanwhile the pool keeps starting new
        # parts. cancel_futures=True empties the queue right now.
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
        if self._session:
            try:
                self._session.close()
            except Exception:
                pass
        return {"ok": True}

    # ── Per-part skip / re-queue while the job is running ─────────────────────

    def skip_part(self, idx: int) -> dict:
        """Skip ONE part while the job is running. A pending part is never
        searched; a part currently searching aborts at its next checkpoint."""
        idx = int(idx)
        with self._lock:
            if self._progress.get("status") != "running":
                return {"ok": False, "error": "No active download job."}
            part = self._part_by_idx(idx)
            if not part:
                return {"ok": False, "error": "Unknown part index."}
            if part["status"] not in ("pending", "searching"):
                return {"ok": False, "error": "This part has already finished."}
            self._skipped.add(idx)
            part["status"] = "skipped"
            part["files"]  = []
            self._progress["done"] += 1
            mfr, model = part.get("manufacturer", ""), part.get("part_number", "")
        bd.tprint(f"    [Skip] #{idx:03d} {mfr} {model} — skipped by user")
        return {"ok": True}

    def requeue_part(self, idx: int) -> dict:
        """Put a skipped part back into the running job's queue ('Search' on
        a skipped card). Only works while the job is still running — after
        that, the Results page's 'Search again' takes over."""
        idx = int(idx)
        with self._lock:
            part = self._part_by_idx(idx)
            if not part:
                return {"ok": False, "error": "Unknown part index."}
            if part["status"] != "skipped":
                return {"ok": False, "error": "Only skipped parts can be re-queued."}
            if (self._progress.get("status") != "running"
                    or self._executor is None or self._run_ctx is None):
                return {"ok": False, "error":
                        "The job has finished — use 'Search again' on the Results page."}
            self._skipped.discard(idx)
            part["status"] = "pending"
            self._progress["done"] = max(0, self._progress["done"] - 1)
            mfr   = part.get("manufacturer", "")
            model = part.get("part_number", "")
            desc  = part.get("description", "")
            ctx   = self._run_ctx
        try:
            fut = self._executor.submit(ctx["tracked"], idx, mfr, model,
                                        ctx["folder"], ctx["max_dl"],
                                        self._session, desc)
        except RuntimeError:
            # Executor already shut down — the run ended between the check
            # above and the submit. Restore the skipped state.
            with self._lock:
                self._skipped.add(idx)
                part["status"] = "skipped"
                self._progress["done"] += 1
            return {"ok": False, "error":
                    "The job just finished — use 'Search again' on the Results page."}
        with self._lock:
            self._extra_futs.append(fut)
        bd.tprint(f"    [Re-queue] #{idx:03d} {mfr} {model} — searching again")
        return {"ok": True}

    # ── File access ───────────────────────────────────────────────────────────

    def open_in_viewer(self, path: str) -> dict:
        """Open any file in its default Windows app."""
        try:
            os.startfile(path)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_folder(self, path: str) -> dict:
        """Open the folder containing a file (or the folder itself) in Explorer."""
        target = str(Path(path).parent) if os.path.isfile(path) else path
        try:
            subprocess.Popen(["explorer", target])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def merge_selected(self, selected_indices: list) -> dict:
        """Merge only the PDFs whose indices the user selected.
        Uses the exact file paths stored in _progress — no folder scanning at all."""
        with self._lock:
            prog = copy.deepcopy(self._progress)

        if prog.get("status") not in ("cancelled", "done"):
            return {"ok": False, "error": "No completed or cancelled job to merge."}

        parts_prog = prog.get("parts", [])
        folder     = Path(prog.get("output_folder", ""))

        if not selected_indices:
            return {"ok": False, "error": "No parts selected."}

        selected_set = {int(i) for i in selected_indices}

        # Use the exact paths recorded at download time — never scan the folder
        part_files: dict = {}
        for part in parts_prog:
            idx = part["idx"]
            if idx not in selected_set:
                continue
            raw_paths = part.get("files") or []
            valid = [Path(f) for f in raw_paths if os.path.isfile(f)]
            if valid:
                part_files[idx] = valid

        if not part_files:
            return {
                "ok": False,
                "error": (
                    "None of the selected files could be found on disk. "
                    "They may have been moved or deleted since the download."
                ),
            }

        # Everything not in part_files → skipped → no divider, no placeholder
        skipped = {p["idx"] for p in parts_prog if p["idx"] not in part_files}

        bom_parts = [
            {
                "manufacturer": p.get("manufacturer", ""),
                "part_number":  p.get("part_number",  ""),
                "description":  p.get("description",  ""),
            }
            for p in parts_prog
        ]

        output = str(folder / "BOM_Manuals_Selected.pdf")
        try:
            bd.merge_pdfs(
                bom_stem="selected",
                parts=bom_parts,
                part_files=part_files,
                skipped=skipped,
                output_path=output,
                custom_logo_path=_load_cfg().get("custom_logo_path", ""),
            )
            with self._lock:
                self._progress["output_path"] = output
            return {"ok": True, "path": output, "merged": len(part_files)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_pdf_as(self, source_path: str) -> dict:
        if not os.path.isfile(source_path):
            return {"ok": False, "error": "Merged PDF not found."}
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="BOM_Manuals.pdf",
            file_types=("PDF Files (*.pdf)",),
        )
        if not result:
            return {"ok": False}
        dest = result[0] if isinstance(result, (list, tuple)) else result
        shutil.copy2(source_path, dest)
        return {"ok": True, "path": dest}

    # ── Re-search a single part after the run (GUI review feature) ───────────
    def _settle_old_job(self, timeout: float = 20.0) -> bool:
        """Make sure every thread from the previous job is REALLY dead, then
        install a brand-new cancel event for the next operation.

        The old (broken) version set the module-global cancel event to None
        after a 5s join — but lingering worker threads read that same global,
        so they'd suddenly see "not cancelled" and RESUME downloading the rest
        of the BOM the moment the user clicked 'Search again'.

        Returns True when it is safe to proceed, False if old threads are
        still winding down (caller should tell the user to retry shortly)."""
        # Keep the old event SET the whole time so lingering threads keep
        # seeing _is_cancelled()==True and exit at their next check.
        self._cancel_ev.set()
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
        if self._job_thread and self._job_thread.is_alive():
            self._job_thread.join(timeout=5)
            if self._job_thread.is_alive():
                return False

        deadline = time.monotonic() + timeout
        while True:
            with self._lock:
                if self._active_workers == 0:
                    break
            if time.monotonic() >= deadline:
                return False
            time.sleep(0.15)

        # All old threads are gone — NOW it's safe to arm a fresh, unset event.
        self._cancel_ev = threading.Event()
        bd.set_cancel_event(self._cancel_ev)
        self._executor = None
        # Run-time skips must not leak into post-run research / auto-pick —
        # a lingering idx would make _download_part abort instantly.
        self._skipped.clear()
        return True

    def _part_by_idx(self, idx: int):
        for p in self._progress.get("parts", []):
            if p["idx"] == int(idx):
                return p
        return None

    def research_part(self, idx: int, manufacturer: str = None,
                      part_number: str = None) -> dict:
        """Run the full search cascade again for ONE part and return the
        candidate URLs, annotated so the UI can show suggestions.
        Optional manufacturer/part_number override the stored values — for
        fixing BOM typos (e.g. Rittal 8205.521 → 8208.521) without re-running.
        Corrections are saved on the part so Keep / Auto-pick / the merged
        catalog all use the corrected identity."""
        with self._lock:
            status = self._progress.get("status")
        if status not in ("done", "cancelled"):
            return {"ok": False, "error": "Re-search is available after the job finishes."}
        part = self._part_by_idx(idx)
        if not part:
            return {"ok": False, "error": "Unknown part index."}

        with self._lock:
            if manufacturer is not None and manufacturer.strip() != "":
                part["manufacturer"] = manufacturer.strip()
            if part_number is not None and part_number.strip() != "":
                part["part_number"] = part_number.strip()
        mfr, model = part.get("manufacturer", ""), part.get("part_number", "")
        desc = part.get("description", "")
        if not self._settle_old_job():      # stop lingering threads first
            return {"ok": False, "error":
                    "Still stopping the previous downloads — wait a few seconds and try again."}
        bd.set_progress_callback(None)          # don't disturb finished-job counters
        bd.tprint(f"    [Re-search] {mfr} {model}: starting candidate search…")
        my_ev = self._cancel_ev                 # the event governing THIS op
        s = _make_session()

        # Ensure catalog sources are loaded/configured in case of standalone calls
        cfg = _load_cfg()
        bd.set_catalog_dirs([p.strip() for p in
                             str(cfg.get("catalog_folder", "")).split(";")
                             if p.strip()])

        # Count this op as a live worker: if the user closes the modal or starts
        # a re-search for ANOTHER part, _settle_old_job() cancels this op and
        # WAITS for it to exit before installing a fresh (unset) cancel event —
        # so this op can never be accidentally "un-cancelled" and keep running.
        with self._lock:
            self._active_workers += 1
        bd.set_search_deadline(120)             # bounded — DDG rate-limits can otherwise
                                                # silently grind for many minutes
        try:
            cands = bd._find_pdfs(model, mfr, s, desc, force=True) or []
        except Exception as e:
            return {"ok": False, "error": str(e)}
        finally:
            bd.set_search_deadline(None)
            with self._lock:
                self._active_workers -= 1
            try: s.close()
            except Exception: pass

        if my_ev.is_set():                      # closed / superseded mid-search
            bd.tprint(f"    [Re-search] {mfr} {model}: aborted by user")
            return {"ok": False, "error": "Search cancelled.", "cancelled": True}

        # Check catalog matches and prepend them to candidate list
        cat_hits = []
        try:
            for name, ref in bd._catalog_lookup(model, mfr, desc, s):
                url = ref.get("url") or f"catalog://{ref.get('id') or name}"
                cat_hits.append({
                    "url": url,
                    "title": name,
                    "referer": "Catalog fallback",
                    "ref": ref
                })
        except Exception as e:
            bd.tprint(f"    [Re-search] catalog lookup failed: {e}")

        # Combine cands, prioritizing catalog results
        seen_urls = set()
        combined_cands = []
        for c in cat_hits:
            if c["url"] not in seen_urls:
                seen_urls.add(c["url"])
                combined_cands.append(c)
        for c in cands:
            if c["url"] not in seen_urls:
                seen_urls.add(c["url"])
                combined_cands.append(c)

        self._research_cache[int(idx)] = {c["url"]: c for c in combined_cands}
        out = []
        for i, c in enumerate(combined_cands):
            if c.get("ref"):
                out.append({
                    "url":       c["url"],
                    "title":     c["title"],
                    "domain":    "Catalog fallback",
                    "kind":      "manual" if bd._name_doc_type(c["url"], c["title"]) == "manual" else "datasheet",
                    "official":  True,
                    "suggested": i == 0,
                })
            else:
                kind = bd._name_doc_type(c["url"], c["title"]) or ""
                out.append({
                    "url":       c["url"],
                    "title":     c["title"],
                    "domain":    bd._dom(c["url"]),
                    "kind":      kind if kind in ("manual", "datasheet") else "",
                    "official":  bd._is_own_mfr_domain(c["url"], mfr),
                    "suggested": i == 0,
                })
        return {"ok": True, "part": {"idx": part["idx"], "manufacturer": mfr,
                                     "part_number": model},
                "candidates": out}

    def keep_candidate(self, idx: int, url: str) -> dict:
        """Download ONE user-chosen candidate into the output folder, strip/
        reject protection, verify, and mark the part as found."""
        part = self._part_by_idx(idx)
        if not part:
            return {"ok": False, "error": "Unknown part index."}
        cand = self._research_cache.get(int(idx), {}).get(url)
        if not cand:
            # Pasted/unknown URL — derive a usable title from the last
            # non-empty path segment (handles trailing slashes, no .pdf ext)
            seg = [s for s in url.split("?")[0].split("/") if s]
            cand = {"url": url, "title": (seg[-1] if seg else "document"),
                    "referer": None}
        folder = Path(self._progress.get("output_folder") or
                      _load_cfg()["output_folder"])
        folder.mkdir(parents=True, exist_ok=True)
        mfr, model = part.get("manufacturer", ""), part.get("part_number", "")

        if not self._settle_old_job():
            return {"ok": False, "error":
                    "Still stopping the previous downloads — wait a few seconds and try again."}
        bd.set_progress_callback(None)
        dest = bd.candidate_dest(folder, int(idx), mfr, model, cand["title"])
        s = _make_session()
        with self._lock:
            self._active_workers += 1
        try:
            if cand.get("ref"):
                if not bd._fetch_catalog_pdf(cand["ref"], dest, s):
                    return {"ok": False, "error": "Failed to copy/download catalog file."}
            else:
                if not bd._write_pdf(url, dest, s, cand.get("referer")):
                    return {"ok": False, "error": "Download failed (not a valid PDF or blocked)."}
        finally:
            with self._lock:
                self._active_workers -= 1
            try: s.close()
            except Exception: pass

        ok, note = bd._ensure_unprotected(dest)
        if not ok:
            try: dest.unlink()
            except Exception: pass
            return {"ok": False, "error": f"Rejected: {note}"}

        dtype, _q, vnote = bd._verify_pdf(
            dest, model, mfr, bd._name_doc_type(url, cand["title"]),
            src_url=url, src_title=cand["title"],
            desc=part.get("description", ""))
        if dtype == "unreadable":
            try: dest.unlink()
            except Exception: pass
            return {"ok": False, "error": "The downloaded file is corrupt/unreadable."}

        with self._lock:
            was_found = part["status"] == "found"
            part["status"] = "found"
            part["files"]  = [str(dest)]
            if not was_found:
                self._progress["found"]     = self._progress.get("found", 0) + 1
                self._progress["not_found"] = max(0, self._progress.get("not_found", 0) - 1)
        warn = dtype in bd._REJECT_TYPES or dtype == "unknown"
        return {"ok": True, "file": str(dest), "doc_type": dtype,
                "note": vnote, "warning": warn}

    def upload_own_pdf(self, idx: int) -> dict:
        """Let the user pick a PDF from their own disk for ONE part — copies it
        into the output folder under the canonical name, verifies it, and marks
        the part as found. Mirrors keep_candidate with a local file as source."""
        part = self._part_by_idx(idx)
        if not part:
            return {"ok": False, "error": "Unknown part index."}
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=(
                "PDF Files (*.pdf)",
                "All Files (*.*)",
            ),
        )
        if not result:
            return {"ok": False, "cancelled": True}
        src = Path(result[0] if isinstance(result, (list, tuple)) else result)
        if not src.is_file():
            return {"ok": False, "error": "File not found."}
        try:
            with open(src, "rb") as f:
                if f.read(4) != b"%PDF":
                    return {"ok": False, "error": "That file is not a valid PDF."}
        except Exception as e:
            return {"ok": False, "error": str(e)}

        folder = Path(self._progress.get("output_folder") or
                      _load_cfg()["output_folder"])
        folder.mkdir(parents=True, exist_ok=True)
        mfr, model = part.get("manufacturer", ""), part.get("part_number", "")
        dest = bd.candidate_dest(folder, int(idx), mfr, model, src.stem)
        try:
            shutil.copy2(src, dest)
        except Exception as e:
            return {"ok": False, "error": f"Could not copy the file: {e}"}

        ok, note = bd._ensure_unprotected(dest)
        if not ok:
            try: dest.unlink()
            except Exception: pass
            return {"ok": False, "error": f"Rejected: {note}"}

        dtype, _q, vnote = bd._verify_pdf(
            dest, model, mfr, None,
            src_url=str(src), src_title=src.stem,
            desc=part.get("description", ""))
        if dtype == "unreadable":
            try: dest.unlink()
            except Exception: pass
            return {"ok": False, "error": "The file is corrupt/unreadable."}

        with self._lock:
            was_found = part["status"] == "found"
            part["status"] = "found"
            part["files"]  = [str(dest)]
            if not was_found:
                self._progress["found"]     = self._progress.get("found", 0) + 1
                self._progress["not_found"] = max(0, self._progress.get("not_found", 0) - 1)
        warn = dtype in bd._REJECT_TYPES or dtype == "unknown"
        return {"ok": True, "file": str(dest), "doc_type": dtype,
                "note": vnote, "warning": warn}

    def add_part(self, after_idx: int, manufacturer: str = "",
                 part_number: str = "", description: str = "") -> dict:
        """Insert a new part into the finished results, right after the row
        whose idx == after_idx (-1 appends at the end, 0 inserts at the top).
        Every part is renumbered so the merged PDF follows the new order."""
        if not (part_number or "").strip():
            return {"ok": False, "error": "Part number is required."}
        with self._lock:
            if self._progress.get("status") not in ("done", "cancelled"):
                return {"ok": False, "error": "Available after the job finishes."}
            parts = self._progress.get("parts", [])
            after_idx = int(after_idx)
            if after_idx < 0:
                pos = len(parts)
            elif after_idx == 0:
                pos = 0
            else:
                pos = next((i + 1 for i, p in enumerate(parts)
                            if p["idx"] == after_idx), len(parts))
            new_part = {
                "idx":          0,          # assigned in the renumber below
                "manufacturer": (manufacturer or "").strip(),
                "part_number":  part_number.strip(),
                "description":  (description or "").strip(),
                "status":       "not_found",
                "files":        [],
            }
            parts.insert(pos, new_part)
            # Renumber sequentially; remap the research cache to the new idxs
            old_cache, new_cache = self._research_cache, {}
            for i, p in enumerate(parts):
                if p is not new_part and p["idx"] in old_cache:
                    new_cache[i + 1] = old_cache[p["idx"]]
                p["idx"] = i + 1
            self._research_cache = new_cache
            self._progress["total"]     = len(parts)
            self._progress["not_found"] = self._progress.get("not_found", 0) + 1
            snap = copy.deepcopy(self._progress)
        return {"ok": True, "new_idx": pos + 1, "parts": snap["parts"],
                "found":     snap.get("found", 0),
                "not_found": snap.get("not_found", 0),
                "total":     snap.get("total", 0)}

    def auto_pick(self, idx: int) -> dict:
        """Let the verified download pipeline pick the best document for ONE
        part (same logic as the main run). Falls back to returning the
        candidate list when nothing passes verification."""
        with self._lock:
            status = self._progress.get("status")
        if status not in ("done", "cancelled"):
            return {"ok": False, "error": "Available after the job finishes."}
        part = self._part_by_idx(idx)
        if not part:
            return {"ok": False, "error": "Unknown part index."}
        folder = Path(self._progress.get("output_folder") or
                      _load_cfg()["output_folder"])
        folder.mkdir(parents=True, exist_ok=True)
        mfr, model = part.get("manufacturer", ""), part.get("part_number", "")

        if not self._settle_old_job():      # stop lingering threads first
            return {"ok": False, "error":
                    "Still stopping the previous downloads — wait a few seconds and try again."}
        bd.set_progress_callback(None)
        bd.tprint(f"    [Auto-pick] {mfr} {model}: searching and verifying…")
        my_ev = self._cancel_ev
        s = _make_session()
        with self._lock:
            self._active_workers += 1
        bd.set_search_deadline(150)
        try:
            res = bd._download_part(int(idx), mfr, model, folder, 1, s,
                                    part.get("description", ""), force=True)
        except Exception as e:
            return {"ok": False, "error": str(e)}
        finally:
            bd.set_search_deadline(None)
            with self._lock:
                self._active_workers -= 1
            try: s.close()
            except Exception: pass

        if my_ev.is_set():                      # closed / superseded mid-search
            bd.tprint(f"    [Auto-pick] {mfr} {model}: aborted by user")
            return {"ok": False, "error": "Search cancelled.", "cancelled": True}

        self._research_cache[int(idx)] = {c["url"]: c for c in res.get("candidates", [])}
        saved = res.get("saved") or []
        if saved:
            with self._lock:
                was_found = part["status"] == "found"
                part["status"] = "found"
                part["files"]  = [str(f) for f in saved]
                if not was_found:
                    self._progress["found"]     = self._progress.get("found", 0) + 1
                    self._progress["not_found"] = max(0, self._progress.get("not_found", 0) - 1)
            return {"ok": True, "file": str(saved[0])}
        cands = [{"url": c["url"], "title": c["title"], "domain": bd._dom(c["url"]),
                  "kind": (bd._name_doc_type(c["url"], c["title"]) or "")
                          if (bd._name_doc_type(c["url"], c["title"]) or "") in ("manual", "datasheet") else "",
                  "official": bd._is_own_mfr_domain(c["url"], mfr),
                  "suggested": i == 0}
                 for i, c in enumerate(res.get("candidates", []))]
        return {"ok": False, "error": "Nothing passed verification — pick manually below.",
                "candidates": cands}

    def cancel_research(self) -> dict:
        """Abort any in-flight re-search / auto-pick / keep download.
        Called when the user closes the Search-again modal. Safe to call when
        nothing is running. Never touches an active main download job."""
        if self._job_thread and self._job_thread.is_alive():
            return {"ok": False, "error": "A download job is running — use Stop instead."}
        self._cancel_ev.set()
        return {"ok": True}

    def open_url(self, url: str) -> dict:
        """Open a URL in the system browser (candidate preview)."""
        try:
            import webbrowser
            webbrowser.open(url)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── PDF Merge Tool ────────────────────────────────────────────────────────
    def receive_merge_pdf(self, name: str, data_b64: str) -> dict:
        """Receive a PDF as base64, save to a session temp file, return its path.
        Used as a fallback when drag-and-drop doesn't expose the file system path."""
        try:
            raw = base64.b64decode(data_b64)
            if not raw[:4] == b'%PDF':
                return {"ok": False, "error": "Not a valid PDF file."}
            # Lazy-init a per-session temp directory
            if not hasattr(self, '_merge_tmp') or self._merge_tmp is None:
                self._merge_tmp = Path(tempfile.mkdtemp(prefix="bom_merge_"))
            tmp_dir: Path = self._merge_tmp
            # Sanitize filename, avoid collisions
            safe = ''.join(c if c.isalnum() or c in '-_. ' else '_' for c in name)[:120]
            out = tmp_dir / safe
            counter = 0
            while out.exists():
                counter += 1
                out = tmp_dir / f"{Path(safe).stem}_{counter}.pdf"
            out.write_bytes(raw)
            return {"ok": True, "path": str(out)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_pdfs_merge_dialog(self) -> dict:
        """Open a multi-select file dialog for picking PDFs to merge."""
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=(
                "PDF Files (*.pdf)",
                "All Files (*.*)",
            ),
        )
        if result:
            return {"ok": True, "paths": list(result)}
        return {"ok": False}

    def browse_merge_output(self) -> dict:
        """Save dialog to choose the output path for the merged PDF."""
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="Merged.pdf",
            file_types=("PDF Files (*.pdf)",),
        )
        if not result:
            return {"ok": False}
        dest = result[0] if isinstance(result, (list, tuple)) else result
        return {"ok": True, "path": dest}

    def merge_pdf_files(self, pdf_paths: list, output_path: str) -> dict:
        """Merge a list of PDF files into one ordered PDF."""
        from pypdf import PdfWriter, PdfReader
        try:
            if not pdf_paths:
                return {"ok": False, "error": "No PDF files provided."}

            writer = PdfWriter()
            total_pages = 0
            errors: list[str] = []

            for path in pdf_paths:
                if not os.path.isfile(path):
                    errors.append(f"Not found: {Path(path).name}")
                    continue
                try:
                    reader = PdfReader(str(path))
                    if reader.is_encrypted:
                        try:
                            if int(reader.decrypt("")) == 0:
                                errors.append(f"{Path(path).name}: password-protected — skipped")
                                continue
                        except Exception:
                            errors.append(f"{Path(path).name}: password-protected — skipped")
                            continue
                    for page in reader.pages:
                        writer.add_page(page)
                    total_pages += len(reader.pages)
                except Exception as e:
                    errors.append(f"{Path(path).name}: {e}")

            if total_pages == 0:
                msg = "No pages could be read from the provided PDFs."
                if errors:
                    msg += " Errors: " + "; ".join(errors)
                return {"ok": False, "error": msg}

            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)
            with open(out, "wb") as f:
                writer.write(f)

            mb = out.stat().st_size / (1024 * 1024)
            return {
                "ok":      True,
                "path":    str(out),
                "pages":   total_pages,
                "size_mb": round(mb, 2),
                "errors":  errors,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── Settings ──────────────────────────────────────────────────────────────

    def load_settings(self) -> dict:
        return {"ok": True, "settings": _load_cfg()}

    def save_settings(self, settings: dict) -> dict:
        _save_cfg({**_DEFAULTS, **settings})
        return {"ok": True}

    def get_default_output(self) -> dict:
        return {"ok": True, "path": _load_cfg()["output_folder"]}