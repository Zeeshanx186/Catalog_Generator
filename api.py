"""
api.py — Python API exposed to the frontend via pywebview's JS bridge.

Every public method is callable from JavaScript as:
    const result = await window.pywebview.api.method_name(args)
All methods return plain dicts (JSON-serialisable).
"""

from __future__ import annotations

import base64
import copy
import json
import os
import shutil
import subprocess
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
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
    s = requests.Session()
    s.headers.update({"User-Agent": bd.BROWSER_UA})
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=20,
        pool_maxsize=20,
        max_retries=requests.adapters.Retry(total=1, backoff_factor=0.3),
    )
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    return s


# ── Main API class ────────────────────────────────────────────────────────────

class API:
    """Exposed to JavaScript as window.pywebview.api"""

    def __init__(self) -> None:
        self.window: webview.Window | None = None

        self._lock      = threading.Lock()
        self._cancel_ev = threading.Event()
        self._job_thread: threading.Thread | None = None
        self._session: requests.Session | None = None   # kept so cancel() can close it
        self._output_folder: str = _DEFAULTS["output_folder"]

        self._progress: dict = {"status": "idle"}

    # ── File I/O ──────────────────────────────────────────────────────────────

    def open_bom_dialog(self) -> dict:
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=(
                "BOM Files (*.pdf;*.xlsx;*.xls;*.md;*.txt)",
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

        self._cancel_ev.clear()
        self._output_folder = options.get("output_folder", _DEFAULTS["output_folder"])

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

        def _on_event(ev: dict) -> None:
            with self._lock:
                t = ev.get("type")
                if t == "part_start":
                    for p in self._progress["parts"]:
                        if p["idx"] == ev["idx"]:
                            p["status"] = "searching"
                            break
                elif t == "part_done":
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
                elif t == "log":
                    log = self._progress["log"]
                    log.append(ev.get("message", ""))
                    if len(log) > 300:
                        self._progress["log"] = log[-300:]

        bd.set_progress_callback(_on_event)
        bd.set_cancel_event(self._cancel_ev)
        self._session = _make_session()

        # Use a manually managed executor so we can shutdown(wait=False) on cancel,
        # which lets in-flight network calls die naturally without blocking the UI.
        ex = ThreadPoolExecutor(max_workers=workers)
        try:
            folder.mkdir(parents=True, exist_ok=True)
            part_results: dict[int, dict] = {}

            fut_map = {
                ex.submit(
                    bd._download_part,
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
                except Exception:
                    part_results[idx] = {"saved": [], "candidates": []}

        except Exception as e:
            with self._lock:
                self._progress["status"] = "error"
                self._progress["log"].append(f"FATAL ERROR: {e}")
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
            bd.set_cancel_event(None)

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
        output = str(folder / "BOM_Manuals_Merged.pdf")
        bd.merge_pdfs(
            bom_stem="bom_manuals",
            parts=parts,
            part_files=final_files,
            skipped=set(),
            output_path=output,
        )

        with self._lock:
            self._progress["status"]       = "done"
            self._progress["output_path"]  = output
            self._progress["output_folder"] = str(folder)

    def _collect_files(self, folder: Path, n_parts: int) -> dict:
        result = {}
        for idx in range(1, n_parts + 1):
            files = bd._scan_folder(folder, idx)
            if files:
                result[idx] = files
        return result

    # ── Progress polling ──────────────────────────────────────────────────────

    def get_progress(self) -> dict:
        with self._lock:
            return copy.deepcopy(self._progress)

    def cancel(self) -> dict:
        """Signal cancel AND close the HTTP session so blocked socket reads fail fast."""
        self._cancel_ev.set()
        if self._session:
            try:
                self._session.close()
            except Exception:
                pass
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
