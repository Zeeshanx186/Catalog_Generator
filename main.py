"""
main.py — Entry point for the BOM Manual Downloader desktop app.
"""

import sys
import os
from pathlib import Path

import webview
from api import API


def resource(relative: str) -> str:
    """Resolve a path that works both in dev and when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    return str(base / relative)


def main() -> None:
    api = API()

    # Find index.html — check ui/ subfolder first, then the root folder
    for candidate in ["ui/index.html", "index.html"]:
        html_path = resource(candidate)
        if Path(html_path).exists():
            break
    else:
        raise FileNotFoundError(
            "index.html not found. Place it at ui/index.html (or the app root)."
        )

    html_content = Path(html_path).read_text(encoding="utf-8")

    window = webview.create_window(
        title="BOM Manual Downloader",
        html=html_content,
        js_api=api,
        width=1280,
        height=820,
        min_size=(980, 660),
        background_color="#080d18",
        text_select=False,
    )

    api.window = window

    webview.start(
        debug="--debug" in sys.argv,
        private_mode=False,
        storage_path=str(Path(os.path.expandvars(r"%APPDATA%")) / "BomDownloader"),
    )


if __name__ == "__main__":
    main()
