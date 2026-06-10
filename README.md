# BOM Manual Downloader

A desktop application that automatically downloads datasheets and manuals for every part in a Bill of Materials, then merges them into a single, bookmarked PDF package — ready for engineering handoff.

![App Screenshot](docs/screenshot.png)

---

## What it does

Give it a BOM file. It reads every part number, searches manufacturer websites and the web, downloads the matching PDF manual, and stitches them into one ordered document — with a cover page, per-part dividers, and "Not Found" placeholders for anything it couldn't locate.

```
BOM.xlsx  →  [auto-search + download]  →  BOM_Manuals.pdf
```

---

## Features

- **Multiple BOM formats** — PDF, Excel (`.xlsx` / `.xls`), Markdown (`.md`)
- **Smart search** — DirectProbe hits manufacturer CDNs first, falls back to DuckDuckGo
- **Concurrent downloads** — configurable worker threads (default: 5)
- **Merged PDF output** — cover page, per-part divider pages, "Not Found" placeholders
- **Partial results** — cancel mid-run and still merge whatever was downloaded
- **Checkbox selection** — choose exactly which found manuals to include in the merge
- **Standalone PDF merger** — combine any PDFs from your filesystem into one file, with drag-and-drop reordering
- **Optional paid search backends** — plug in API keys for Google CSE, Bing, Nexar, Mouser, Farnell, and more
- **Dark UI** — native desktop window via pywebview

---

## Screenshots

| Step | Description |
|------|-------------|
| **1 — Input** | Load a BOM file or drag it onto the window |
| **2 — Parts** | Review extracted parts before downloading |
| **3 — Download** | Live progress with per-part status |
| **4 — Results** | Checkboxes to select which manuals to merge |

---

## Installation

### Requirements

- Python 3.10 or newer
- Windows (pywebview uses the system WebView2 / Edge runtime)

### Steps

```bash
# 1. Clone the repo
git clone https://github.com/your-username/BomDownloader.git
cd BomDownloader

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python main.py
```

### `requirements.txt`

```
pywebview>=4.4.1,<5.0
requests>=2.31.0
beautifulsoup4>=4.12.0
pdfplumber>=0.10.0
pypdf>=3.17.0
reportlab>=4.0.0
openpyxl>=3.1.0
ddgs>=6.1.0
```

---

## Usage

### Step 1 — Load your BOM

Click **Load BOM** and select a PDF, Excel, or Markdown file. The app will extract manufacturer and part number columns automatically.

Alternatively, click **Load Parts List** to provide a plain CSV/text file of parts directly.

### Step 2 — Review parts

Check the extracted parts list. You can rename or remove rows before proceeding.

### Step 3 — Download

Click **Start Download**. The app searches for each part's datasheet concurrently and shows live status (Found / Not Found / Searching…). You can **Cancel** at any time — partial results are preserved.

### Step 4 — Merge

Once complete, every found manual is pre-checked in the results table.

- **Uncheck** any manuals you don't want to include
- Click **Merge Selected (N)** to build the PDF
- Use **Open Merged PDF** or **Save PDF As…** to access the output

---

## Output PDF structure

```
[Cover page]          — lists all parts with Found ✓ / Not Found ✗ / Skipped markers
  [Divider — Part 001]  RITTAL  8108245
  [Manual pages]
  [Divider — Part 002]  HONEYWELL  FC-PDB-0824P
  [Manual pages]
  [Divider — Part 003]  MOXA  MB3270I-T
  [Placeholder]         — "DOCUMENT NOT FOUND — please source manually"
  ...
```

---

## PDF Merge Tool

The sidebar **Merge PDFs** tool lets you combine arbitrary PDFs from your filesystem — independently of any BOM workflow.

1. Click the **Layers icon** in the sidebar to open the tool
2. Drag PDFs onto the drop zone or click **Add More** to browse
3. Reorder files with the ↑ ↓ buttons
4. Set an output path and click **Merge PDFs**

---

## Optional search backends

The app works out of the box with DuckDuckGo and direct manufacturer CDN probing (both free, no key needed). For higher hit rates or larger BOMs, add API keys in `bom_downloader.py`:

| Backend | Variable | Free tier |
|---------|----------|-----------|
| Nexar | `NEXAR_CLIENT_ID` / `NEXAR_CLIENT_SECRET` | 1,000 / month |
| Google Custom Search | `GOOGLE_CSE_KEY` / `GOOGLE_CSE_CX` | 100 / day |
| Bing Search | `BING_API_KEY` | 1,000 / month (Azure) |
| Mouser | `MOUSER_API_KEY` | Free |
| Farnell / element14 | `FARNELL_API_KEY` | Free |
| Serper | `SERPER_API_KEY` | 2,500 / month |
| SerpAPI | `SERPAPI_KEY` | Paid |
| Tavily | `TAVILY_API_KEY` | 1,000 / month |
| Exa | `EXA_API_KEY` | 1,000 / month |
| Brave Search | `BRAVE_API_KEY` | $3 / 1,000 |

Keys are optional — leave any field blank to skip that backend.

---

## Project structure

```
BomDownloader/
├── main.py              # App entry point, creates pywebview window
├── api.py               # Python ↔ JavaScript bridge (all pyapi() calls)
├── bom_downloader.py    # Core: BOM parsing, search, download, PDF merge
├── index.html           # Single-file frontend (HTML + CSS + JS)
└── requirements.txt
```

| File | Responsibility |
|------|---------------|
| `main.py` | Boots pywebview, resolves paths for dev vs. PyInstaller |
| `api.py` | Exposes Python methods to the JS frontend; manages download state and progress |
| `bom_downloader.py` | BOM parsers (PDF/Excel/MD), search engine clients, concurrent downloader, PDF writer |
| `index.html` | All UI — 4-step workflow, settings panel, PDF merge tool, dark theme |

---

## Settings

Accessible via the ⚙ icon in the app:

| Setting | Default | Description |
|---------|---------|-------------|
| Output folder | `Documents\BOM Manuals` | Where downloaded files and merged PDFs are saved |
| Worker threads | 5 | Parallel download connections |
| Skip merging | Off | Save files only, skip the merge step |

---

## Supported BOM formats

| Format | How parts are extracted |
|--------|------------------------|
| **PDF** | `pdfplumber` text extraction, regex patterns for MFR / P/N columns |
| **Excel** | `openpyxl` column detection by header keywords |
| **Markdown** | Table row parsing |

---

## Building a standalone executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name BomDownloader main.py
```

The app uses `sys._MEIPASS` to resolve bundled assets when frozen.

---

## License

MIT
