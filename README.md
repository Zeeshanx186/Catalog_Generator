<div align="center">

<img src="docs/banner.png" alt="BOM Manual Downloader" width="100%"/>

# BOM Manual Downloader

**Automatically fetches datasheets for every part in your BOM and merges them into one clean PDF.**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://github.com)
[![License](https://img.shields.io/badge/License-MIT-22c55e?style=for-the-badge)](LICENSE)
[![pywebview](https://img.shields.io/badge/UI-pywebview-f59e0b?style=for-the-badge)](https://pywebview.flowrl.com)

<br/>

*Load a BOM → parts are searched & downloaded → everything merges into one ordered PDF*

</div>

---

## ✨ Features

<table>
<tr>
<td width="50%">

### 📥 Smart Downloading
- Reads **PDF, Excel (.xlsx/.xls) and Markdown** BOMs
- Probes manufacturer CDNs directly before falling back to web search
- **Concurrent downloads** — up to 5 parallel workers by default
- Live per-part progress (Searching → Found / Not Found)

</td>
<td width="50%">

### 📄 Professional PDF Output
- **Cover page** listing every part with Found ✓ / Not Found ✗
- Per-part **divider pages** for easy navigation
- Placeholder pages for anything that couldn't be located
- Cancel mid-run and still merge what was found

</td>
</tr>
<tr>
<td width="50%">

### ✅ Selective Merging
- Results page shows **checkboxes** for every found manual
- Pre-selects all found parts — uncheck anything you don't want
- Uses exact download paths, never stale files from previous runs
- Outputs a clean `BOM_Manuals_Selected.pdf`

</td>
<td width="50%">

### 🔧 Standalone PDF Merger
- Sidebar tool to merge **any PDFs** from your filesystem
- Drag-and-drop file list with ↑ ↓ reordering
- Click to browse or drag files directly onto the drop zone
- Works independently of any BOM workflow

</td>
</tr>
</table>

---

## 🖥️ Screenshots

| Input & Parts | Live Download | Results & Merge |
|:---:|:---:|:---:|
| ![Input](docs/screen-input.png) | ![Download](docs/screen-download.png) | ![Results](docs/screen-results.png) |
| Load BOM, review parts | Watch per-part progress live | Select manuals, merge to PDF |

---

## 🚀 Getting Started

### Prerequisites

- Python **3.10+**
- Windows (uses the system WebView2 / Edge runtime via pywebview)

### Install & Run

```bash
# Clone
git clone https://github.com/your-username/BomDownloader.git
cd BomDownloader

# Install dependencies
pip install -r requirements.txt

# Launch
python main.py
```

### Dependencies

| Package | Purpose |
|---------|---------|
| `pywebview` | Native desktop window |
| `pdfplumber` | Extract text from PDF BOMs |
| `pypdf` | Merge and manipulate PDFs |
| `reportlab` | Generate cover, divider & placeholder pages |
| `openpyxl` | Read Excel BOMs |
| `requests` + `beautifulsoup4` | Download and scrape manufacturer sites |
| `ddgs` | DuckDuckGo search (free, no key needed) |

---

## 📖 How to Use

```
1. Load BOM      →  PDF, Excel, or Markdown file
2. Review Parts  →  Confirm extracted manufacturers & part numbers
3. Download      →  Watch progress — cancel any time
4. Merge         →  Check the manuals you want → click Merge Selected
```

> **Cancelled a run?** No problem. The Results page still shows everything that was downloaded with checkboxes, so you can merge partial results right away.

---

## 📂 Output Structure

```
BOM_Manuals_Selected.pdf
│
├── [Cover Page]              all 50 parts listed — Found ✓ / Not Found ✗ / Skipped
│
├── [Divider]  001  RITTAL  8108245
├── [Manual pages ...]
│
├── [Divider]  002  HONEYWELL  FC-PDB-0824P
├── [Manual pages ...]
│
├── [Divider]  003  MOXA  MB3270I-T
└── [Placeholder]             "DOCUMENT NOT FOUND — please source manually"
```

---

## 🔑 Optional Search Backends

Works out of the box with **DuckDuckGo + DirectProbe** (both free, no setup). For higher hit rates or larger BOMs, add API keys in `bom_downloader.py`:

<details>
<summary><b>Click to expand API key configuration</b></summary>

<br/>

| Provider | Variable(s) | Free Tier | Link |
|----------|-------------|-----------|------|
| **Nexar** | `NEXAR_CLIENT_ID` / `NEXAR_CLIENT_SECRET` | 1,000 / month | [nexar.com/api](https://nexar.com/api) |
| **Google CSE** | `GOOGLE_CSE_KEY` / `GOOGLE_CSE_CX` | 100 / day | [console.cloud.google.com](https://console.cloud.google.com) |
| **Bing Search** | `BING_API_KEY` | 1,000 / month | [Azure Portal](https://portal.azure.com) |
| **Mouser** | `MOUSER_API_KEY` | Free | [mouser.com/api-hub](https://www.mouser.com/api-hub) |
| **Farnell / element14** | `FARNELL_API_KEY` | Free | [partner.element14.com](https://partner.element14.com) |
| **Serper** | `SERPER_API_KEY` | 2,500 / month | [serper.dev](https://serper.dev) |
| **Tavily** | `TAVILY_API_KEY` | 1,000 / month | [app.tavily.com](https://app.tavily.com) |
| **Exa** | `EXA_API_KEY` | 1,000 / month | [dashboard.exa.ai](https://dashboard.exa.ai) |
| **Brave Search** | `BRAVE_API_KEY` | $3 / 1,000 | [brave.com/search/api](https://brave.com/search/api) |

Leave any field blank to skip that backend. A built-in circuit breaker disables a backend after 3 consecutive failures so one bad API key doesn't slow down the whole run.

</details>

---

## 🗂️ Project Structure

```
BomDownloader/
├── main.py              ← Entry point — creates the pywebview window
├── api.py               ← Python ↔ JS bridge (all pyapi() calls from the frontend)
├── bom_downloader.py    ← Core engine: BOM parsing, search, download, PDF generation
├── index.html           ← Complete frontend (HTML + CSS + JS, single file)
└── requirements.txt
```

---

## ⚙️ Settings

| Setting | Default | Description |
|---------|---------|-------------|
| Output folder | `Documents\BOM Manuals` | Where files and merged PDFs are saved |
| Worker threads | `5` | Parallel download connections |
| Skip merging | `off` | Download only, skip the merge step |

---

## 📦 Build a Standalone EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name BomDownloader main.py
```

The app detects `sys._MEIPASS` automatically when frozen, so no extra configuration needed.

---

## 📜 License

Released under the [MIT License](LICENSE). Free to use, modify, and distribute.

---

<div align="center">
  <sub>Built with Python · pywebview · pypdf · reportlab · DuckDuckGo</sub>
</div>
