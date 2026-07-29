# AI Agent Setup Guide — BOM Downloader & Catalog Generator

> **Note for AI Agent**: This guide is an automated setup protocol to initialize, restore, and verify the **BomDownloader** project on a new machine. Follow these instructions step-by-step using system commands.

---

## 1. System Requirements & Environment Check

* **Operating System**: Windows 10/11 (or Linux/macOS)
* **Python Version**: Python 3.10+ (Verify with `python --version`)
* **Git**: Installed and configured
* **Target Directory**: `d:\APP\BomDownloader` (or user's workspace path)

---

## 2. Setup Instructions for AI Agent

### Step 1: Obtain Project Files
If the repository is not yet cloned on the new laptop, run:
```powershell
git clone https://github.com/Zeeshanx186/Catalog_Generator.git
cd Catalog_Generator
```

If restoring from `BomDownloader_export.zip`:
* Extract the contents of `BomDownloader_export.zip` into the target project folder.
* Ensure data directories `manuals/`, `assets/`, `docs/`, `scratchpad/`, and `BOM.xlsx` are present in the project root.

---

### Step 2: Create Python Virtual Environment
Do **NOT** reuse old virtual environments from other operating systems or laptops. Create a fresh one:

```powershell
# Create virtual environment named 'myvenv'
python -m venv myvenv

# Activate environment (PowerShell)
.\myvenv\Scripts\Activate.ps1

# (Alternative: CMD)
# myvenv\Scripts\activate.bat
```

---

### Step 3: Install Required Dependencies
Install all Python requirements:

```powershell
# Upgrade pip
.\myvenv\Scripts\python.exe -m pip install --upgrade pip

# Install dependencies from requirements.txt
.\myvenv\Scripts\pip.exe install -r requirements.txt
```

*Note: If `playwright` is present in requirements, install browser binaries:*
```powershell
.\myvenv\Scripts\playwright.exe install chromium
```

---

### Step 4: Verify Project Imports & Dependencies
Run a sanity check script to verify core modules import cleanly without errors:

```powershell
.\myvenv\Scripts\python.exe -c "import api, bom_downloader, build_catalog_map; print('SUCCESS: Core modules loaded clean!')"
```

---

### Step 5: Run Application / Development Server

#### Option A: Running Python API / GUI Backend
```powershell
.\myvenv\Scripts\python.exe main.py
```

#### Option B: Rebuilding PyInstaller Executable (Optional)
If building standalone `.exe` binaries:
```powershell
# Using PyInstaller spec file
.\myvenv\Scripts\pyinstaller.exe BomDownloader.spec

# Or running build.bat script
.\build.bat
```

---

## 3. Verification & Troubleshooting Checklist for AI Agent

- [ ] **Check Python Path**: Ensure commands use `.\myvenv\Scripts\python.exe` so dependencies are loaded correctly.
- [ ] **Check Data Files**: Verify `BOM.xlsx` and `manuals/` folder exist in project root.
- [ ] **Check UI Assets**: Ensure `ui/index.html` exists and loads cleanly in browser.
- [ ] **Check Network & API Keys**: Verify internet connection for downloading manuals and check `.env` if API keys are configured.

---
*Setup protocol complete. The BOM Downloader application is now fully configured and ready for operation.*
