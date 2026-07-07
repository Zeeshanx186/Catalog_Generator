# BomDownloader.spec — PyInstaller build recipe for the portable app.
#
# Build with:  pyinstaller BomDownloader.spec --noconfirm
# (or just run build.bat)
#
# Key portability fixes baked in here — do NOT build with a bare
# "pyinstaller main.py" command, it will miss all of this:
#   * datas=('ui','ui')            — bundles ui/index.html; without it the
#                                    windowed exe dies instantly with a silent
#                                    FileNotFoundError.
#   * runtime_hooks=rthook_ssl.py  — points requests/httpx/curl_cffi at the
#                                    bundled certifi CA file on machines where
#                                    cert env vars are unset/broken.
#   * upx=False                    — UPX-packed DLLs trip antivirus and crash
#                                    on some machines; not worth the size win.
import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect ALL reportlab files (binaries, data, hidden imports)
rl_datas, rl_binaries, rl_hiddenimports = collect_all('reportlab')

# Optional icon — build works with or without one
_icon = 'assets/icon.ico' if os.path.exists('assets/icon.ico') else None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[] + rl_binaries,
    datas=[
        ('ui', 'ui'),
    ] + rl_datas,
    hiddenimports=[
        'api',
        'bom_downloader',
        'pdfplumber',
        'pdfplumber.utils',
        'pypdf',
        'pypdf.filters',
        'openpyxl',
        'openpyxl.styles',
        'docx',                              # python-docx — Word BOM support
        'requests',
        'requests.adapters',
        'certifi',
        'bs4',
        'ddgs',
        'primp',                             # HTTP engine used by ddgs 9.x
        'webview',
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',    # WebView2 — the actual renderer
    ] + rl_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['rthook_ssl.py'],
    excludes=[
        'tkinter', 'matplotlib', 'numpy', 'scipy',
        'pandas', 'cv2', 'torch',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BomDownloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='BomDownloader',
)
