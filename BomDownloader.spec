# BomDownloader.spec
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect ALL reportlab files (binaries, data, hidden imports)
rl_datas, rl_binaries, rl_hiddenimports = collect_all('reportlab')

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
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.utils',
        'reportlab.lib.colors',
        'reportlab.lib.styles',
        'reportlab.lib.units',
        'reportlab.pdfbase',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.pdfbase._fontdata',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfutils',
        'reportlab.graphics',
        'reportlab.graphics.shapes',
        'reportlab.platypus',
        'openpyxl',
        'openpyxl.styles',
        'requests',
        'requests.adapters',
        'bs4',
        'ddgs',
        'webview',
        'webview.platforms.winforms',
        'webview.platforms.mshtml',
        'json',
        'threading',
        'concurrent.futures',
        'tempfile',
        'base64',
        'shutil',
        'subprocess',
        'pathlib',
    ] + rl_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BomDownloader',
)