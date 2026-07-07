@echo off
:: build.bat — One-click build: produces the portable folder + portable zip
:: Run from the project root with: build.bat
::
:: IMPORTANT: always build via BomDownloader.spec (this script does).
:: A bare "pyinstaller main.py" produces a broken exe that exits silently,
:: because the ui\ folder and the SSL runtime hook are only wired up in the spec.

setlocal enabledelayedexpansion
title BOM Downloader - Build

echo.
echo  ============================================================
echo   BOM Manual Downloader - Build Script
echo  ============================================================
echo.

:: ── Check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Install Python 3.10+ and add to PATH.
    pause & exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo  Python: %%i

:: ── Check / install dependencies ─────────────────────────────────────────────
echo.
echo  [1/3] Installing Python dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 ( echo  [ERROR] pip install failed. & pause & exit /b 1 )
pip install pyinstaller --quiet
echo       Done.

:: ── PyInstaller ───────────────────────────────────────────────────────────────
echo.
echo  [2/3] Building executable with PyInstaller...
if exist dist\BomDownloader rmdir /s /q dist\BomDownloader
if exist build rmdir /s /q build

pyinstaller BomDownloader.spec --noconfirm
if errorlevel 1 ( echo  [ERROR] PyInstaller build failed. & pause & exit /b 1 )
echo       Executable built: dist\BomDownloader\BomDownloader.exe

:: ── Portable ZIP ─────────────────────────────────────────────────────────────
echo.
echo  [3/3] Creating portable ZIP...
if exist "dist\BomDownloader-Portable.zip" del "dist\BomDownloader-Portable.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\BomDownloader' -DestinationPath 'dist\BomDownloader-Portable.zip'"
if errorlevel 1 ( echo  [WARN]  Could not create ZIP ^(PowerShell missing?^). Skipping. )
echo       Portable ZIP: dist\BomDownloader-Portable.zip

:: ── Summary ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
echo   Build complete!
echo.
echo   Portable folder:  dist\BomDownloader\   (run BomDownloader.exe)
echo   Portable ZIP   :  dist\BomDownloader-Portable.zip
echo.
echo   Target machine needs: Windows 10/11 with the WebView2
echo   runtime (preinstalled on Win 11 and any PC with MS Edge).
echo  ============================================================
echo.
pause
