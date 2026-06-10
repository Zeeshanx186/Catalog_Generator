@echo off
:: build.bat — One-click build: produces installer + portable zip
:: Run from the project root with: build.bat

setlocal enabledelayedexpansion
title BOM Downloader — Build

echo.
echo  ============================================================
echo   BOM Manual Downloader — Build Script
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
echo  [1/4] Installing Python dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 ( echo  [ERROR] pip install failed. & pause & exit /b 1 )
pip install pyinstaller --quiet
echo       Done.

:: ── Create assets folder + placeholder icon if missing ───────────────────────
if not exist assets mkdir assets
if not exist assets\icon.ico (
    echo  [INFO]  No icon found at assets\icon.ico — using default.
    echo  [INFO]  Replace assets\icon.ico with your own 256x256 .ico file.
    :: Copy Python's own icon as a placeholder
    for %%p in ("%LOCALAPPDATA%\Programs\Python\Python3*\python.exe") do (
        if exist "%%p" ( copy "%%~dpp..\*.ico" assets\icon.ico >nul 2>&1 )
    )
)

:: ── PyInstaller ───────────────────────────────────────────────────────────────
echo.
echo  [2/4] Building executable with PyInstaller...
if exist dist\BomDownloader rmdir /s /q dist\BomDownloader
if exist build rmdir /s /q build

pyinstaller BomDownloader.spec --noconfirm
if errorlevel 1 ( echo  [ERROR] PyInstaller build failed. & pause & exit /b 1 )
echo       Executable built: dist\BomDownloader\BomDownloader.exe

:: ── Portable ZIP ─────────────────────────────────────────────────────────────
echo.
echo  [3/4] Creating portable ZIP...
if exist "dist\BomDownloader-Portable.zip" del "dist\BomDownloader-Portable.zip"
powershell -Command "Compress-Archive -Path 'dist\BomDownloader\*' -DestinationPath 'dist\BomDownloader-Portable.zip'"
if errorlevel 1 ( echo  [WARN]  Could not create ZIP ^(PowerShell missing?^). Skipping. )
echo       Portable ZIP: dist\BomDownloader-Portable.zip

:: ── Inno Setup installer (optional) ─────────────────────────────────────────
echo.
echo  [4/4] Building Windows installer...
set ISCC=
for %%p in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
) do ( if exist "%%~p" set ISCC=%%~p )

if defined ISCC (
    "!ISCC!" installer.iss
    if errorlevel 1 ( echo  [WARN]  Inno Setup compile failed. Installer skipped. )
    else ( echo       Installer: dist\BomDownloader-Setup-1.0.0.exe )
) else (
    echo  [INFO]  Inno Setup not found — skipping installer.
    echo  [INFO]  Download free at: https://jrsoftware.org/isinfo.php
)

:: ── Summary ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
echo   Build complete!
echo.
echo   Portable :  dist\BomDownloader-Portable.zip
if exist "dist\BomDownloader-Setup-1.0.0.exe" (
    echo   Installer:  dist\BomDownloader-Setup-1.0.0.exe
)
echo  ============================================================
echo.
pause
