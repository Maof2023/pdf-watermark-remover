@echo off
rem ============================================================
rem  Build Pdf2WordCleaner.exe (onefile, windowed) with PyInstaller
rem  Double-click this file in the project directory to build.
rem ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] python not found in PATH.
  echo         Install Python 3.9+ and check "Add Python to PATH".
  pause
  exit /b 1
)

echo [1/2] Ensuring PyInstaller is installed...
python -m pip install --quiet pyinstaller
if errorlevel 1 (
  echo [ERROR] Failed to install PyInstaller. Check your network or use:
  echo         python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
  pause
  exit /b 1
)

echo [2/2] Building exe ...
python -m PyInstaller --noconfirm --onefile --windowed --name Pdf2WordCleaner ^
  --collect-all pdf2docx --collect-all pymupdf --collect-all pymupdf_layout app.py
if errorlevel 1 (
  echo [ERROR] Build failed. See messages above.
) else (
  echo.
  echo [OK] Build finished: dist\Pdf2WordCleaner.exe
)
pause
