@echo off
setlocal
cd /d "%~dp0\virtual-journal-reader"

where pythonw >nul 2>nul
if not errorlevel 1 (
  start "" /b pythonw serve_reader.py %*
  exit /b 0
)

where py >nul 2>nul
if not errorlevel 1 (
  start "" /b pyw -3 serve_reader.py %*
  exit /b 0
)

where python >nul 2>nul
if not errorlevel 1 (
  start "" /min python serve_reader.py %*
  exit /b 0
)

echo Could not find Python. Install Python for Windows and ensure it is on PATH.
echo.
pause
