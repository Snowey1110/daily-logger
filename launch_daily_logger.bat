@echo off
setlocal
cd /d "%~dp0"

where pythonw >nul 2>nul
if not errorlevel 1 (
  start "" /b pythonw daily_logger.py
  exit /b 0
)

where py >nul 2>nul
if not errorlevel 1 (
  start "" /b pyw -3 daily_logger.py
  exit /b 0
)

echo Could not find pythonw/pyw for GUI-only launch.
echo Install Python for Windows and ensure pythonw is on PATH.
echo.
pause
