@echo off
setlocal
cd /d "%~dp0"

python daily_logger.py
if errorlevel 1 (
  echo.
  echo If this failed, install requirements with:
  echo   pip install openpyxl
  echo.
  pause
)
