@echo off
REM Runs the app straight from the Python source (no build) - handy for testing changes.
cd /d "%~dp0"
python -m pip install pywebview openpyxl fpdf2 >nul 2>&1
python timesheet_app.py
if errorlevel 1 pause
