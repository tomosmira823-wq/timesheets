@echo off
REM Runs the hosted (online) version on this PC for testing: http://localhost:8000
cd /d "%~dp0"
python -m pip install -r web\requirements.txt >nul 2>&1
set DATA_ROOT=%~dp0web\data
python -m uvicorn web.server:app --host 127.0.0.1 --port 8000
pause
