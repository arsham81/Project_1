@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 (
 pause
 exit /b 1
)
start "" http://127.0.0.1:8501
.venv\Scripts\python app.py
pause
