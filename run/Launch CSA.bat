@echo off
setlocal
cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
    echo CSA virtual environment was not found.
    echo Create it with: py -m venv .venv
    echo Then install dependencies with: .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

".venv\Scripts\python.exe" app.py
if errorlevel 1 pause
