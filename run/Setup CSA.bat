@echo off
setlocal
cd /d "%~dp0.."

where py >nul 2>nul
if not errorlevel 1 goto python_found
where python >nul 2>nul
if not errorlevel 1 goto python_found

where winget >nul 2>nul
if errorlevel 1 (
    echo Python 3 was not found, and Windows Package Manager is unavailable.
    echo Install Python 3 from https://www.python.org/downloads/
    echo Enable "Add Python to PATH", then run this setup file again.
    pause
    exit /b 1
)

echo Installing Python 3 with Windows Package Manager...
winget install --id Python.Python.3.13 --exact --source winget
if errorlevel 1 (
    echo Python installation did not complete.
    pause
    exit /b 1
)

:python_found
where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=py"
    goto create_environment
)
set "PYTHON=python"

:create_environment
echo Using Python:
%PYTHON% --version
if not exist ".venv\Scripts\python.exe" (
    echo Creating the CSA virtual environment...
    %PYTHON% -m venv .venv
    if errorlevel 1 (
        echo Could not create the virtual environment.
        pause
        exit /b 1
    )
)

echo Installing CSA requirements...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
)

echo CSA setup is complete.
echo You can now double-click: run\Launch CSA.bat
pause
