@echo off
echo ================================
echo   Python Virtual Env Setup
echo ================================

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    pause
    exit /b 1
)

REM Create virtual environment
if not exist venv (
    echo [INFO] Creating virtual environment...
    python -m venv venv
) else (
    echo [INFO] Virtual environment already exists.
)

REM Activate venv
call venv\Scripts\activate

REM Install requirements
if exist requirements.txt (
    echo [INFO] Installing dependencies...
    pip install -r requirements.txt
) else (
    echo [WARNING] requirements.txt not found.
)

echo ================================
echo   Setup completed successfully
echo ================================
pause
