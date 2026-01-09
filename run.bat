@echo off

REM Change to project directory
cd /d %~dp0

REM Activate virtual environment
call .venv\Scripts\activate.bat

REM Run your Python app
python -m src.app.app

REM Keep CMD open after execution
pause
