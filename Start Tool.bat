@echo off
cd /d "%~dp0app"
if exist "%~dp0venv\Scripts\python.exe" (
    start "Astra Monthly Update Tool - DO NOT CLOSE" cmd /k ""%~dp0venv\Scripts\python.exe" app.py"
) else (
    echo Could not find venv\Scripts\python.exe — run setup again (see README).
    pause
    exit /b 1
)
timeout /t 3 /nobreak >nul
start "" http://localhost:5050
