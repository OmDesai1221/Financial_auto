@echo off
cd /d "%~dp0app"
start "Astra Monthly Update Tool - DO NOT CLOSE" cmd /k python app.py
timeout /t 3 /nobreak >nul
start "" http://localhost:5050
