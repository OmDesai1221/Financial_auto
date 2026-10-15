@echo off
cd /d "%~dp0app"
set PYEXE=%~dp0venv\Scripts\python.exe
if not exist "%PYEXE%" goto novenv
start "Astra Monthly Update Tool - DO NOT CLOSE" cmd /k ""%PYEXE%" app.py"
timeout /t 3 /nobreak >nul
start "" http://localhost:5050
goto end

:novenv
echo Could not find venv\Scripts\python.exe - run setup again (see README).
pause

:end
