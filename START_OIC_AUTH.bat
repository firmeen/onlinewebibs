@echo off
setlocal
cd /d "%~dp0"

if not exist "oic_auth.py" (
    echo ERROR: oic_auth.py not found.
    pause
    exit /b 1
)

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" oic_auth.py
) else (
    python oic_auth.py
)

set "RC=%ERRORLEVEL%"
pause
exit /b %RC%