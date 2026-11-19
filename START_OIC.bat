@echo off
setlocal EnableExtensions
REM Launch the existing Textual UI in a dedicated PowerShell console.
cd /d "%~dp0" || exit /b 1
title OIC XMLA Collector

if not exist "oic_collector.py" (
    echo [ERROR] oic_collector.py not found.
    pause
    exit /b 1
)

set "PYTHON="
set "PY_ARGS="
if exist ".venv\Scripts\python.exe" (
    set "PYTHON=%CD%\.venv\Scripts\python.exe"
    goto :RUN
)
if exist "venv\Scripts\python.exe" (
    set "PYTHON=%CD%\venv\Scripts\python.exe"
    goto :RUN
)
if exist "python.exe" (
    set "PYTHON=%CD%\python.exe"
    goto :RUN
)
where py.exe >nul 2>&1
if not errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=py.exe"
        set "PY_ARGS=-3"
        goto :RUN
    )
)
where python.exe >nul 2>&1
if not errorlevel 1 (
    python --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=python.exe"
        goto :RUN
    )
)
echo [ERROR] Python 3 not found.
echo Install Python 3 or create a .venv in this folder.
pause
exit /b 1

:RUN
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PowerShell not found.
    pause
    exit /b 1
)
start "OIC XMLA Collector" /D "%CD%" powershell.exe -NoLogo -NoProfile -NoExit -Command "& '%PYTHON%' %PY_ARGS% 'oic_collector.py'"
if errorlevel 1 (
    echo [ERROR] PowerShell could not be started.
    pause
    exit /b 1
)
endlocal
exit /b 0
