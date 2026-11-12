@echo off
setlocal EnableExtensions

REM START_OIC.bat
REM Place this file in the same folder as oic_collector.py

cd /d "%~dp0" || exit /b 1
title OIC XMLA Collector

REM --------------------------------------------------
REM Check application
REM --------------------------------------------------

if not exist "oic_collector.py" (
    echo [ERROR] oic_collector.py not found.
    pause
    exit /b 1
)

REM --------------------------------------------------
REM Open Windows Terminal in fullscreen
REM --------------------------------------------------

if /I not "%~1"=="--inside" (
    where wt.exe >nul 2>&1
    if not errorlevel 1 (
        wt.exe -w new --fullscreen -d "%CD%" cmd.exe /c "call ""%~f0"" --inside"
        if not errorlevel 1 exit /b 0
    )
)

REM --------------------------------------------------
REM Find Python automatically
REM --------------------------------------------------

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
        set "PYTHON=py -3"
        goto :RUN
    )
)

where python.exe >nul 2>&1
if not errorlevel 1 (
    python --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=python"
        goto :RUN
    )
)

echo [ERROR] Python 3 not found.
echo Install Python 3 or create a .venv in this folder.
pause
exit /b 1

REM --------------------------------------------------
REM Run OIC Terminal UI
REM --------------------------------------------------

:RUN
%PYTHON% "oic_collector.py"

if errorlevel 1 (
    echo.
    echo [ERROR] OIC XMLA Collector exited with an error.
    pause
)

endlocal
exit /b 0