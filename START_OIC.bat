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
REM Launch the main UI in maximized PowerShell (never Windows Terminal)
REM --------------------------------------------------

if /I not "%~1"=="--inside" (
    where powershell.exe >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Windows PowerShell is required.
        pause
        exit /b 1
    )
    powershell.exe -NoLogo -NoProfile -NonInteractive -Command "$bat = ('%~f0').Replace([char]39, [string][char]39+[char]39); $script = '& ' + [char]39 + $bat + [char]39 + ' --inside'; $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script)); Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList ('-NoLogo -NoProfile -NoExit -EncodedCommand ' + $enc) -WorkingDirectory '%CD%' -WindowStyle Maximized"
    exit /b %ERRORLEVEL%
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