@echo off
setlocal EnableExtensions
cd /d "%~dp0" || exit /b 1

title OIC XMLA Collector
if not exist "oic_collector.py" (
    echo [ERROR] oic_collector.py not found.
    pause
    exit /b 1
)

REM Open a native PowerShell console only once.
if /I not "%~1"=="--inside" (
    set "PS="
    where pwsh.exe >nul 2>&1 && set "PS=pwsh.exe"
    if not defined PS (
        if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
            set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
        )
    )
    if not defined PS (
        echo [ERROR] PowerShell not found.
        pause
        exit /b 1
    )
    REM The temporary CMD launcher exits; only one visible PowerShell window stays.
    start "OIC XMLA Collector" "%PS%" -NoLogo -NoProfile -NoExit -Command "& '%~f0' --inside"
    if errorlevel 1 (
        echo [ERROR] PowerShell could not be opened.
        pause
        exit /b 1
    )
    exit /b 0
)

REM Prefer the project's own environment.
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

:RUN
if /I "%PYTHON%"=="py -3" (
    py -3 "oic_collector.py"
) else (
    "%PYTHON%" "oic_collector.py"
)
if errorlevel 1 (
    echo.
    echo [ERROR] OIC XMLA Collector exited with an error.
    pause
)
endlocal
exit /b 0
