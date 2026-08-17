@echo off
setlocal
cd /d "%~dp0"

REM Vérification du script Python dans src/ ou racine
if exist "src\unlock.py" (
    set "PY_SCRIPT=src\unlock.py"
) else (
    set "PY_SCRIPT=unlock.py"
)

REM Vérification du script PowerShell dans src/ ou racine
if exist "src\unlock.ps1" (
    set "PS_SCRIPT=src\unlock.ps1"
) else (
    set "PS_SCRIPT=unlock.ps1"
)

REM 1. Si un environnement virtuel .venv ou Python est installé, on l'utilise
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "%PY_SCRIPT%" %*
    goto :end
)

where python >nul 2>&1
if %errorlevel% equ 0 (
    python "%PY_SCRIPT%" %*
    goto :end
)

REM 2. Si Python N'EST PAS installé, bascule automatiquement sur PowerShell natif Windows (100% sans Python)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0%PS_SCRIPT%" %*

:end
if "%~1"=="" (
    pause
)
