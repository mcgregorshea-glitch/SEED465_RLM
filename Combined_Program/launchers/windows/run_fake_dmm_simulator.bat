@echo off
setlocal
cd /d "%~dp0..\..\sim"

set VENV=venv

if not exist "%VENV%\Scripts\python.exe" (
    echo [SETUP] Virtual environment not found. Creating it now...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment. Is Python in your PATH?
        pause
        exit /b 1
    )
    echo [SETUP] Installing dependencies...
    "%VENV%\Scripts\pip" install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] pip install failed. Check requirements.txt and your internet connection.
        pause
        exit /b 1
    )
    echo [SETUP] Setup complete.
    echo.
)

REM dmm_manager.py's DMM_LIST hard-codes six instruments at IDs 20-25 (VINV,
REM VINP, IINP, VSYS, SAUX, SINV). Each needs its own simulator process bound
REM to 127.0.0.<id>:5025 — one loopback octet per instrument, fixed port 5025.
REM Point the real app's "DMM IP Prefix" field (Sender panel > DMM section) at
REM 127.0.0 (no port) to reach all six. No seed-control-center code changes needed.
echo Starting 6 Fake DMM Simulators (127.0.0.20-25:5025)...
start "Fake DMM - VINV (20)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 20 --name VINV
start "Fake DMM - VINP (21)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 21 --name VINP
start "Fake DMM - IINP (22)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 22 --name IINP
start "Fake DMM - VSYS (23)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 23 --name VSYS
start "Fake DMM - SAUX (24)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 24 --name SAUX
start "Fake DMM - SINV (25)" "%VENV%\Scripts\python.exe" fake_dmm.py --id 25 --name SINV
