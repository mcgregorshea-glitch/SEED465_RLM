@echo off
setlocal
cd /d "%~dp0..\.."

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

:loop
echo Launching SEED Control Center...
"%VENV%\Scripts\python.exe" src\main.py

echo.
echo Application exited.
echo Press any key to restart, or close this window to exit.
pause >nul
cls
goto loop
