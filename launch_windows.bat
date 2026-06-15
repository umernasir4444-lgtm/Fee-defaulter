@echo off
setlocal
cd /d "%~dp0"
title Fee Defaulter Report Generator

:: ── Pick Python ───────────────────────────────────────────────────────────
set "CODEX_PY=C:\Users\umern\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%CODEX_PY%" (
    set "PY=%CODEX_PY%"
    goto :got_python
)
where py >nul 2>nul
if %errorlevel%==0 ( set "PY=py" & goto :got_python )
where python >nul 2>nul
if %errorlevel%==0 ( set "PY=python" & goto :got_python )

echo [ERROR] Python was not found. Install Python and try again.
pause & exit /b 1

:got_python
echo [OK] Using Python: %PY%

:: ── Ensure required packages are available ─────────────────────────────
"%PY%" -c "import openpyxl" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] Installing openpyxl...
    "%PY%" -m pip install openpyxl --quiet
)

"%PY%" -c "import cryptography" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] Installing cryptography for HTTPS support...
    "%PY%" -m pip install cryptography --quiet
)

:: ── HTTPS Setup ───────────────────────────────────────────────────────────
if not exist "cert.pem" (
    echo [INFO] Generating self-signed certificates for HTTPS...
    "%PY%" generate_certs.py
)

set "PROTOCOL=http"
if exist "cert.pem" (
    set "PROTOCOL=https"
)

:: ── Password Setup ────────────────────────────────────────────────────────
echo.
echo ======================================================================
echo SECURITY SETUP
echo ======================================================================
set /p APP_PASSWORD="Set a password for LAN access (press Enter to skip): "
echo ======================================================================
echo.

:: ── Start server, THEN open browser after a short delay ──────────────────
echo [INFO] Starting server on %PROTOCOL%://0.0.0.0:8765
echo        Other users on your network can access it using your IP address.
echo        Browser will open automatically. Press Ctrl+C to stop.
echo.

:: Open browser in background after 3-second delay
start "" /b cmd /c "ping 127.0.0.1 -n 4 >nul && start %PROTOCOL%://127.0.0.1:8765"

:: Run server (this line BLOCKS until you Ctrl+C)
"%PY%" launcher.py

echo.
echo [INFO] Server stopped.
pause