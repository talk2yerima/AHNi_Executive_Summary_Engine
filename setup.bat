@echo off
REM ============================================================
REM  setup.bat  -  First-time setup for AHNi Executive Service Engine
REM  Double-click or run from Command Prompt.
REM ============================================================
echo.
echo ============================================
echo  AHNi Executive Service Engine  -  Setup
echo ============================================
echo.

REM -- 1. Check Python 3.12 --
py -3.12 --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python 3.12 not found.
    echo         Install Python 3.12 from https://www.python.org
    echo         IMPORTANT: Check "Add Python to PATH" during install.
    pause & exit /b 1
)
echo [1/5] Python 3.12 OK:
py -3.12 --version

REM -- 2. Create virtual environment --
echo.
IF EXIST ".venv" (
    IF NOT EXIST ".venv\Scripts\activate.bat" (
        echo [2/5] Existing .venv is invalid - removing and recreating...
        rmdir /s /q .venv
    )
)
IF NOT EXIST ".venv" (
    echo [2/5] Creating virtual environment with Python 3.12...
    py -3.12 -m venv .venv
    IF %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create virtual environment.
        pause & exit /b 1
    )
    echo       .venv created.
) ELSE (
    echo [2/5] .venv already exists - skipping.
)

REM -- 3. Install packages --
echo.
echo [3/5] Installing packages...
call .venv\Scripts\activate.bat
python -m pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] pip install failed. Check your internet connection.
    pause & exit /b 1
)
echo       Packages installed.

REM -- 4. Create folders and .env --
echo.
echo [4/5] Preparing folders and config...
IF NOT EXIST "output" mkdir output
IF NOT EXIST "cache"  mkdir cache
IF NOT EXIST "logs"   mkdir logs
echo       output\  cache\  logs\  ready.

IF NOT EXIST ".env" (
    copy .env.example .env >nul
    echo       .env created from .env.example
    echo.
    echo  *** ACTION REQUIRED ***
    echo  Open .env and fill in your credentials before continuing:
    echo    DHIS2_URL, DHIS2_USER, DHIS2_PASS
    echo    AZURE_CONNECTION_STRING, AZURE_CONTAINER_NAME, AZURE_EXCEL_BLOB_NAME
) ELSE (
    echo       .env already exists.
)

REM -- 5. Install the Windows Service --
echo.
echo [5/5] Installing AHNi Executive Service Engine...
echo       (This will open a UAC prompt to grant Administrator rights)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_service.ps1"

echo.
echo ============================================
echo  Setup complete!
echo ============================================
echo.
echo  The service is now running as: AHNi_Executive_Service_Engine
echo  It will run a full pull daily at midnight and
echo  incremental pulls every 3 hours in between.
echo.
echo  To check status:
echo    nssm\nssm.exe status AHNi_Executive_Service_Engine
echo.
echo  To view live logs:
echo    powershell -Command "Get-Content logs\AESE.log -Tail 50 -Wait"
echo.
pause
