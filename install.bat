@echo off
setlocal enabledelayedexpansion

:: =========================================================================
::  MS Word Redline Creator v1.0 - Source Install Script
::  Author: Jack Wang
::
::  Use this if you want to run from source instead of building an .exe.
::  This script:
::   1. Checks for Python
::   2. Creates a virtual environment
::   3. Installs all dependencies
::   4. Creates a desktop shortcut (optional)
::   5. Launches the application
:: =========================================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo   MS Word Redline Creator v1.0 - Source Install
echo   Author: Jack Wang
echo ============================================================
echo.

:: ----- Step 1: Check Python -----------------------------------------------
echo [1/4] Checking Python installation...
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://www.python.org
    echo Make sure to check "Add Python to PATH" during install.
    goto :fail
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo       Found: %PYVER%

:: ----- Step 2: Create virtual environment ---------------------------------
echo.
echo [2/4] Setting up virtual environment...
if not exist "venv" (
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        goto :fail
    )
    echo       Created venv/
) else (
    echo       venv/ already exists, reusing.
)

:: Activate
call venv\Scripts\activate.bat

:: ----- Step 3: Install dependencies ---------------------------------------
echo.
echo [3/4] Installing dependencies...
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: Failed to install one or more dependencies.
    echo        Check your internet connection and try again.
    goto :fail
)
echo       All dependencies installed.

:: ----- Step 4: Create launcher script -------------------------------------
echo.
echo [4/4] Creating launcher...

:: Update the launch.bat to use the venv
(
    echo @echo off
    echo cd /d "%%~dp0"
    echo call venv\Scripts\activate.bat
    echo python gui.py
    echo if errorlevel 1 pause
) > "launch.bat"

echo       launch.bat updated to use virtual environment.

:: ----- Ask about desktop shortcut -----------------------------------------
echo.
set /p SHORTCUT="Create desktop shortcut? (Y/N): "
if /i "%SHORTCUT%"=="Y" (
    set DESKTOP=%USERPROFILE%\Desktop
    set SCRIPT_DIR=%~dp0

    powershell -NoProfile -Command ^
        "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\MS Word Redline Creator.lnk'); $s.TargetPath = '%SCRIPT_DIR%launch.bat'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Description = 'MS Word Redline Creator v1.0'; $s.Save()"

    if errorlevel 1 (
        echo       Could not create shortcut. You can run launch.bat directly.
    ) else (
        echo       Desktop shortcut created.
    )
)

:: ----- Done ---------------------------------------------------------------
echo.
echo ============================================================
echo   INSTALL COMPLETE
echo ============================================================
echo.
echo   To run the application:
echo     - Double-click launch.bat
if /i "%SHORTCUT%"=="Y" echo     - Or use the desktop shortcut
echo.
echo   To uninstall:
echo     - Delete this folder and the desktop shortcut
echo ============================================================
echo.

:: Offer to launch now
set /p LAUNCH="Launch the application now? (Y/N): "
if /i "%LAUNCH%"=="Y" (
    echo.
    echo Starting MS Word Redline Creator...
    start "" python gui.py
)

call venv\Scripts\deactivate.bat 2>nul
pause
exit /b 0

:fail
echo.
echo ============================================================
echo   INSTALL FAILED - See errors above.
echo ============================================================
echo.
call venv\Scripts\deactivate.bat 2>nul
pause
exit /b 1
