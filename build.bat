@echo off
setlocal enabledelayedexpansion

:: =========================================================================
::  MS Word Redline Creator v1.0 - One-Click Build Script
::  Author: Jack Wang
::
::  This script:
::   1. Checks for Python
::   2. Creates a virtual environment (if not present)
::   3. Installs all dependencies + PyInstaller
::   4. Builds a standalone .exe via PyInstaller
::   5. Copies the launcher batch file into the dist folder
::   6. Creates a ready-to-distribute ZIP package
:: =========================================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo   MS Word Redline Creator v1.0 - Build Script
echo   Author: Jack Wang
echo ============================================================
echo.

:: ----- Step 1: Check Python -----------------------------------------------
echo [1/6] Checking Python installation...
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://www.python.org
    goto :fail
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo       Found: %PYVER%

:: ----- Step 2: Create virtual environment ---------------------------------
echo.
echo [2/6] Setting up virtual environment...
if not exist "build_env" (
    python -m venv build_env
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        goto :fail
    )
    echo       Created build_env/
) else (
    echo       build_env/ already exists, reusing.
)

:: Activate the venv
call build_env\Scripts\activate.bat

:: ----- Step 3: Install dependencies ---------------------------------------
echo.
echo [3/6] Installing dependencies...
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: Failed to install requirements.
    goto :fail
)
pip install pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller.
    goto :fail
)
echo       All dependencies installed.

:: ----- Step 4: Build with PyInstaller -------------------------------------
echo.
echo [4/6] Building standalone executable with PyInstaller...
echo       This may take 1-3 minutes...

:: Clean previous build artifacts
if exist "build\RedlineCreator" rmdir /s /q "build\RedlineCreator"
if exist "dist\RedlineCreator" rmdir /s /q "dist\RedlineCreator"

pyinstaller redline_creator.spec --noconfirm --clean
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    echo        Check the output above for details.
    goto :fail
)
echo       Build complete.

:: ----- Step 5: Add extras to dist folder ----------------------------------
echo.
echo [5/6] Preparing distribution package...

:: Copy the launcher
(
    echo @echo off
    echo cd /d "%%~dp0"
    echo start "" "RedlineCreator.exe"
) > "dist\RedlineCreator\Launch Redline Creator.bat"

:: Copy README
(
    echo MS Word Redline Creator v1.0
    echo Author: Jack Wang
    echo.
    echo ============================================================
    echo.
    echo QUICK START:
    echo   Double-click "Launch Redline Creator.bat"
    echo   or run RedlineCreator.exe directly.
    echo.
    echo WHAT IT DOES:
    echo   Compares two Word document revisions and generates a new
    echo   .docx file with tracked changes ^(redlines^) showing every
    echo   difference. Optionally carries over review comments from
    echo   the earlier revision.
    echo.
    echo REQUIREMENTS:
    echo   - Microsoft Word ^(for highest-fidelity comparison^)
    echo   - Or use "Force XML" mode for basic comparison without Word
    echo.
    echo FILES:
    echo   RedlineCreator.exe  - Main application
    echo   Launch Redline Creator.bat - Convenience launcher
    echo.
    echo For full documentation, click the "? Help" button in the app.
    echo ============================================================
) > "dist\RedlineCreator\README.txt"

echo       Distribution folder ready: dist\RedlineCreator\

:: ----- Step 6: Create ZIP archive ----------------------------------------
echo.
echo [6/6] Creating ZIP package...

:: Use PowerShell to create the zip
set ZIPNAME=RedlineCreator_v1.0_Windows.zip
if exist "dist\%ZIPNAME%" del "dist\%ZIPNAME%"

powershell -NoProfile -Command ^
    "Compress-Archive -Path 'dist\RedlineCreator\*' -DestinationPath 'dist\%ZIPNAME%' -Force"

if errorlevel 1 (
    echo WARNING: Could not create ZIP archive.
    echo          You can manually zip the dist\RedlineCreator folder.
) else (
    echo       Package created: dist\%ZIPNAME%
)

:: ----- Done ---------------------------------------------------------------
echo.
echo ============================================================
echo   BUILD SUCCESSFUL
echo ============================================================
echo.
echo   Executable:  dist\RedlineCreator\RedlineCreator.exe
echo   Launcher:    dist\RedlineCreator\Launch Redline Creator.bat
echo   ZIP Package: dist\%ZIPNAME%
echo.
echo   To deploy: copy the ZIP to any Windows machine and extract.
echo   No Python installation required on the target machine.
echo ============================================================
echo.

call build_env\Scripts\deactivate.bat 2>nul
pause
exit /b 0

:fail
echo.
echo ============================================================
echo   BUILD FAILED - See errors above.
echo ============================================================
echo.
call build_env\Scripts\deactivate.bat 2>nul
pause
exit /b 1
