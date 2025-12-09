@echo off
echo ================================================
echo  PDF UTILITY - DEPENDENCY INSTALLER
echo ================================================
echo.
echo This script will install all required Python libraries:
echo   - pypdf (for PDF operations)
echo   - Pillow (for image handling)
echo   - msoffcrypto-tool (for Excel encryption)
echo.
echo ================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python from https://www.python.org/
    echo.
    pause
    exit /b 1
)

echo Python detected. Proceeding with installation...
echo.

REM Upgrade pip first
echo [1/4] Upgrading pip...
python -m pip install --upgrade pip
echo.

REM Install pypdf
echo [2/4] Installing pypdf...
pip install pypdf
echo.

REM Install Pillow
echo [3/4] Installing Pillow...
pip install Pillow
echo.

REM Install msoffcrypto-tool
echo [4/4] Installing msoffcrypto-tool...
pip install msoffcrypto-tool
echo.

echo ================================================
echo  INSTALLATION COMPLETE!
echo ================================================
echo.
echo All dependencies have been installed successfully.
echo You can now run pdf_util.py
echo.
pause