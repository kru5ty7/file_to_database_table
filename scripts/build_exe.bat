@echo off
REM ========================================
REM File to Database Converter - Build Script
REM Creates standalone Windows executable
REM ========================================

echo.
echo ========================================
echo Building File to Database Converter
echo ========================================
echo.

REM Change to project root directory
cd /d "%~dp0\.."

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.10+ from python.org
    pause
    exit /b 1
)

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    echo.
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller
        pause
        exit /b 1
    )
    echo.
)

REM Verify all dependencies are installed
echo Checking dependencies...
python -c "import pandas, pyodbc, openpyxl, cryptography" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Some dependencies are missing!
    echo Installing dependencies from requirements.txt...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

echo Dependencies OK
echo.

REM Clean previous builds
echo Cleaning previous build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo.

REM Build the executable
echo ========================================
echo Building executable with PyInstaller...
echo ========================================
echo.
echo This may take 2-5 minutes...
echo.

pyinstaller FileToDBConverter.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
    echo Check the error messages above for details.
    echo Common issues:
    echo   - Missing dependencies
    echo   - Syntax errors in code
    echo   - Insufficient disk space
    echo.
    echo Check build\FileToDBConverter\warn-FileToDBConverter.txt for warnings
    echo.
    pause
    exit /b 1
)

REM Check if executable was created
if not exist "dist\FileToDBConverter.exe" (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
    echo Executable was not created!
    echo Check the error messages above.
    echo.
    pause
    exit /b 1
)

REM Get file size
for %%A in ("dist\FileToDBConverter.exe") do set "filesize=%%~zA"
set /a filesizeMB=%filesize% / 1048576

echo.
echo ========================================
echo BUILD SUCCESSFUL!
echo ========================================
echo.
echo Executable created: dist\FileToDBConverter.exe
echo File size: %filesizeMB% MB
echo.
echo You can now run the application by double-clicking:
echo   dist\FileToDBConverter.exe
echo.
echo To distribute:
echo   1. Copy dist\FileToDBConverter.exe to target machine
echo   2. Ensure SQL Server ODBC Driver 17+ is installed
echo   3. Run the executable - no Python needed!
echo.
echo Note: The executable includes Python runtime and all dependencies.
echo.

REM Ask if user wants to test the executable
echo.
set /p test="Do you want to test the executable now? (Y/N): "
if /i "%test%"=="Y" (
    echo.
    echo Launching executable...
    start "" "dist\FileToDBConverter.exe"
)

echo.
echo ========================================
echo Build process complete!
echo ========================================
pause
