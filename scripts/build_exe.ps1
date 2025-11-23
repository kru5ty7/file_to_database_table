<#
.SYNOPSIS
    Build script for File to Database Converter

.DESCRIPTION
    Creates a standalone Windows executable using PyInstaller.
    No Python installation required on target machines.

.EXAMPLE
    .\build_exe.ps1

.NOTES
    Author: File to Database Converter Team
    Requires: Python 3.10+, PyInstaller
#>

# Set error action preference
$ErrorActionPreference = "Stop"

# Colors for output
function Write-Header {
    param([string]$Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Error-Message {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Write-Warning-Message {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

# Change to project root directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location (Split-Path -Parent $scriptDir)

Write-Header "Building File to Database Converter"

# Check Python installation
Write-Host "Checking Python installation..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Success "✓ Python found: $pythonVersion"
} catch {
    Write-Error-Message "✗ ERROR: Python is not installed or not in PATH"
    Write-Host "`nPlease install Python 3.10+ from python.org"
    Read-Host "Press Enter to exit"
    exit 1
}

# Check PyInstaller
Write-Host "`nChecking PyInstaller..." -ForegroundColor Yellow
try {
    python -c "import PyInstaller" 2>$null
    Write-Success "✓ PyInstaller is installed"
} catch {
    Write-Warning-Message "PyInstaller not found. Installing..."
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Error-Message "✗ Failed to install PyInstaller"
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Success "✓ PyInstaller installed successfully"
}

# Check dependencies
Write-Host "`nChecking dependencies..." -ForegroundColor Yellow
try {
    python -c "import pandas, pyodbc, openpyxl, cryptography" 2>$null
    Write-Success "✓ All dependencies are installed"
} catch {
    Write-Warning-Message "Some dependencies are missing. Installing from requirements.txt..."
    pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Error-Message "✗ Failed to install dependencies"
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Success "✓ Dependencies installed successfully"
}

# Clean previous builds
Write-Host "`nCleaning previous build artifacts..." -ForegroundColor Yellow
if (Test-Path "build") {
    Remove-Item -Path "build" -Recurse -Force
    Write-Success "✓ Removed build directory"
}
if (Test-Path "dist") {
    Remove-Item -Path "dist" -Recurse -Force
    Write-Success "✓ Removed dist directory"
}

# Build executable
Write-Header "Building executable with PyInstaller"
Write-Host "This may take 2-5 minutes..." -ForegroundColor Yellow
Write-Host ""

$buildStartTime = Get-Date

pyinstaller FileToDBConverter.spec --clean --noconfirm

if ($LASTEXITCODE -ne 0) {
    Write-Header "BUILD FAILED!"
    Write-Error-Message "Check the error messages above for details."
    Write-Host "`nCommon issues:"
    Write-Host "  - Missing dependencies"
    Write-Host "  - Syntax errors in code"
    Write-Host "  - Insufficient disk space"
    Write-Host "`nCheck build\FileToDBConverter\warn-FileToDBConverter.txt for warnings"
    Read-Host "`nPress Enter to exit"
    exit 1
}

$buildEndTime = Get-Date
$buildDuration = $buildEndTime - $buildStartTime

# Verify executable was created
if (-not (Test-Path "dist\FileToDBConverter.exe")) {
    Write-Header "BUILD FAILED!"
    Write-Error-Message "Executable was not created!"
    Write-Host "Check the error messages above."
    Read-Host "`nPress Enter to exit"
    exit 1
}

# Get file size
$fileSize = (Get-Item "dist\FileToDBConverter.exe").Length
$fileSizeMB = [math]::Round($fileSize / 1MB, 2)

# Success message
Write-Header "BUILD SUCCESSFUL!"
Write-Success "✓ Executable created: dist\FileToDBConverter.exe"
Write-Success "✓ File size: $fileSizeMB MB"
Write-Success "✓ Build time: $($buildDuration.TotalSeconds) seconds"

Write-Host "`nYou can now run the application by double-clicking:"
Write-Host "  dist\FileToDBConverter.exe" -ForegroundColor White -BackgroundColor DarkGray

Write-Host "`nTo distribute:" -ForegroundColor Cyan
Write-Host "  1. Copy dist\FileToDBConverter.exe to target machine"
Write-Host "  2. Ensure SQL Server ODBC Driver 17+ is installed"
Write-Host "  3. Run the executable - no Python needed!"

Write-Host "`nNote: The executable includes Python runtime and all dependencies." -ForegroundColor Gray

# Offer to test the executable
Write-Host ""
$test = Read-Host "Do you want to test the executable now? (Y/N)"
if ($test -eq "Y" -or $test -eq "y") {
    Write-Host "`nLaunching executable..." -ForegroundColor Yellow
    Start-Process "dist\FileToDBConverter.exe"
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Build process complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

Read-Host "`nPress Enter to exit"
