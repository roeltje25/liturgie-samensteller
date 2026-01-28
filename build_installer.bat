@echo off
REM Build script for Liturgie Samensteller Installer
REM Creates a Windows installer using Inno Setup

echo ========================================
echo Building Liturgie Samensteller Installer
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check if Inno Setup is installed
where iscc >nul 2>&1
if errorlevel 1 (
    echo ERROR: Inno Setup not found. Please install from https://jrsoftware.org/isinfo.php
    echo        and add it to your PATH.
    pause
    exit /b 1
)

REM Install/upgrade build dependencies
echo Installing build dependencies...
pip install --upgrade pyinstaller pywin32

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "installer_output" rmdir /s /q installer_output

REM Run PyInstaller (folder mode for installer)
echo.
echo Running PyInstaller...
python -m PyInstaller build.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

REM Run Inno Setup
echo.
echo Creating installer with Inno Setup...
iscc installer.iss

if errorlevel 1 (
    echo.
    echo ERROR: Inno Setup failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Installer created successfully!
echo ========================================
echo.
echo Output: installer_output\LiturgieSamensteller_Setup_1.0.0.exe
echo.
pause
