@echo off
REM Build script for Liturgie Samensteller
REM Creates a standalone Windows executable

echo ========================================
echo Building Liturgie Samensteller
echo ========================================
echo.
echo Options:
echo   1. Single executable (one .exe file)
echo   2. Folder distribution (faster startup)
echo.
set /p choice="Select option (1 or 2): "

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Install/upgrade build dependencies
echo.
echo Installing build dependencies...
pip install --upgrade pyinstaller pywin32

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build

REM Run PyInstaller
echo.
echo Running PyInstaller...

if "%choice%"=="1" (
    python -m PyInstaller build_onefile.spec --clean --noconfirm
) else (
    python -m PyInstaller build.spec --clean --noconfirm
)

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build successful!
echo ========================================
echo.
if "%choice%"=="1" (
    echo Output: dist\LiturgieSamensteller.exe
) else (
    echo Output folder: dist\LiturgieSamensteller\
    echo Run: dist\LiturgieSamensteller\LiturgieSamensteller.exe
)
echo.
pause
