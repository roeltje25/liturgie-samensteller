@echo off
REM Build script for Liturgie Samensteller
REM Creates a standalone Windows executable

echo ========================================
echo Building Liturgie Samensteller
echo ========================================
echo.
echo Options:
echo   1. Single executable (PyInstaller - one .exe file)
echo   2. Folder distribution (PyInstaller - faster startup)
echo   3. Compiled native (Nuitka - compiles to C, best performance)
echo.
set /p choice="Select option (1, 2, or 3): "

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build

if "%choice%"=="3" (
    goto nuitka_build
) else (
    goto pyinstaller_build
)

:pyinstaller_build
REM Install/upgrade PyInstaller dependencies
echo.
echo Installing PyInstaller dependencies...
pip install --upgrade pyinstaller pywin32

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
exit /b 0

:nuitka_build
REM Install/upgrade Nuitka and project dependencies
echo.
echo Installing Nuitka and project dependencies...
echo Note: First build may take longer as Nuitka downloads C compiler if needed.
pip install --upgrade nuitka pywin32 ordered-set zstandard
pip install -r requirements.txt

REM Get version from Python
for /f "tokens=*" %%i in ('python -c "from src import __version__; print(__version__)"') do set APP_VERSION=%%i
echo Building version: %APP_VERSION%

REM Run Nuitka
echo.
echo Running Nuitka (compiling Python to C, this may take several minutes)...

python -m nuitka ^
    --standalone ^
    --onefile ^
    --windows-console-mode=disable ^
    --enable-plugin=pyqt6 ^
    --include-data-dir=src/i18n=src/i18n ^
    --include-package=pptx ^
    --include-package-data=pptx ^
    --include-package=lxml ^
    --include-package=openpyxl ^
    --include-package=win32com ^
    --include-module=pythoncom ^
    --include-module=pywintypes ^
    --include-package=yt_dlp ^
    --include-package=requests ^
    --include-package=PIL ^
    --include-package=certifi ^
    --include-package=charset_normalizer ^
    --include-package=urllib3 ^
    --include-package=idna ^
    --nofollow-import-to=tkinter ^
    --nofollow-import-to=matplotlib ^
    --nofollow-import-to=numpy ^
    --nofollow-import-to=pandas ^
    --nofollow-import-to=scipy ^
    --nofollow-import-to=cv2 ^
    --output-dir=dist ^
    --output-filename=LiturgieSamensteller.exe ^
    --company-name="PowerPoint Mixer" ^
    --product-name="Liturgie Samensteller" ^
    --file-version=%APP_VERSION% ^
    --product-version=%APP_VERSION% ^
    --assume-yes-for-downloads ^
    --msvc=latest ^
    run.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    echo.
    echo If Nuitka failed, make sure you have a C compiler installed.
    echo You can install Visual Studio Build Tools or MinGW64.
    echo Nuitka will try to download MinGW64 automatically on first run.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build successful! (Nuitka - Native Compiled)
echo ========================================
echo.
echo Output: dist\LiturgieSamensteller.exe
echo.
echo Note: Nuitka-compiled executables are:
echo   - Faster at runtime (native machine code)
echo   - Better protected (no Python bytecode)
echo   - Larger file size (includes optimizations)
echo.
pause
exit /b 0
