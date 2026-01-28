@echo off
REM Clean build artifacts and temporary files

echo Cleaning build artifacts...

REM Remove PyInstaller build folders
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "installer_output" rmdir /s /q installer_output

REM Remove Python cache
for /d /r . %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
for /r . %%f in (*.pyc) do @if exist "%%f" del /q "%%f"
for /r . %%f in (*.pyo) do @if exist "%%f" del /q "%%f"

REM Remove other temp files
if exist "nul" del /q nul

echo.
echo Clean complete!
pause
