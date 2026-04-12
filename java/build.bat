@echo off
REM ============================================================
REM  Liturgie Samensteller – Native Windows build
REM  Produces: target\gluonfx\x86_64-windows\liturgie-samensteller.exe
REM
REM  Prerequisites:
REM    1. GraalVM JDK 21+  (winget install GraalVM.GraalVM.21)
REM       Set JAVA_HOME to the GraalVM JDK directory.
REM    2. Visual Studio 2022 Build Tools with "Desktop development
REM       with C++" workload.
REM    3. Run this script from a normal Command Prompt (GluonFX
REM       sets up the MSVC environment automatically via vcvars64).
REM
REM  For a JVM dev run (no native build needed):
REM    mvnw javafx:run
REM ============================================================
cd /d "%~dp0"

if not defined JAVA_HOME (
    echo ERROR: JAVA_HOME is not set. Please install GraalVM JDK 21+.
    exit /b 1
)

echo Building native executable...
echo JAVA_HOME = %JAVA_HOME%
echo.

call mvnw.cmd -Pnative gluonfx:build

if %ERRORLEVEL% neq 0 (
    echo.
    echo Build FAILED. See output above.
    exit /b %ERRORLEVEL%
)

echo.
echo ============================================================
echo Build SUCCESS
echo Output: target\gluonfx\x86_64-windows\liturgie-samensteller.exe
echo ============================================================
