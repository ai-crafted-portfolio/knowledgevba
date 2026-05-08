@echo off
setlocal

REM ============================================================
REM  Install-KnowledgevbaModules.bat
REM
REM  Drag & drop your target .xlsm onto this bat, OR run from cmd:
REM    Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm" [/demo]
REM
REM  Prerequisite (one-time, on the target Excel):
REM    File -> Options -> Trust Center -> Trust Center Settings ->
REM      Macro Settings -> "Trust access to the VBA project object model" ON
REM ============================================================

if "%~1"=="" (
    echo [ERROR] Drop your target .xlsm onto this bat, or pass it as arg 1.
    echo.
    echo Usage:
    echo   Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm"
    echo   Install-KnowledgevbaModules.bat "C:\path\to\book.xlsm" /demo
    pause
    exit /b 1
)

set "TARGET_XLSM=%~1"
set "DEMO_FLAG="
if /I "%~2"=="/demo" set "DEMO_FLAG=-Demo"

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Install-KnowledgevbaModules.ps1"

if not exist "%PS1%" (
    echo [ERROR] PowerShell script not found: %PS1%
    pause
    exit /b 1
)

echo ============================================================
echo  knowledgevba VBA module installer
echo ============================================================
echo  Target xlsm: %TARGET_XLSM%
echo  Demo mode  : %DEMO_FLAG%
echo  PowerShell : %PS1%
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -XlsmPath "%TARGET_XLSM%" %DEMO_FLAG%

echo.
echo ============================================================
echo  Done. Check console output above for any warnings.
echo ============================================================
pause
