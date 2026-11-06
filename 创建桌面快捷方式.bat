@echo off
setlocal
cd /d "%~dp0"

rem ============================================================
rem  Create a Desktop shortcut for the converter.
rem  ASCII-only on purpose (cmd.exe mis-parses UTF-8 .bat files).
rem ============================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_shortcut.ps1"
if errorlevel 1 (
    echo.
    echo   [ERROR] Failed to create the shortcut.
    echo.
    pause
    exit /b 1
)

echo.
pause
exit /b 0
