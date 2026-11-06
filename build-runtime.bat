@echo off
setlocal
title Build portable runtime
cd /d "%~dp0"

rem ============================================================
rem  Build the portable runtime (runtime\python + runtime\everything)
rem  ASCII-only on purpose: cmd.exe mis-parses UTF-8 text in .bat files.
rem ============================================================

rem Prefer PowerShell 7 (defaults to UTF-8); fall back to Windows PowerShell.
set "PS="
for %%I in (pwsh.exe) do if not defined PS set "PS=%%~$PATH:I"
if not defined PS set "PS=powershell.exe"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-runtime.ps1"
if errorlevel 1 (
    echo.
    echo   [ERROR] Portable runtime build failed. See the message above.
    echo.
    pause
    exit /b 1
)

echo.
pause
exit /b 0
