@echo off
setlocal
title DOCM to DOCX Converter
cd /d "%~dp0"

rem ============================================================
rem  One-click launcher.
rem  Keep this file ASCII-only: cmd.exe mis-parses UTF-8 text
rem  in .bat files. All Chinese messages come from Python.
rem ============================================================

set "PY="
for %%I in (python.exe) do if not defined PY set "PY=%%~$PATH:I"

if not defined PY goto :nopython

rem --- dependency precheck (visible console so errors can be read) ---
"%PY%" "%~dp0docm2docx.py" --check
if errorlevel 1 goto :deps

rem --- launch the GUI with pythonw (no black console window) ---
set "PYW="
for %%I in (pythonw.exe) do if not defined PYW set "PYW=%%~$PATH:I"
if not defined PYW set "PYW=%PY%"

start "DOCM to DOCX Converter" "%PYW%" "%~dp0docm2docx.py" %*
exit /b 0

:nopython
echo.
echo   [ERROR] Python was not found on PATH.
echo.
echo   Please install Python 3.8 or newer:
echo       https://www.python.org/downloads/
echo   Remember to tick "Add Python to PATH" during setup.
echo.
pause
exit /b 1

:deps
echo.
echo   [ERROR] Dependencies are missing or the environment is not ready.
echo   See the message above, then install them with:
echo       "%PY%" -m pip install -r "%~dp0requirements.txt"
echo.
pause
exit /b 1
