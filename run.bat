@echo off
setlocal
title DOCM to DOCX Converter
cd /d "%~dp0"

rem ============================================================
rem  One-click launcher.
rem  Priority: bundled portable runtime, then system Python.
rem  Keep this file ASCII-only: cmd.exe mis-parses UTF-8 text
rem  in .bat files. All Chinese messages come from Python.
rem ============================================================

set "PYHOME="
if exist "%~dp0runtime\python\python.exe" set "PYHOME=%~dp0runtime\python"
if not defined PYHOME for %%I in (python.exe) do if not defined PYHOME set "PYHOME=%%~dpI"

if not defined PYHOME goto :nopython
set "PYEXE=%PYHOME%\python.exe"
if not exist "%PYEXE%" goto :nopython

rem --- dependency precheck (visible console so errors can be read) ---
"%PYEXE%" "%~dp0docm2docx.py" --check
if errorlevel 1 goto :deps

rem --- launch the GUI (pythonw = no black console window) ---
set "PYW=%PYHOME%\pythonw.exe"
if not exist "%PYW%" set "PYW=%PYEXE%"

start "DOCM to DOCX Converter" "%PYW%" "%~dp0docm2docx.py" %*
exit /b 0

:nopython
echo.
echo   [ERROR] Python was not found.
echo.
echo   Option A - build the bundled portable runtime (no install needed):
echo       run  build-runtime.bat  once on a machine that has Python,
echo       then copy the whole folder anywhere.
echo.
echo   Option B - install Python 3.8 or newer:
echo       https://www.python.org/downloads/
echo   Remember to tick "Add Python to PATH" during setup.
echo.
pause
exit /b 1

:deps
echo.
echo   [ERROR] Dependencies are missing or the environment is not ready.
echo   See the message above, then install them with:
echo       "%PYEXE%" -m pip install -r "%~dp0requirements.txt"
echo.
pause
exit /b 1
