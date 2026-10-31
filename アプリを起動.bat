@echo off
setlocal
cd /d "%~dp0"
title Pro Investor Analyzer (local server)
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
  where python >nul 2>nul
  if not errorlevel 1 set "PY=python"
)
if not defined PY goto nopython
%PY% server.py %*
if errorlevel 1 pause
exit /b

:nopython
echo.
echo   Python 3 was not found on this PC.
echo   Install it from https://www.python.org/downloads/
echo   (turn on "Add python.exe to PATH" in the installer), then run this file again.
echo.
pause
exit /b 1
