@echo off
setlocal

rem Quick check of proxy channels to Wildberries (~10 seconds).
rem Put your proxies into proxy.txt next to this file, then run it.
rem All messages are printed by the program itself, in Russian.

cd /d "%~dp0"

if exist "wb_stocks.exe" (
  wb_stocks.exe --test
  goto done
)

set PY=
python --version >nul 2>&1
if %errorlevel%==0 set PY=python
if defined PY goto runpy
py --version >nul 2>&1
if %errorlevel%==0 set PY=py
if defined PY goto runpy

echo.
echo wb_stocks.exe not found, and Python is not installed.
echo Put this file next to wb_stocks.exe, or install Python
echo from https://www.python.org/downloads/ ("Add Python to PATH").
echo.
goto done

:runpy
%PY% wb_stocks.py --test

:done
echo.
pause
