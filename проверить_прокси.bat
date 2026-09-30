@echo off
chcp 65001 >nul
rem Быстрая проверка каналов до Wildberries: ~10 секунд вместо полного прогона.
rem Положите адреса в proxy.txt рядом и запустите этот файл двойным щелчком.

if exist "%~dp0wb_stocks.exe" (
  "%~dp0wb_stocks.exe" --test
  goto end
)

set PY=
python --version >nul 2>&1
if %errorlevel%==0 set PY=python
if defined PY goto run
py --version >nul 2>&1
if %errorlevel%==0 set PY=py
if defined PY goto run

echo Не найден ни wb_stocks.exe, ни Python.
pause
exit /b 1

:run
%PY% "%~dp0wb_stocks.py" --test

:end
pause
