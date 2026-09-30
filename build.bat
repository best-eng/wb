@echo off
setlocal

rem Builds wb_stocks.exe. Run from the folder containing wb_stocks.py.
rem Uses "python -m PyInstaller": a bare "pyinstaller" command is often
rem missing from PATH even when the package is installed.

cd /d "%~dp0"

set PY=
python --version >nul 2>&1
if %errorlevel%==0 set PY=python
if defined PY goto found
py --version >nul 2>&1
if %errorlevel%==0 set PY=py
if defined PY goto found

echo.
echo Python not found.
echo Install it from https://www.python.org/downloads/ and tick
echo "Add Python to PATH". Then close this window, open it again
echo and run build.bat once more - an open console does not see
echo the new PATH.
echo.
goto end

:found
echo Using: %PY%
echo.

%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail

rem --collect-all curl_cffi is required: without it libcurl is left out
rem of the exe, TLS impersonation silently stops working and Wildberries
rem starts answering 403 again.
%PY% -m PyInstaller --onefile --console --name wb_stocks ^
  --collect-all curl_cffi --collect-all openpyxl --hidden-import wb_browser --hidden-import websocket ^
  --clean --noconfirm wb_stocks.py
if errorlevel 1 goto fail

echo.
echo Done: dist\wb_stocks.exe
goto end

:fail
echo.
echo Build failed. See the messages above.

:end
echo.
pause
