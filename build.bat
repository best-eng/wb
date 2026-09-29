@echo off
chcp 65001 >nul
rem Сборка wb_stocks.exe. Запускать из папки с wb_stocks.py.

rem Ищем Python сами: голая команда pyinstaller часто не находится,
rem потому что папка Scripts не попадает в PATH. Через -m это не важно.
set PY=
python --version >nul 2>&1
if %errorlevel%==0 set PY=python
if defined PY goto found

py --version >nul 2>&1
if %errorlevel%==0 set PY=py
if defined PY goto found

echo.
echo Python не найден.
echo.
echo Установите его с https://www.python.org/downloads/ и при установке
echo обязательно отметьте "Add Python to PATH".
echo Затем закройте это окно и запустите build.bat заново: уже открытая
echo консоль о новом PATH не знает.
echo.
pause
exit /b 1

:found
echo Использую: %PY%
echo.

%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail

rem --collect-all curl_cffi обязателен: без него в exe не попадают
rem библиотеки libcurl, подмена TLS-отпечатка молча отключается
rem и Wildberries снова начинает отвечать 403.
%PY% -m PyInstaller ^
  --onefile ^
  --console ^
  --name wb_stocks ^
  --collect-all curl_cffi ^
  --collect-all openpyxl ^
  --clean ^
  --noconfirm ^
  wb_stocks.py
if errorlevel 1 goto fail

echo.
echo Готово: dist\wb_stocks.exe
goto end

:fail
echo.
echo Сборка не удалась. Смотрите сообщения выше.

:end
pause
