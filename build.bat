@echo off
chcp 65001 >nul
rem Сборка wb_stocks.exe. Запускать из папки с wb_stocks.py.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail

rem --collect-all curl_cffi обязателен: без него в exe не попадают
rem библиотеки libcurl и сертификаты, и подмена TLS-отпечатка не работает.
python -m PyInstaller ^
  --onefile ^
  --console ^
  --name wb_stocks ^
  --collect-all curl_cffi ^
  --collect-all openpyxl ^
  --clean ^
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
