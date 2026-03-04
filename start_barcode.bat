@echo off
setlocal
timeout /t 5 /nobreak >nul

set "APPDIR=C:\Users\Admira\Documents\Barcode"
set "PY=%APPDIR%\venv\Scripts\python.exe"
set "SCRIPT=%APPDIR%\barcode.py"

echo APPDIR=%APPDIR%
echo PY=%PY%
echo SCRIPT=%SCRIPT%
echo.

if not exist "%PY%" (
  echo [ERROR] No existe el python del venv: %PY%
  pause
  exit /b 1
)

echo Ejecutando python...
"%PY%" "%SCRIPT%"

echo.
echo El script terminó. Revisa el mensaje anterior.
pause
