@echo off
rem Makes "Pembuat Barcode.exe" and "Pembuat Barcode.zip" in this folder (see release.py).
cd /d "%~dp0"
set "PY=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not exist "%PY%" (
    where py >nul 2>nul && (set "PY=py") || (set "PY=python")
)
"%PY%" release.py %*
echo.
pause
