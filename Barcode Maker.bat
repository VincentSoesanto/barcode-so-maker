@echo off
cd /d "%~dp0"
rem Prefer Python 3.14 if installed, else the default Python.
set "PYW=%LOCALAPPDATA%\Programs\Python\Python314\pythonw.exe"
if exist "%PYW%" (
    start "" "%PYW%" start_app.pyw
    exit /b
)
where pyw >nul 2>nul && (
    start "" pyw -3.14 start_app.pyw
    exit /b
)
start "" pythonw start_app.pyw
