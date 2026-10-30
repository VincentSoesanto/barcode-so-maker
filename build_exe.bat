@echo off
rem Rebuilds "Pembuat Barcode.exe" (in this folder) after barcode_app.py has been changed.
cd /d "%~dp0"
set "PY=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -m PyInstaller --noconfirm --onefile --windowed --name "Pembuat Barcode" ^
    --collect-data docx --hidden-import xlrd --hidden-import openpyxl ^
    --add-data "%~dp0SO\Form SO (Sales Order).xlsx;SO" ^
    --distpath "%~dp0_build\dist" --workpath "%~dp0_build\work" --specpath "%~dp0_build" ^
    start_app.pyw
if errorlevel 1 (
    echo.
    echo BUILD GAGAL - lihat pesan di atas.
    pause
    exit /b 1
)

copy /y "_build\dist\Pembuat Barcode.exe" "%~dp0" >nul
if errorlevel 1 (
    echo.
    echo Gagal menyalin - apakah Pembuat Barcode.exe sedang terbuka? Tutup dulu lalu coba lagi.
    pause
    exit /b 1
)
rmdir /s /q "_build"
echo.
echo Selesai: "%~dp0Pembuat Barcode.exe"
pause
