@echo off
title BPS App — Build Portable EXE
cd /d "%~dp0"

echo ============================================
echo   BPS REPORT GENERATOR — Build Portable EXE
echo ============================================
echo.

if not exist .venv\Scripts\python.exe (
    echo [ERROR] .venv tak jumpa. Jalankan setup dulu.
    pause
    exit /b 1
)

echo [1/3] Compiling BPS.exe (single file)...
.venv\Scripts\python.exe -m PyInstaller ^
    --noconsole ^
    --onefile ^
    --name BPS ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import weasyprint ^
    --icon "BPS.ico" ^
    app.py

if errorlevel 1 (
    echo [ERROR] Build gagal!
    pause
    exit /b 1
)

echo [2/3] Menyediakan folder BPS-Portable...
if exist dist\BPS-Portable rmdir /s /q dist\BPS-Portable
mkdir dist\BPS-Portable
move dist\BPS.exe dist\BPS-Portable\ >nul
echo     - BPS.exe dipindah.

echo [3/3] Salin static + buat folder output...
xcopy static dist\BPS-Portable\static\ /E /I /Y /Q >nul
mkdir dist\BPS-Portable\output >nul 2>&1
echo OUTPUT_PATH = >> dist\BPS-Portable\config.txt 2>nul

echo.
echo ============================================
echo   BUILD SIAP!
echo ============================================
echo.
echo Folder siap: dist\BPS-Portable\
echo   - BPS.exe          (double-click untuk jalan)
echo   - static\          (letterhead hospital)
echo   - output\          (laporan disimpan sini)
echo   - config.txt       (ubah OUTPUT_PATH kalau nak)
echo.
echo CARA HANTAR KE KAWAN:
echo   1. Zip folder BPS-Portable
echo   2. Hantar via USB / email / WhatsApp
echo   3. Kawan unzip & double-click BPS.exe
echo   4. Browser terbuka sendiri ke localhost:5000
echo.
echo NOTE: Auto-update TAK aktif dalam EXE.
echo   Kalau app dikemaskini, rebuild & hantar balik.
echo.
pause
