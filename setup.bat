@echo off
title BPS Report Generator — Setup
color 17
echo ============================================
echo   BPS REPORT GENERATOR — Portable Setup
echo   Laporan Penilaian Biopsikososial (PKSP)
echo   Kementerian Kesihatan Malaysia
echo ============================================
echo.

:: ── Step 1: Check / Install Python ──
echo Step 1: Checking Python...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Python ditemui!
    python --version
    set _PYTHON=python
    goto :VENV
)

:: Python tak jumpa — cuba pakai uv
echo [INFO] Python tak dijumpai. Cuba guna uv portable...

if exist uv.exe (
    echo [OK] uv portable ditemui.
) else (
    echo [INFO] Downloading uv.exe (portable)...
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip' -OutFile 'uv.zip'" >nul 2>&1
    if exist uv.zip (
        powershell -Command "Expand-Archive -Force 'uv.zip' '.'; Remove-Item 'uv.zip'" >nul 2>&1
        move uv-x86_64-pc-windows-msvc\uv.exe uv.exe >nul 2>&1
        rmdir /s /q uv-x86_64-pc-windows-msvc >nul 2>&1
        echo [OK] uv.exe downloaded!
    ) else (
        echo [ERROR] Gagal download uv.exe.
        echo Sila download manual: https://astral.sh/uv/install
        echo Letak uv.exe dalam folder BPS-App ni, then jalankan setup.bat lagi.
        pause
        exit /b
    )
)

:: Guna uv untuk install Python
echo [INFO] uv sedang install Python portable...
uv python install 3.11 >nul 2>&1
echo [OK] Python installed by uv.
set _PYTHON=uv run python

:VENV
echo.
:: ── Step 2: Virtual Environment ──
echo Step 2: Creating virtual environment...
if exist .venv (
    echo [OK] Virtual environment already exists.
) else (
    if "%_PYTHON%"=="uv run python" (
        uv venv .venv >nul 2>&1
    ) else (
        python -m venv .venv
    )
    echo [OK] Virtual environment created.
)

echo.
:: ── Step 3: Install Packages ──
echo Step 3: Installing packages...
if "%_PYTHON%"=="uv run python" (
    uv pip install --system flask python-docx fpdf2 requests -q
) else (
    .venv\Scripts\python -m pip install --upgrade pip -q
    .venv\Scripts\pip install flask python-docx fpdf2 requests -q
)
echo [OK] Packages installed.

echo.
:: ── Step 4: Config ──
echo Step 4: Creating config.txt (jika belum ada)...
echo # OUTPUT_PATH = >> config.txt 2>nul
echo [OK] config.txt ready.

echo.
:: ── Step 5: Shortcut di Desktop ──
echo Step 5: Creating shortcuts di Desktop...
powershell -ExecutionPolicy Bypass -Command ^
"$WS = New-Object -ComObject WScript.Shell; ^
$SC = $WS.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\BPS Report.lnk'); ^
$SC.TargetPath = [Environment]::GetFolderPath('System') + '\wscript.exe'; ^
$SC.Arguments = '%~dp0Jalan BPS.vbs'; ^
$SC.WorkingDirectory = '%~dp0'; ^
$SC.Description = 'BPS Report Generator - Laporan Penilaian Biopsikososial'; ^
$SC.Save(); ^
$SC2 = $WS.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Stop BPS.lnk'); ^
$SC2.TargetPath = [Environment]::GetFolderPath('System') + '\wscript.exe'; ^
$SC2.Arguments = '%~dp0Stop BPS.vbs'; ^
$SC2.WorkingDirectory = '%~dp0'; ^
$SC2.Description = 'Matikan Server BPS'; ^
$SC2.Save(); ^
Write-Host '  Shortcuts created!'"

echo.
echo ============================================
echo   SETUP SELESAI!
echo ============================================
echo.
echo Cara guna:
echo   1. Double-click shortcut "BPS Report" kat Desktop
echo   2. Buka http://localhost:5000 dalam browser
echo.
echo Laporan disimpan ke folder:
echo   %~dp0output\
echo.
echo Note: Untuk guna AI generate, PC kena ada internet.
echo       Tanpa internet, guna butang "Generate Syor (Offline)".
echo.
echo 📌 LETAK LETTERHEAD HOSPITAL:
echo   Salin gambar letterhead hospital korang ke folder static/:
echo     static\letterhead-header.png
echo     static\letterhead-footer.png
echo   Kalau takde, app tetap jalan — header/footer kosong.
echo.
pause