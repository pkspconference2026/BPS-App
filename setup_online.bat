@echo off
title BPS App — Online Setup
cd /d "%~dp0"

echo ============================================
echo   BPS REPORT GENERATOR — Setup Online
echo   Muat turun dari GitHub + Setup Auto
echo ============================================
echo.

:: 1) Check internet
echo [1/5] Memeriksa internet...
ping -n 1 raw.githubusercontent.com >nul 2>&1
if errorlevel 1 (
    echo ❌ TIADA INTERNET! Pastikan PC kawan ada internet.
    echo    Guna setup.bat (offline/USB) sebagai ganti.
    pause
    exit /b 1
)
echo    ✅ Internet OK

:: 2) Download zip dari GitHub
echo [2/5] Muat turun BPS-App dari GitHub...
set "ZIP=%TEMP%\BPS-App.zip"
set "DEST=%USERPROFILE%\Desktop\BPS-App"

:: Guna PowerShell untuk download
powershell -Command "& {try{Invoke-WebRequest -Uri 'https://github.com/pkspconference2026/BPS-App/archive/refs/heads/main.zip' -OutFile '%ZIP%' -ErrorAction Stop; Write-Host '   ✅ Muat turun siap'}catch{Write-Host '❌ Gagal download: ' $_; exit 1}}"
if errorlevel 1 pause & exit /b 1

:: 3) Extract
echo [3/5] Extract...
if exist "%DEST%" rmdir /s /q "%DEST%"
powershell -Command "& {Expand-Archive -Path '%ZIP%' -DestinationPath '%TEMP%\BPS-Extract' -Force; Move-Item '%TEMP%\BPS-Extract\BPS-App-main\*' '%DEST%' -Force; rmdir /s /q '%TEMP%\BPS-Extract'; Remove-Item '%ZIP%' -Force}"
echo    ✅ Extract ke Desktop

:: 4) Setup venv + install deps
echo [4/5] Pasang Python environment...
cd /d "%DEST%"

:: Check python
python --version >nul 2>&1
if errorlevel 1 (
    echo    ❌ Python tak jumpa! Cuba guna uv.exe...
    if exist "uv.exe" (
        echo    ✅ Guna uv.exe portable
        uv.exe venv .venv
        uv.exe pip install --python .venv\Scripts\python.exe flask python-docx fpdf2 requests
    ) else (
        echo    ❌ Takde Python dan takde uv.exe. Pasang Python dulu.
        echo       https://www.python.org/downloads/
        pause
        exit /b 1
    )
) else (
    python -m venv .venv
    .venv\Scripts\python -m pip install --quiet flask python-docx fpdf2 requests
)
echo    ✅ Environment siap

:: 5) Buat shortcut
echo [5/5] Buat shortcut Desktop...
powershell -Command "& {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\BPS Report.lnk')
    $Shortcut.TargetPath = '%SystemRoot%\System32\wscript.exe'
    $Shortcut.Arguments = '%DEST%\Jalan BPS.vbs'
    $Shortcut.WorkingDirectory = '%DEST%'
    $Shortcut.Description = 'BPS Report Generator'
    if (Test-Path '%DEST%\BPS.ico') { $Shortcut.IconLocation = '%DEST%\BPS.ico' }
    $Shortcut.Save()
    # Stop shortcut
    $Shortcut2 = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\Stop BPS.lnk')
    $Shortcut2.TargetPath = '%SystemRoot%\System32\wscript.exe'
    $Shortcut2.Arguments = '%DEST%\Stop BPS.vbs'
    $Shortcut2.WorkingDirectory = '%DEST%'
    $Shortcut2.Description = 'Matikan Server BPS'
    $Shortcut2.Save()
}"
echo    ✅ Shortcut 'BPS Report' & 'Stop BPS' dah kat Desktop

:: Selesai
echo ============================================
echo   ✅ SETUP SIAP!
echo ============================================
echo.
echo  Buka app: double-click 'BPS Report' kat Desktop
echo  atau buka http://localhost:5000
echo.
echo  Untuk update masa depan, tekan butang
echo  🔄 Semak Update dalam app tu sendiri.
echo.
echo 📌 LETAK LETTERHEAD HOSPITAL:
echo   Salin gambar letterhead hospital korang ke folder static/:
echo     static\letterhead-header.png
echo     static\letterhead-footer.png
echo   Kalau takde, app tetap jalan — header/footer kosong.
echo.
pause
start "" "%DEST%\.venv\Scripts\python.exe" "%DEST%\app.py"