@echo off
title BPS Report Generator
cd /d "%~dp0"
echo ============================================
echo   BPS REPORT GENERATOR
echo   Laporan Penilaian Biopsikososial
echo ============================================
echo.
if not exist .venv (
    echo [ERROR] Virtual environment tak dijumpai!
    echo Sila jalankan setup.bat dulu.
    pause
    exit /b
)
echo Starting server...
echo Buka http://localhost:5000 dalam browser
echo Ctrl+C untuk berhenti
echo.
.venv\Scripts\python app.py
pause