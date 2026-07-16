@echo off
title Instalasi BSPS DB
echo =======================================================
echo          INSTALASI APLIKASI BSPS DB
echo =======================================================
echo.

cd /d "%~dp0"

echo [1/3] Memeriksa Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python belum terinstal atau belum ada di PATH!
    echo Silakan install Python dari https://www.python.org/downloads/
    echo CENTANG "Add Python to PATH" saat instalasi.
    pause
    exit /b 1
)
echo       Python ditemukan!

echo [2/3] Memeriksa Node.js...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js belum terinstal atau belum ada di PATH!
    echo Silakan install Node.js dari https://nodejs.org/
    pause
    exit /b 1
)
echo       Node.js ditemukan!

echo [3/3] Menginstall dependencies...
echo.
echo --- Backend (Python) ---
cd bsps-db-app\backend
pip install --only-binary :all: -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Gagal install dependencies Python!
    pause
    exit /b 1
)
echo       Backend OK!

echo.
echo --- Frontend (Node.js) ---
cd ..\frontend
call npm install
if %errorlevel% neq 0 (
    echo [ERROR] Gagal install dependencies Frontend!
    pause
    exit /b 1
)
echo       Frontend OK!

echo.
echo =======================================================
echo Instalasi selesai! Untuk menjalankan aplikasi:
echo   klik ganda jalankan_server.bat
echo =======================================================
pause
