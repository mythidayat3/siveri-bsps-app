@echo off
title Memulai Server BSPS DB
echo =======================================================
echo          MEMULAI LAYANAN APLIKASI BSPS DB
echo =======================================================
echo.

:: Pindah ke direktori tempat file .bat berada
cd /d "%~dp0"

echo [1/3] Menjalankan Server Backend (FastAPI)...
start "Backend - BSPS DB" cmd /k "cd /d bsps-db-app\backend && uvicorn app:app --host 127.0.0.1 --port 8000"

echo [2/3] Menjalankan Server Frontend (Vite React)...
start "Frontend - BSPS DB" cmd /k "cd /d bsps-db-app\frontend && npx vite --host 127.0.0.1 --port 3000"

echo [3/3] Menunggu server siap...
timeout /t 5 >nul

echo Membuka aplikasi di browser...
start http://127.0.0.1:3000/

echo.
echo =======================================================
echo Layanan telah berhasil dijalankan!
echo Silakan biarkan kedua jendela CMD server tetap terbuka.
echo =======================================================
echo.
pause
