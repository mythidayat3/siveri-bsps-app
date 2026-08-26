@echo off
title Memulai Server BSPS DB
echo =======================================================
echo          MEMULAI LAYANAN APLIKASI BSPS DB
echo =======================================================
echo.

:: Pindah ke direktori tempat file .bat berada
cd /d "%~dp0"

echo [1/3] Menjalankan Server Backend (FastAPI)...
start "Backend - BSPS DB" cmd /k "cd /d bsps-db-app\backend && uvicorn app:app --host 0.0.0.0 --port 8000 --reload"

echo [2/3] Menjalankan Server Frontend (Vite React)...
start "Frontend - BSPS DB" cmd /k "cd /d bsps-db-app\frontend && npx vite --host 0.0.0.0 --port 3000"

echo [3/3] Menunggu server siap...
timeout /t 5 >nul

echo Membuka aplikasi di browser laptop ini...
start http://localhost:3000/

echo.
echo =======================================================
echo  Layanan telah berhasil dijalankan!
echo  - Akses dari Laptop Ini: http://localhost:3000/
echo  - Akses Rekan Kerja (Wi-Fi): http://192.168.18.187:3000/
echo =======================================================
echo.
pause
