@echo off
title Update Aplikasi BSPS DB
echo =======================================================
echo          UPDATE APLIKASI BSPS DB
echo =======================================================
echo.

:: Deteksi lokasi project (cari folder bsps-db-app)
cd /d "%~dp0"
if not exist "bsps-db-app" (
    echo [ERROR] Folder bsps-db-app tidak ditemukan!
    echo Pastikan file update.bat berada di root folder project.
    pause
    exit /b 1
)

:: Step 1: Backup database
echo [1/3] Backup database...
set DB_PATH=bsps-db-app\backend\bsps_db.sqlite
set DB_PATH2=bsps-db-app\backend\bsps_data.db
set TIMESTAMP=%DATE:~10,4%%DATE:~7,2%%DATE:~4,2%

if exist "%DB_PATH%" (
    copy "%DB_PATH%" "%DB_PATH%.backup_%TIMESTAMP%" >nul 2>&1
    echo       Database bsps_db.sqlite berhasil di-backup
) else if exist "%DB_PATH2%" (
    copy "%DB_PATH2%" "%DB_PATH2%.backup_%TIMESTAMP%" >nul 2>&1
    echo       Database bsps_data.db berhasil di-backup
) else (
    echo       Tidak ada database ditemukan (baru pertama kali)
)

:: Step 2: Install/Update Python dependencies
echo.
echo [2/3] Install dependencies Python...
cd bsps-db-app\backend
pip install -r requirements.txt --quiet 2>nul
if %errorlevel% neq 0 (
    echo       [WARNING] pip install gagal, mencoba tanpa --quiet...
    pip install -r requirements.txt
)
echo       Dependencies Python sudah ter-update!
cd ..\..

:: Step 3: Selesai
echo.
echo [3/3] Verifikasi file...
set MISSING=0
if not exist "bsps-db-app\backend\app.py" (
    echo       [ERROR] app.py tidak ditemukan!
    set MISSING=1
)
if not exist "bsps-db-app\backend\database.py" (
    echo       [ERROR] database.py tidak ditemukan!
    set MISSING=1
)
if not exist "bsps-db-app\frontend\dist\index.html" (
    echo       [ERROR] frontend/dist tidak ditemukan!
    set MISSING=1
)

if %MISSING%==1 (
    echo.
    echo [ERROR] Beberapa file penting tidak ditemukan!
    echo Silakan extract ulang file update ke folder project.
    pause
    exit /b 1
)

echo.
echo =======================================================
echo Update berhasil! Data Anda tetap aman.
echo.
echo Untuk menjalankan aplikasi:
echo   klik ganda jalankan_server.bat
echo =======================================================
pause
