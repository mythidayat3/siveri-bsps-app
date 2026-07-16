@echo off
title Buat Paket Update BSPS DB
echo =======================================================
echo          BUAT PAKET UPDATE BSPS DB
echo =======================================================
echo.

cd /d "%~dp0"

:: Step 1: Build frontend
echo [1/5] Build frontend...
cd bsps-db-app\frontend
call npx vite build
if %errorlevel% neq 0 (
    echo [ERROR] Frontend build gagal!
    pause
    exit /b 1
)
cd ..\..
echo       Frontend build OK!

:: Step 2: Buat folder sementara
echo.
echo [2/5] Menyiapkan folder sementara...
set TEMP_DIR=_update_package_temp
if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%"
mkdir "%TEMP_DIR%"
mkdir "%TEMP_DIR%\bsps-db-app\backend"
mkdir "%TEMP_DIR%\bsps-db-app\backend\templates"
mkdir "%TEMP_DIR%\bsps-db-app\frontend\dist"

:: Step 3: Copy file kode (TANPA database)
echo.
echo [3/5] Mengcopy file kode...

:: Backend - Python files
copy "bsps-db-app\backend\app.py" "%TEMP_DIR%\bsps-db-app\backend\" >nul
copy "bsps-db-app\backend\database.py" "%TEMP_DIR%\bsps-db-app\backend\" >nul
copy "bsps-db-app\backend\requirements.txt" "%TEMP_DIR%\bsps-db-app\backend\" >nul
echo       Backend Python files...

:: Backend - Templates
copy "bsps-db-app\backend\templates\*.xlsx" "%TEMP_DIR%\bsps-db-app\backend\templates\" >nul
copy "bsps-db-app\backend\templates\*.docx" "%TEMP_DIR%\bsps-db-app\backend\templates\" >nul
echo       Backend templates...

:: Frontend - dist
xcopy /E /Q /Y "bsps-db-app\frontend\dist\*" "%TEMP_DIR%\bsps-db-app\frontend\dist\" >nul
echo       Frontend dist...

:: Root files
copy "jalankan_server.bat" "%TEMP_DIR%\" >nul
copy "install.bat" "%TEMP_DIR%\" >nul
copy "update.bat" "%TEMP_DIR%\" >nul
echo       Root files...

:: Step 4: Buat zip
echo.
echo [4/5] Membuat file zip...
set TIMESTAMP=%DATE:~10,4%%DATE:~7,2%%DATE:~4,2%
set ZIP_NAME=UPDATE_BSPS_DB_%TIMESTAMP%.zip

:: Hapus zip lama jika ada
if exist "%ZIP_NAME%" del "%ZIP_NAME%"

:: Zip menggunakan PowerShell
powershell -Command "Compress-Archive -Path '%TEMP_DIR%\*' -DestinationPath '%ZIP_NAME%' -Force"
if %errorlevel% neq 0 (
    echo [ERROR] Gagal membuat zip!
    pause
    exit /b 1
)

:: Step 5: Bersihkan folder sementara
echo.
echo [5/5] Membersihkan folder sementara...
rmdir /s /q "%TEMP_DIR%"

echo.
echo =======================================================
echo PAKET UPDATE BERHASIL DIBUAT!
echo.
echo File: %ZIP_NAME%
echo.
echo Kirim file ini ke user untuk di-update.
echo User hanya perlu:
echo   1. Extract zip ke folder project
echo   2. Jalankan update.bat
echo =======================================================
pause
