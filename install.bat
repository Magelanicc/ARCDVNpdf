@echo off
setlocal
title ARCDVNpdf Installer

echo.
echo ==========================================
echo          ARCDVNpdf Installer
echo ==========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py"
) else (
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] Python tidak ditemukan.
        echo Install Python 3.10+ terlebih dahulu.
        pause
        exit /b 1
    )
    set "PY=python"
)

echo [1/3] Installing ARCDVNpdf...
%PY% -m pip install --upgrade pip
%PY% -m pip install .
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Instalasi gagal.
    pause
    exit /b 1
)

echo.
echo [2/3] Locating Python Scripts directory...
for /f "delims=" %%P in ('%PY% -c "import sysconfig; print(sysconfig.get_path(\"scripts\"))"') do set "SCRIPTS=%%P"

if not defined SCRIPTS (
    echo [ERROR] Scripts directory tidak ditemukan.
    pause
    exit /b 1
)

echo Scripts: %SCRIPTS%

echo.
echo [3/3] Adding Scripts directory to USER PATH...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p=[Environment]::GetEnvironmentVariable('Path','User'); if(-not $p){$p=''}; if(($p -split ';') -notcontains '%SCRIPTS%'){[Environment]::SetEnvironmentVariable('Path', (($p.TrimEnd(';') + ';%SCRIPTS%').Trim(';')), 'User')}; Write-Host 'User PATH updated.'"

echo.
echo ==========================================
echo Installation selesai.
echo ==========================================
echo.
echo TUTUP CMD/PowerShell yang lama, lalu buka terminal baru.
echo Setelah itu coba:
echo.
echo     ARCDVNpdf --version
echo     ARCDVNpdf doctor
echo     ARCDVNpdf
echo.
pause
