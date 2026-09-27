@echo off
setlocal
title ARCDVNpdf Uninstaller

echo Menghapus ARCDVNpdf dari environment Python...
where py >nul 2>nul
if %errorlevel%==0 (
    py -m pip uninstall arcdvnpdf -y
) else (
    python -m pip uninstall arcdvnpdf -y
)

echo.
echo ARCDVNpdf sudah di-uninstall dari Python environment.
echo Catatan: PATH Python Scripts tidak dihapus otomatis karena bisa dipakai
echo oleh package Python lain.
echo.
pause
