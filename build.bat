@echo off
setlocal
title ARCDVNpdf Build

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")

echo Installing build tools...
%PY% -m pip install --upgrade build wheel

echo.
echo Building wheel...
%PY% -m build

echo.
echo Build selesai. File ada di folder dist.
pause
