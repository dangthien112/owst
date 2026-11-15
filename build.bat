@echo off
chcp 65001 >nul
setlocal

echo ============================================
echo   Dong goi Order2PDF thanh file .exe
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Khong tim thay python. Cai Python 3.11+ tu python.org
    echo       va nho tich chon "Add Python to PATH".
    pause
    exit /b 1
)

if not exist .venv (
    echo [1/4] Tao moi truong ao...
    python -m venv .venv
)
call .venv\Scripts\activate.bat

echo [2/4] Cai thu vien...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt pyinstaller

echo [3/4] Tao icon...
if not exist assets\icon.ico (
    python tools\make_icon.py
)

echo [4/4] Dong goi...
set ICONARG=
if exist assets\icon.ico set ICONARG=--icon "assets\icon.ico"

pyinstaller --noconfirm --clean --onefile --windowed --name Order2PDF ^
    %ICONARG% ^
    --add-data "fonts;fonts" ^
    --add-data "products.json;." ^
    --add-data "config.json;." ^
    app.py

if errorlevel 1 (
    echo.
    echo [LOI] Dong goi that bai. Xem log phia tren.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   XONG:  dist\Order2PDF.exe
echo   Copy file .exe ra Desktop hoac thu muc bat ky.
echo ============================================
echo.
pause
