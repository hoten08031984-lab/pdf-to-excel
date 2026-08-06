@echo off
chcp 65001 > nul
title Quet Hoa Don PDF sang Excel
echo ==================================================
echo      QUET VA XU LY HOA DON PDF SANG EXCEL
echo ==================================================
echo.

:: Kiem tra Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [LOI] May tinh chua cai dat Python hoac chua bien Python vao PATH!
    echo Vui long cai dat Python tu https://www.python.org va tich chon "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

:: Kiem tra va tu dong cai dat thu vien neu thieu
echo Dang kiem tra thu vien Python...
python -c "import fitz, openpyxl; from rapidocr_onnxruntime import RapidOCR" >nul 2>&1
if %errorlevel% neq 0 (
    echo [THONG BAO] Dang tu dong cai dat cac thu vien phu thuoc...
    python -m pip install -r "%~dp0requirements.txt"
    if %errorlevel% neq 0 (
        echo [LOI] Cai dat thu vien that bai! Kiem tra ket noi internet.
        pause
        exit /b 1
    )
    echo.
)

:: Chay script xu ly hoa don
echo Dang chay chuong trinh xu ly hoa don...
python "%~dp0process_invoices.py"

echo.
echo ==================================================
echo Hoan tat! Nhan phim bat ky de thoat...
echo ==================================================
pause
