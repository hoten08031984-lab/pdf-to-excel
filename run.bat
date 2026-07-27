@echo off
chcp 65001 > NUL
title Quet Hoa Don PDF sang Excel
echo ==================================================
echo      QUET VA XU LY HOA DON PDF SANG EXCEL
echo ==================================================
echo.
python "%~dp0process_invoices.py"
echo.
echo ==================================================
echo Nhan phim bat ky de thoat...
pause > NUL
