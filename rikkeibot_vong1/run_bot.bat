@echo off
title Rikkei Bot Vong 1
echo ==================================================
echo   DANG KHOI DONG RIKKEI BOT VONG 1 (PORT 5000)...
echo ==================================================
echo.
echo 1. Dang tu dong mo Dashboard tren trinh duyet...
start "" http://localhost:5000
echo 2. Dang khoi chay Web Server...
echo (Vui long khong tat cua so nay trong qua trinh su dung bot)
echo.
python app_v1.py
if %errorlevel% neq 0 (
    echo.
    echo [LOI]: Khong the chay python. Vui long kiem tra xem Python da duoc cai dat va add vao PATH chua.
    pause
)
