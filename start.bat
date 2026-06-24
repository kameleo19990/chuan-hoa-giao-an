@echo off
title Chuẩn Hóa Giáo Án & Đề Thi
color 0B
echo.
echo  =====================================================
echo   Dang khoi dong ung dung Chuan Hoa Giao An...
echo  =====================================================
echo.

cd /d "%~dp0"

:: Mở trình duyệt sau 3 giây
start "" timeout /t 3 /nobreak >nul
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:8000"

:: Khởi chạy server (cửa sổ này giữ server sống)
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload

pause
