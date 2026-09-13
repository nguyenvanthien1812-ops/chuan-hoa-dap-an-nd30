@echo off
chcp 65001 >nul
title DocNormalizer Pro - Chuẩn Hóa Đề Thi
color 0B

echo =======================================================================
echo     DOCNORMALIZER PRO - CHUẨN HÓA ĐỀ THI ^& VĂN BẢN HÀNH CHÍNH NĐ 30
echo =======================================================================
echo.
echo [*] Đang mở ứng dụng trên trình duyệt web...
echo.

:: 1. Thu mo qua may chu cuc bo neu dang chay
start http://localhost:8000 2>nul

:: 2. Mo truc tiep file index.html qua trinh duyet mac dinh
start "" "%~dp0index.html"

echo [V] Đã gửi lệnh mở ứng dụng!
echo [*] Địa chỉ truy cập: http://localhost:8000
echo [*] Ứng dụng chạy 100%% offline, an toàn và bảo mật tuyệt đối.
echo.
echo Nhấn phím bất kỳ hoặc đóng cửa sổ này khi đã vào được ứng dụng...
timeout /t 5 >nul
exit
