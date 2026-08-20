@echo off
chcp 65001 >nul
title DocNormalizer Pro - Chuẩn Hóa Đề Thi
color 0B

echo =======================================================================
echo     DOCNORMALIZER PRO - CHUẨN HÓA ĐỀ THI ^& VĂN BẢN HÀNH CHÍNH NĐ 30
echo =======================================================================
echo.
echo [*] Đang mở ứng dụng trên trình duyệt web mặc định...
echo.

:: Mo file index.html bang trinh duyet mac dinh
start "" "%~dp0index.html"

echo [V] Đã mở thành công!
echo [*] Ứng dụng chạy 100%% offline, an toàn và bảo mật tuyệt đối.
echo.
echo Cửa sổ này sẽ tự đóng sau 2 giây...
timeout /t 2 /nobreak >nul
exit
