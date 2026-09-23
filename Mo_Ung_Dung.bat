@echo off
chcp 65001 >nul
title DocNormalizer Pro - Chuẩn Hóa Đề Thi
color 0B

echo =======================================================================
echo     DOCNORMALIZER PRO - CHUẨN HÓA ĐỀ THI ^& VĂN BẢN HÀNH CHÍNH NĐ 30
echo =======================================================================
echo.

:: Kiem tra neu may chu port 8000 da chay
netstat -an 2>nul | findstr ":8000 " | findstr "LISTENING" >nul
if %ERRORLEVEL% equ 0 (
    echo [V] Máy chủ cục bộ đã đang chạy trên cổng 8000.
    echo [*] Đang mở ứng dụng...
    start http://localhost:8000
    timeout /t 2 >nul
    exit
)

:: Khoi dong may chu Python neu chua chay
where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [*] Đang khởi động máy chủ cục bộ trên cổng 8000...
    start /min "" python -m http.server 8000
    timeout /t 2 >nul
    echo [V] Máy chủ đã khởi động!
    echo [*] Đang mở ứng dụng tại: http://localhost:8000
    start http://localhost:8000
    echo.
    echo [!] LƯU Ý: Không đóng cửa sổ CMD nhỏ xuất hiện ở thanh taskbar.
    echo     (Đó là máy chủ đang chạy - đóng sẽ mất kết nối đăng nhập)
    echo.
    timeout /t 5 >nul
    exit
)

:: Fallback: Node.js
where node >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [*] Đang khởi động máy chủ Node.js...
    start /min "" cmd /c "npx -y serve ""%~dp0"" -l 8000"
    timeout /t 3 >nul
    start http://localhost:8000
    timeout /t 5 >nul
    exit
)

:: Cảnh báo nếu không có Python / Node
echo.
echo [!] CẢNH BÁO: Không tìm thấy Python hoặc Node.js!
echo.
echo     Ứng dụng cần chạy qua máy chủ HTTP để đăng nhập Google hoạt động.
echo     Vui lòng cài đặt Python tại: https://www.python.org/downloads/
echo     (Tích chọn "Add Python to PATH" khi cài)
echo.
echo     Hoặc truy cập phiên bản online tại: https://docnormalizer.vercel.app
echo.
pause
