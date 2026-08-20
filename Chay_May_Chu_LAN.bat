@echo off
chcp 65001 >nul
title DocNormalizer Pro - Chia Sẻ Mạng Cục Bộ (LAN / Wi-Fi)
color 0E

echo =======================================================================
echo     DOCNORMALIZER PRO - CHẠY MÁY CHỦ CỤC BỘ ^& CHIA SẺ MẠNG LAN
echo =======================================================================
echo.
echo [*] Đang kiểm tra môi trường Python / Node.js...

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [V] Đã tìm thấy Python.
    echo [*] Đang lấy địa chỉ IP mạng nội bộ của bạn...
    echo.
    echo -----------------------------------------------------------------------
    echo  Các máy tính và điện thoại trong cùng mạng Wi-Fi/LAN có thể truy cập:
    for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /r /c:"IPv4 Address" /c:"Địa chỉ IPv4"') do (
        echo   ==>  http:%%a:8000
    )
    echo   ==>  http://localhost:8000 (Trên máy tính này)
    echo -----------------------------------------------------------------------
    echo.
    echo [*] Đang mở trình duyệt...
    start http://localhost:8000
    echo [*] Máy chủ đang chạy. Nhấn Ctrl + C để dừng máy chủ khi không sử dụng.
    echo.
    python -m http.server 8000
    goto :eof
)

where node >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [V] Đã tìm thấy Node.js.
    echo [*] Đang khởi chạy máy chủ npx serve...
    start http://localhost:3000
    npx -y serve "%~dp0" -l 3000
    goto :eof
)

echo [!] Máy tính chưa cài sẵn Python hoặc Node.js.
echo [*] Đang mở trực tiếp file index.html qua trình duyệt...
start "" "%~dp0index.html"
timeout /t 3 >nul
exit
