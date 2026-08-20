@echo off
chcp 65001 >nul
title Tạo Shortcut Desktop - DocNormalizer Pro
color 0A

echo =======================================================================
echo     TẠO PHÍM TẮT TRUY CẬP NHANH RA MÀN HÌNH CHÍNH (DESKTOP)
echo =======================================================================
echo.
echo [*] Đang tạo Shortcut DocNormalizer Pro trên Desktop...

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $d = [System.Environment]::GetFolderPath('Desktop'); $lnk = $ws.CreateShortcut((Join-Path $d 'DocNormalizer Pro.lnk')); $lnk.TargetPath = '%~dp0Mo_Ung_Dung.bat'; $lnk.WorkingDirectory = '%~dp0'; $lnk.IconLocation = 'shell32.dll,260'; $lnk.Description = 'DocNormalizer Pro - Chuẩn Hóa Đề Thi & NĐ30'; $lnk.Save()"

if %ERRORLEVEL% equ 0 (
    echo.
    echo [THÀNH CÔNG] Đã tạo biểu tượng "DocNormalizer Pro" trên màn hình Desktop!
    echo Bạn có thể mở ứng dụng trực tiếp từ màn hình Desktop bất cứ lúc nào.
) else (
    echo.
    echo [!] Có lỗi xảy ra khi tạo Shortcut.
)

echo.
pause
exit
