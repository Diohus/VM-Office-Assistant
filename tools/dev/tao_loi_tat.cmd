@echo off
rem Tao bieu tuong mo app. Mac dinh dat tren Desktop cua nguoi dung nay, co the dua thu muc khac lam tham so.
rem Chi tao mot file .lnk, khong doi cai dat nao cua may. Chay mot lan, sau khi cai app.
rem Ten file khong dau: PowerShell doc bien moi truong theo bang ma cua may, chu co dau bi hong.
setlocal
chcp 65001 > nul
set "TARGET=%~dp0mo_app.cmd"
set "FOLDER=%~1"
if "%FOLDER%"=="" for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "FOLDER=%%d"
if not exist "%FOLDER%" (
  echo Không tìm thấy thư mục để đặt biểu tượng.
  set "RC=1"
  goto end
)
set "LINK=%FOLDER%\VM-Office-Assistant.lnk"
powershell -NoProfile -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:LINK); $s.TargetPath = $env:TARGET; $s.WorkingDirectory = (Split-Path $env:TARGET); $s.Description = 'VM-Office-Assistant'; $s.Save()" || (set "RC=1" & goto end)
echo Đã tạo biểu tượng: %LINK%
set "RC=0"

:end
rem Nhấp đúp thì cửa sổ tự đóng ngay, người dùng không kịp đọc kết quả (Owner gặp 10/11). Dừng lại, trừ khi chạy tự động.
if "%VM_NO_PAUSE%"=="" pause
exit /b %RC%
