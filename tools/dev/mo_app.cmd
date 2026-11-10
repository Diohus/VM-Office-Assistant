@echo off
rem Cho bang ping 127.0.0.1 (khong ra mang), goi duong dan day du: timeout cua Git co the dung truoc trong PATH,
rem con timeout.exe cua Windows thoat ngay khi dau vao bi chuyen huong (do 10/11). ping -n 4 la cho khoang 3 giay.
rem Mo app sau khi khoi dong lai may, mot lan nhap. Khong doi cai dat nao cua may, khong dung lai image, khong doc .env.
rem VM_PROJECT, WEB_PORT, VM_DOCKER, VM_DOCKER_DESKTOP, VM_WAIT_SECONDS, VM_NO_BROWSER, VM_NO_PAUSE: chi de thu bang bo rieng.
setlocal
chcp 65001 > nul
cd /d "%~dp0..\.."

if "%VM_PROJECT%"=="" set "VM_PROJECT=vm-office"
if "%WEB_PORT%"=="" set "WEB_PORT=8088"
if "%VM_WAIT_SECONDS%"=="" set "VM_WAIT_SECONDS=180"
set "APP_URL=http://127.0.0.1:%WEB_PORT%"

if not exist ".local\compose.env" goto not_installed

rem Tim docker. Docker Desktop co the cai cho mot nguoi dung hoac cho ca may.
if not "%VM_DOCKER%"=="" goto have_docker
set "VM_DOCKER=docker"
where docker > nul 2>&1 && goto have_docker
set "VM_DOCKER=%LOCALAPPDATA%\Programs\DockerDesktop\resources\bin\docker.exe"
if exist "%VM_DOCKER%" goto have_docker
set "VM_DOCKER=%ProgramFiles%\Docker\Docker\resources\bin\docker.exe"
if exist "%VM_DOCKER%" goto have_docker
goto no_docker

:have_docker
echo Đang mở app, xin chờ. Không cần làm gì thêm.
call "%VM_DOCKER%" version > nul 2>&1 && goto docker_ready

echo Phần chạy nền của app ^(Docker Desktop^) chưa bật. Đang bật, có thể mất vài phút...
if not "%VM_DOCKER_DESKTOP%"=="" goto start_desktop
set "VM_DOCKER_DESKTOP=%LOCALAPPDATA%\Programs\DockerDesktop\Docker Desktop.exe"
if exist "%VM_DOCKER_DESKTOP%" goto start_desktop
set "VM_DOCKER_DESKTOP=%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
:start_desktop
if not exist "%VM_DOCKER_DESKTOP%" goto no_docker
start "" "%VM_DOCKER_DESKTOP%"

set /a WAITED=0
:wait_docker
call "%VM_DOCKER%" version > nul 2>&1 && goto docker_ready
if %WAITED% GEQ %VM_WAIT_SECONDS% goto docker_timeout
"%SystemRoot%\System32\PING.EXE" -n 4 127.0.0.1 > nul
set /a WAITED+=3
goto wait_docker

:docker_ready
rem Khong co --build: chi bat lai ban da dung. Thieu image thi bao, khong tu tai tu Internet.
call "%VM_DOCKER%" compose -p "%VM_PROJECT%" --env-file .local/compose.env up -d --wait --no-build > nul 2>&1 || goto app_failed

set /a WAITED=0
:wait_app
curl.exe -fs -o nul "%APP_URL%/api/health" && goto app_ready
if %WAITED% GEQ %VM_WAIT_SECONDS% goto app_failed
"%SystemRoot%\System32\PING.EXE" -n 3 127.0.0.1 > nul
set /a WAITED+=2
goto wait_app

:app_ready
echo App đã sẵn sàng: %APP_URL%
if "%VM_NO_BROWSER%"=="" start "" "%APP_URL%"
exit /b 0

:not_installed
echo.
echo App chưa được cài trên máy này, nên chưa mở được. Dữ liệu của anh/chị không bị đổi.
echo Hãy nhờ người cài app chạy bước cài đặt một lần.
set "RC=11"
goto fail

:no_docker
echo.
echo Máy này chưa có Docker Desktop, phần chạy nền của app, nên chưa mở được app. Dữ liệu của anh/chị không bị đổi.
echo Hãy nhờ người cài app kiểm tra Docker Desktop.
set "RC=12"
goto fail

:docker_timeout
echo.
echo Docker Desktop, phần chạy nền của app, chưa bật xong sau %VM_WAIT_SECONDS% giây, nên chưa mở được app. Dữ liệu của anh/chị không bị đổi.
echo Hãy nhìn biểu tượng cá voi ở góc dưới bên phải màn hình. Nếu có hộp thông báo thì đọc và làm theo, rồi nhấp lại biểu tượng mở app.
echo Vẫn không được thì khởi động lại máy, hoặc nhờ người cài app.
set "RC=13"
goto fail

:app_failed
echo.
echo Docker Desktop đã chạy nhưng app chưa bật được. Dữ liệu của anh/chị không bị đổi.
echo Hãy đợi một phút rồi nhấp lại biểu tượng mở app. Vẫn không được thì nhờ người cài app.
set "RC=14"
goto fail

:fail
echo.
if "%VM_NO_PAUSE%"=="" pause
exit /b %RC%
