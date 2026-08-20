@echo off
chcp 65001 > nul
title Medit PDF Parser

echo ================================
echo  Medit PDF Parser 설치 및 실행
echo ================================
echo.

cd /d %~dp0

:: Python 3.12 설치 확인
py -3.12 --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Python 3.12가 없습니다. 다운로드 및 설치 중...
    curl -o python_installer.exe https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe
    python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    del python_installer.exe
    echo Python 3.12 설치 완료!
    echo.
)

:: 패키지 설치 여부 확인
if not exist ".installed" (
    echo 필요한 패키지 설치 중... 시간이 걸릴 수 있습니다.
    echo.
    py -3.12 -m pip install --upgrade pip
    py -3.12 -m pip install -r requirements.txt --no-cache-dir
    if %errorlevel% neq 0 (
        echo 패키지 설치 실패. 관리자 권한으로 다시 실행해주세요.
        pause
        exit /b 1
    )
    echo. > .installed
    echo 패키지 설치 완료!
    echo.
)

:: 서버 시작
echo 서버 시작 중...
start /b py -3.12 app.py

:: Flask 응답 대기 후 브라우저 실행
:waitloop
timeout /t 1 > nul
curl -s http://127.0.0.1:5000 > nul 2>&1
if %errorlevel% neq 0 goto waitloop

echo 서버 준비 완료!
start http://127.0.0.1:5000
echo.
echo ================================
echo  http://127.0.0.1:5000 접속됨
echo  종료하려면 이 창을 닫으세요
echo ================================
pause