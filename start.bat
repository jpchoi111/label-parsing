@echo off
chcp 65001 > nul
title Medit PDF Parser
cd /d %~dp0

echo 서버 시작 중...
start /b py -3.12 app.py

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