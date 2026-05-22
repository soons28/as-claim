@echo off
title 채권신고 취하서 자동 타이핑 시스템
echo ========================================================
echo   채권신고 취하서 자동 타이핑 시스템을 기동하고 있습니다...
echo ========================================================
echo.

:: 1. 조금 뒤에 브라우저 실행
timeout /t 1 /nobreak > nul
start http://localhost:3000

:: 2. Node.js 서버 실행
node server.js

if %errorlevel% neq 0 (
    echo.
    echo [오류] Node.js 설치 상태를 확인해 주세요.
    pause
)
