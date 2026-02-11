@echo off
chcp 65001 >nul
echo ========================================
echo   重启服务器 - 局域网模式
echo ========================================
echo.

REM 查找并停止现有的uvicorn进程
echo 正在停止现有服务...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq uvicorn*" 2>nul
timeout /t 2 /nobreak >nul

echo.
echo 正在启动服务器（绑定到所有网络接口）...
echo.
echo 本机访问: http://127.0.0.1:8001
echo 局域网访问: http://192.168.31.125:8001
echo.
echo ========================================
echo.

REM 激活虚拟环境
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

REM 启动服务（关键：--host 0.0.0.0）
uvicorn app.main:app --host 0.0.0.0 --port 8001 --reload

pause

