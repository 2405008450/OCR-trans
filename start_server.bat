@echo off
chcp 65001 >nul
echo ========================================
echo   图片OCR翻译系统 - 局域网部署
echo ========================================
echo.
echo 正在启动服务器...
echo.
echo 本机访问地址:
echo   http://127.0.0.1:8001
echo   http://localhost:8001
echo.
echo 局域网访问地址:
echo   http://192.168.31.125:8001
echo.
echo 其他设备请使用上述局域网地址访问
echo.
echo ========================================
echo.

REM 激活虚拟环境（如果存在）
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

REM 启动服务（绑定到所有网络接口）
uvicorn app.main:app --host 0.0.0.0 --port 8001 --reload

pause


