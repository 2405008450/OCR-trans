@echo off
chcp 65001 >nul
echo ========================================
echo   图片OCR翻译系统 - 云服务器部署
echo ========================================
echo.

REM 激活虚拟环境
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

REM 加载环境变量（如果存在.env文件）
if exist .env (
    echo 加载环境变量...
    for /f "tokens=1,2 delims==" %%a in (.env) do set %%a=%%b
)

REM 设置默认值
if "%HOST%"=="" set HOST=0.0.0.0
if "%PORT%"=="" set PORT=8001
if "%DEBUG%"=="" set DEBUG=False

echo 服务器配置:
echo   HOST: %HOST%
echo   PORT: %PORT%
echo   DEBUG: %DEBUG%
echo.

if not "%SERVER_IP%"=="" (
    echo 访问地址:
    echo   http://%SERVER_IP%:%PORT%
    echo.
)

echo 启动服务...
echo ========================================
echo.

REM 根据DEBUG模式选择启动方式
if "%DEBUG%"=="False" (
    echo 生产模式：使用多进程启动
    uvicorn app.main:app --host %HOST% --port %PORT% --workers 4 --log-level info
) else (
    echo 开发模式：单进程启动（支持热重载）
    uvicorn app.main:app --host %HOST% --port %PORT% --reload --log-level debug
)

pause

