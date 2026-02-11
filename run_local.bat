@echo off
chcp 65001 >nul
echo ========================================
echo   本地生产环境启动（使用当前 Python）
echo ========================================
echo.

REM 必须在启动 Python 之前设置，避免 Paddle OneDNN/PIR 报错
set FLAGS_use_mkldnn=0
set FLAGS_use_new_executor=0
echo 已设置: FLAGS_use_mkldnn=0, FLAGS_use_new_executor=0
echo.

REM 使用当前环境中的 python 启动 uvicorn
python -m uvicorn app.main:app --host 127.0.0.1 --port 8001 --reload

echo.
pause
