@echo off
chcp 65001 >nul
title Excel Agent Online - 本地Excel处理Agent

echo ============================================
echo   Excel Agent Online
echo   正在启动服务...
echo   访问地址: http://localhost:5000
echo ============================================
echo.

:: 延迟2秒后自动打开浏览器
start "" http://localhost:5000

:: 自动安装依赖
echo 正在检查依赖...
pip install -q -r requirements.txt
echo.

:: 启动Flask服务
python app.py

echo.
echo 服务已停止。
pause
