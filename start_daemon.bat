@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 隐藏窗口运行
if "%_HIDDEN%"=="" (
    powershell -NoProfile -WindowStyle Hidden -Command "$env:_HIDDEN='1'; Start-Process cmd -ArgumentList '/c \"\"%~f0\"\"' -Wait"
    exit /b
)

echo 启动自动报价守护进程...
python app\auto_quote_daemon.py
