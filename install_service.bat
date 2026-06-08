@echo off
chcp 65001 >nul
echo === 安装翻译报价自动执行服务 ===
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 需要管理员权限，正在请求提权...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 安装服务
python D:\baojia\electron\app\auto_quote_service.py install
if %errorlevel% neq 0 (
    echo 安装失败！
    pause
    exit /b 1
)

:: 启动服务
sc start AutoQuoteDaemon
if %errorlevel% equ 0 (
    echo.
    echo === 服务安装并启动成功 ===
    echo 服务名称: AutoQuoteDaemon
    echo 显示名称: 翻译报价自动执行服务
    echo 日志位置: D:\baojia\electron\outputs\logs\service.log
) else (
    echo.
    echo 服务已安装，尝试启动...
    net start AutoQuoteDaemon
)

echo.
pause
