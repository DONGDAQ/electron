@echo off
chcp 65001 >nul
echo === 卸载翻译报价自动执行服务 ===
echo.

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 需要管理员权限，正在请求提权...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

sc stop AutoQuoteDaemon >nul 2>&1
python D:\baojia\electron\app\auto_quote_service.py remove
if %errorlevel% equ 0 (
    echo 服务已卸载
) else (
    echo 卸载失败
)
pause
