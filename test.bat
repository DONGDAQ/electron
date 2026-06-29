@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 如果没有隐藏，通过 PowerShell 静默启动隐藏窗口
if "%_HIDDEN%"=="" (
    powershell -NoProfile -WindowStyle Hidden -Command "$env:_HIDDEN='1'; Start-Process cmd -ArgumentList '/c \"\"%~f0\"\"' -Wait"
    exit /b
)

echo ==============================================
echo          翻译报价系统 - 开发测试模式
echo ==============================================
echo.
echo 正在启动开发模式...
echo 窗口打开后按 Ctrl+Shift+I 打开开发者工具
echo 关闭此窗口即可停止服务
echo.
echo 修改代码后会自动热更新，无需重启
echo ==============================================
echo.

:: 清除可能导致 Electron 启动失败的 Node 选项
set NODE_OPTIONS=

powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -like 'python*.exe' -and $_.CommandLine -like '*quote_system.web_app*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"

:: 直接调用 electron.exe，绕过损坏的 index.js
set ELECTRON_PATH=%~dp0node_modules\electron\dist\electron.exe
"%ELECTRON_PATH%" .

pause
