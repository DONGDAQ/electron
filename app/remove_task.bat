@echo off
chcp 65001 >nul
echo === 删除自动报价定时任务 ===
echo.

schtasks /delete /tn "AutoQuoteDaily" /f
if %errorlevel% equ 0 (
    echo 定时任务已删除
) else (
    echo 任务不存在或删除失败
)

pause
