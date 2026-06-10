@echo off
chcp 65001 >nul
echo === 注册自动报价定时任务 + 看门狗 ===
echo.

:: 删除旧任务（如果存在）
schtasks /delete /tn "AutoQuoteDaily" /f >nul 2>&1

:: 创建新任务：工作日9:00执行，以当前用户身份运行
schtasks /create /tn "AutoQuoteDaily" ^
    /tr "cmd /c cd /d D:\baojia\electron\app && python auto_quote_scheduled.py" ^
    /sc weekly /d MON,TUE,WED,THU,FRI /st 09:00 ^
    /rl HIGHEST /f

if %errorlevel% equ 0 (
    echo 定时任务注册成功
) else (
    echo 定时任务注册失败！
    pause
    exit /b 1
)

:: 注册看门狗到开机启动（防止360删除定时任务）
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "AutoQuoteWatchdog" /t REG_SZ /d "wscript.exe \"D:\baojia\electron\app\watchdog_task.vbs\"" /f >nul 2>&1

if %errorlevel% equ 0 (
    echo 看门狗注册成功
) else (
    echo 看门狗注册失败！
)

echo.
echo === 完成 ===
echo 任务名称: AutoQuoteDaily
echo 执行时间: 工作日 09:00
echo 看门狗:   开机自动检查并恢复定时任务
echo.
pause
