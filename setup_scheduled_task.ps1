$python = "C:\Users\admin\.workbuddy\binaries\python\versions\3.13.12\python.exe"
$script = "D:\baojia\electron\app\auto_quote_scheduled.py"

# 1. 创建定时任务（工作日每天9点）
$action = New-ScheduledTaskAction -Execute $python -Argument $script -WorkingDirectory "D:\baojia\electron\app"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 09:00
$principal = New-ScheduledTaskPrincipal -UserId "admin" -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "翻译报价每日自动报价" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "工作日每天9点自动报价(完美世界3项目+战双2项目+TK+4399)" -Force
Write-Host "定时任务已创建: 翻译报价每日自动报价"

# 2. 注册看门狗到开机启动（防止任务被删）
$watchdogPath = "D:\baojia\electron\app\watchdog_task.vbs"
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "AutoQuoteWatchdog" -Value "wscript.exe `"$watchdogPath`"" -Force
Write-Host "看门狗已注册: AutoQuoteWatchdog"
Write-Host "完成"
