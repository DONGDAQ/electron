$python = "C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe"
$script = "D:\baojia\electron\app\auto_quote_scheduled.py"
$action = New-ScheduledTaskAction -Execute $python -Argument $script -WorkingDirectory "D:\baojia\electron\app"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 09:00
$principal = New-ScheduledTaskPrincipal -UserId "admin" -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "翻译报价每日自动报价" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "工作日每天9点自动报价(完美世界3项目+战双2项目+4399)" -Force
Write-Host "任务已创建"
