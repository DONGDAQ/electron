$action = New-ScheduledTaskAction -Execute "D:\baojia\electron\app\auto_quote_scheduled.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At "09:00"
Register-ScheduledTask -TaskName "WM_AutoQuote" -Action $action -Trigger $trigger -Force
Write-Host "Task created successfully"
