Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

TASK_NAME = "翻译报价每日自动报价"
PYTHON = "C:\Users\admin\.workbuddy\binaries\python\versions\3.13.12\python.exe"
SCRIPT = "D:\baojia\electron\app\auto_quote_scheduled.py"
WORK_DIR = "D:\baojia\electron\app"
LOG_DIR = "D:\baojia\electron\outputs\logs"

Sub EnsureLogDir
    If Not fso.FolderExists(LOG_DIR) Then
        fso.CreateFolder(LOG_DIR)
    End If
End Sub

Sub WriteLog(msg)
    EnsureLogDir
    Set lf = fso.OpenTextFile(LOG_DIR & "\watchdog.log", 8, True)
    lf.WriteLine Date & " " & Time & " - " & msg
    lf.Close
End Sub

' 检查定时任务是否存在（同时检查是否有乱码副本）
checkCmd = "cmd /c schtasks /query /tn " & TASK_NAME & " >nul 2>&1"
result = WshShell.Run(checkCmd, 0, True)

If result <> 0 Then
    ' 任务不存在，用 PowerShell 重建（避免 cmd 编码问题产生乱码）
    psCmd = "powershell -NoProfile -Command " & _
        "$t = Get-ScheduledTask -TaskName '" & TASK_NAME & "' -ErrorAction SilentlyContinue; " & _
        "if (-not $t) { " & _
        "  $action = New-ScheduledTaskAction -Execute '" & PYTHON & "' -Argument '" & SCRIPT & "' -WorkingDirectory '" & WORK_DIR & "'; " & _
        "  $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 09:00; " & _
        "  $principal = New-ScheduledTaskPrincipal -UserId 'admin' -RunLevel Limited; " & _
        "  $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew; " & _
        "  Register-ScheduledTask -TaskName '" & TASK_NAME & "' -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description '工作日每天9点自动报价' -Force; " & _
        "}"
    WshShell.Run psCmd, 0, True
    WriteLog "定时任务 " & TASK_NAME & " 已用 PowerShell 重建（避免乱码）"
End If

' 清理可能存在的乱码副本
garbledCheck = "cmd /c schtasks /query /tn " & Chr(34) & ChrW(&H7F16) & ChrW(&H8BD1) & ChrW(&H62A5) & ChrW(&HA4B9) & ChrW(&H6BCF) & ChrW(&H65E5) & ChrW(&H81EA) & ChrW(&H52A8) & ChrW(&H62A5) & ChrW(&HA4B9) & Chr(34) & " >nul 2>&1"
' 上面的乱码名不可靠，用 PowerShell 搜索并删除
cleanPsCmd = "powershell -NoProfile -Command " & _
    "$tasks = Get-ScheduledTask | Where-Object { $_.Actions.Arguments -like '*auto_quote_scheduled.py*' -and $_.TaskName -ne '" & TASK_NAME & "' }; " & _
    "foreach ($t in $tasks) { Unregister-ScheduledTask -TaskName $t.TaskName -Confirm:`$false; Write-Output ('cleaned: ' + $t.TaskName) }"
WshShell.Run cleanPsCmd, 0, True
