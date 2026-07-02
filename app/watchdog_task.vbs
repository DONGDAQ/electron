Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

TASK_NAME = "翻译报价每日自动报价"
PYTHON = "C:\Users\admin\.workbuddy\binaries\python\versions\3.13.12\python.exe"
SCRIPT = "D:\baojia\electron\app\auto_quote_scheduled.py"
WORK_DIR = "D:\baojia\electron\app"
PS1_PATH = "D:\baojia\electron\app\watchdog_task.ps1"
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

' 用 PowerShell 脚本处理所有任务操作，避免 cmd 编码导致乱码副本
If fso.FileExists(PS1_PATH) Then
    psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & PS1_PATH & """"
    WshShell.Run psCmd, 0, True
    WriteLog "watchdog 已通过 PowerShell 脚本执行任务检查/修复"
Else
    ' PS1 不存在时降级：用 PowerShell 直接内联检查
    checkPs = "powershell -NoProfile -Command ""$t = Get-ScheduledTask -TaskName '" & TASK_NAME & "' -ErrorAction SilentlyContinue; if (-not $t) { Write-Output 'MISSING' }"""
    result = WshShell.Run(checkPs, 0, True)
    WriteLog "watchdog: PS1 不存在，已跳过自动修复"
End If
