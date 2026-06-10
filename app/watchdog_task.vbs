Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

TASK_NAME = "AutoQuoteDaily"
SCRIPT_DIR = "D:\baojia\electron\app"
PYTHON = "C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe"
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

' 检查定时任务是否存在
checkCmd = "cmd /c schtasks /query /tn " & TASK_NAME & " >nul 2>&1"
result = WshShell.Run(checkCmd, 0, True)

If result <> 0 Then
    ' 任务不存在，重新创建
    createCmd = "cmd /c schtasks /create /tn " & TASK_NAME & " /tr " & _
        Chr(34) & "cmd /c cd /d " & SCRIPT_DIR & " && " & _
        PYTHON & " auto_quote_scheduled.py" & Chr(34) & _
        " /sc weekly /d MON,TUE,WED,THU,FRI /st 09:00 /rl HIGHEST /f"
    WshShell.Run createCmd, 0, True
    WriteLog "定时任务 " & TASK_NAME & " 已重建"
End If
