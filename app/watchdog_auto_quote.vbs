Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

ROOT = "D:\baojia\electron\dist\win-unpacked\resources\app"
PYTHONW = "C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\pythonw.exe"
PYTHON = "C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe"
DAEMON = ROOT & "\auto_quote_daemon.py"
LOG_PATH = ROOT & "\outputs\logs"
Q = Chr(34)

Sub EnsureLogDir
    If Not fso.FolderExists(LOG_PATH) Then
        fso.CreateFolder(LOG_PATH)
    End If
End Sub

Sub WriteLog(msg)
    EnsureLogDir
    Set lf = fso.OpenTextFile(LOG_PATH & "\watchdog.log", 8, True)
    lf.WriteLine Date & " " & Time & " - " & msg
    lf.Close
End Sub

' 1. Start daemon if not running
daemonRunning = False
Set svc = GetObject("winmgmts:\\.\root\cimv2")
Set procs = svc.ExecQuery("SELECT * FROM Win32_Process WHERE Name='pythonw.exe'")
For Each p In procs
    If InStr(p.CommandLine, "auto_quote_daemon") > 0 Then
        daemonRunning = True
        Exit For
    End If
Next

If Not daemonRunning Then
    WshShell.Run Q & PYTHONW & Q & " " & Q & DAEMON & Q, 0, False
    WriteLog "Daemon started"
End If

' 2. Check scheduled task (backup)
checkCmd = "cmd /c schtasks /query /tn AutoQuoteDaily >nul 2>&1"
result = WshShell.Run(checkCmd, 0, True)
If result <> 0 Then
    createCmd = "cmd /c schtasks /create /tn AutoQuoteDaily /tr " & _
        Q & "cmd /c cd /d " & ROOT & " && " & _
        PYTHON & " auto_quote_scheduled.py" & Q & _
        " /sc WEEKLY /d MON,TUE,WED,THU,FRI /st 09:00 /f"
    WshShell.Run createCmd, 0, True
    WriteLog "Scheduled task recreated"
End If
