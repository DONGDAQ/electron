' QuoteTool Dev Launcher - double-click to open the app directly
On Error Resume Next

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\baojia\electron"

' Clear env vars that would break electron (set by some IDEs)
WshShell.Environment("Process").Remove("ELECTRON_RUN_AS_NODE")
WshShell.Environment("Process").Remove("NODE_OPTIONS")

' Kill stale flask_server (if any) - silent, no error popup
WshShell.Run "taskkill /F /IM flask_server.exe", 0, True

' Start Electron in dev mode (loads from source, no rebuild needed)
WshShell.Run """D:\baojia\electron\node_modules\electron\dist\electron.exe"" ""D:\baojia\electron""", 0, False
