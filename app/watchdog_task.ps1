# watchdog_task.ps1
# 用 PowerShell 原生 Cmdlet 管理定时任务，避免 cmd 编码导致乱码副本
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File "watchdog_task.ps1"

param(
    [string]$TaskName = "翻译报价每日自动报价",
    [string]$Python   = "C:\Users\admin\.workbuddy\binaries\python\versions\3.13.12\python.exe",
    [string]$Script   = "D:\baojia\electron\app\auto_quote_scheduled.py",
    [string]$WorkDir = "D:\baojia\electron\app"
)

$LOG_DIR = "D:\baojia\electron\outputs\logs"
$LOG_FILE = "$LOG_DIR\watchdog.log"

# 日志函数
function Write-Log($msg) {
    if (-not (Test-Path $LOG_DIR)) { New-Item -ItemType Directory -Path $LOG_DIR -Force | Out-Null }
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

# 1. 清理其他指向 auto_quote_scheduled.py 的乱码/重复任务
$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
    $_.Actions.Arguments -like '*auto_quote_scheduled.py*' -and $_.TaskName -ne $TaskName
}
foreach ($t in $tasks) {
    try {
        Unregister-ScheduledTask -TaskName $t.TaskName -Confirm:$false
        Write-Log "已删除重复/乱码任务: $($t.TaskName)"
    } catch {
        Write-Log "删除任务失败: $($t.TaskName) - $_"
    }
}

# 2. 检查主任务是否存在，不存在则创建
$mainTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $mainTask) {
    $action   = New-ScheduledTaskAction -Execute $Python -Argument $Script -WorkingDirectory $WorkDir
    $trigger   = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "09:00"
    $principal = New-ScheduledTaskPrincipal -UserId "admin" -RunLevel Limited
    $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -WakeToRun
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "工作日每天9点自动报价" -Force
    Write-Log "定时任务 '$TaskName' 已创建"
} else {
    Write-Log "定时任务 '$TaskName' 已存在，无需操作"
}

Write-Log "watchdog 检查完成"
