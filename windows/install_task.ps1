# 在工作排程器建立兩個工作（以目前登入的使用者身分、登入時才跑）：
#   Tardemonkey-live-day   週一到週五 08:50 起每 30 分鐘，到 13:50（日盤）
#   Tardemonkey-live-night 週一到週五 15:06 起每 30 分鐘，到隔天 05:06（夜盤）
# 時間照這台電腦的時區，請確認是台北時間（UTC+8）。重跑會覆蓋舊設定。
# 移除：Unregister-ScheduledTask -TaskName Tardemonkey-live-day,Tardemonkey-live-night -Confirm:$false
$script = Join-Path $PSScriptRoot 'run_live.ps1'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $script) -WorkingDirectory (Split-Path -Parent $PSScriptRoot)
$days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

function Make($at, $hours) {
    $t = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At $at
    $rep = New-ScheduledTaskTrigger -Once -At $at -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Hours $hours -Minutes 1)
    $t.Repetition = $rep.Repetition
    return $t
}
Register-ScheduledTask -TaskName 'Tardemonkey-live-day' -Action $action -Trigger (Make '08:50' 5) -Settings $settings -Description 'Tardemonkey 日盤即時訊號' -Force | Out-Null
Register-ScheduledTask -TaskName 'Tardemonkey-live-night' -Action $action -Trigger (Make '15:06' 14) -Settings $settings -Description 'Tardemonkey 夜盤即時訊號' -Force | Out-Null
Get-ScheduledTask -TaskName 'Tardemonkey-live-*' | Get-ScheduledTaskInfo | Select-Object TaskName, NextRunTime
