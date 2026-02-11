# QuickPaths Watchdog - Task Scheduler setup
# Checks every 3 minutes, restarts QuickPaths if not running

$taskName = 'QuickPaths_Watchdog'
$scriptPath = 'C:\Users\ligon\CCA\QuickPaths\QuickPaths.ps1'

# Remove existing task if any
schtasks /Delete /TN $taskName /F 2>$null

# Create watchdog script
$watchdogPath = 'C:\Users\ligon\CCA\QuickPaths\watchdog.ps1'
$watchdogCode = @'
$mutexName = 'Global\QuickPaths_Singleton'
$m = $null
try {
    $m = [System.Threading.Mutex]::OpenExisting($mutexName)
    # Mutex exists = QuickPaths is running, do nothing
} catch {
    # Mutex doesn't exist = QuickPaths is not running, start it
    Start-Process powershell.exe -ArgumentList '-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\ligon\CCA\QuickPaths\QuickPaths.ps1"' -WindowStyle Hidden
}
'@
[System.IO.File]::WriteAllText($watchdogPath, $watchdogCode, (New-Object System.Text.UTF8Encoding($true)))

# Create scheduled task: run watchdog every 3 minutes
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 3)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force

Write-Host "Watchdog task '$taskName' created successfully."

