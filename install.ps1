<#
    QuickPaths Installer
    Automatically sets up auto-start, watchdog, and launches QuickPaths.
#>

Add-Type -AssemblyName PresentationFramework

$projectDir = $PSScriptRoot
if (-not $projectDir) { $projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path }

$mainScript = Join-Path $projectDir 'QuickPaths.ps1'
$watchdogScript = Join-Path $projectDir 'watchdog.ps1'

# --- Pre-checks ---

if (-not (Test-Path $mainScript)) {
    [System.Windows.MessageBox]::Show(
        "QuickPaths.ps1 not found in:`n$projectDir`n`nPlease run the installer from the QuickPaths folder.",
        'QuickPaths Install', 'OK', 'Error')
    exit 1
}

if ($PSVersionTable.PSVersion.Major -lt 5) {
    [System.Windows.MessageBox]::Show(
        "PowerShell 5.1 or later is required.`nCurrent version: $($PSVersionTable.PSVersion)",
        'QuickPaths Install', 'OK', 'Error')
    exit 1
}

# --- 1. Generate Startup VBS (auto-start on boot) ---

$startupDir = [System.IO.Path]::Combine(
    [Environment]::GetFolderPath('Startup'))  # shell:startup
$vbsPath = Join-Path $startupDir 'QuickPaths.vbs'

$vbsContent = @"
Set ws = CreateObject("WScript.Shell")
ws.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$mainScript""", 0, False
"@
[System.IO.File]::WriteAllText($vbsPath, $vbsContent, [System.Text.Encoding]::ASCII)

# --- 2. Generate watchdog.ps1 (dynamic path) ---

$watchdogCode = @'
$mutexName = 'Global\QuickPaths_Singleton'
$m = $null
try {
    $m = [System.Threading.Mutex]::OpenExisting($mutexName)
} catch {
    $dir = $PSScriptRoot
    if (-not $dir) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $ps1 = Join-Path $dir 'QuickPaths.ps1'
    Start-Process powershell.exe -ArgumentList ('-ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $ps1 + '"') -WindowStyle Hidden
}
'@
$bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($watchdogScript, $watchdogCode, $bom)

# --- 3. Register Task Scheduler watchdog (current user, no admin) ---

$taskName = 'QuickPaths_Watchdog'

# Remove existing task if any
schtasks /Delete /TN $taskName /F 2>$null

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogScript`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 3)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME `
    -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null

# --- 4. Ensure QuickPaths.ps1 has UTF-8 BOM (required for PS 5.1 + CJK) ---

$content = [System.IO.File]::ReadAllText($mainScript, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($mainScript, $content, $bom)

# --- 5. Launch QuickPaths ---

Start-Process powershell.exe `
    -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScript`"" `
    -WindowStyle Hidden

# --- Done ---

[System.Windows.MessageBox]::Show(
    "QuickPaths installed successfully!`n`n" +
    "  Auto-start: Enabled`n" +
    "  Watchdog: Every 3 minutes`n" +
    "  Location: $projectDir`n`n" +
    "The floating dot should appear on your desktop.",
    'QuickPaths Install', 'OK', 'Information')
