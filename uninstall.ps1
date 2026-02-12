<#
    QuickPaths Uninstaller
    Removes auto-start, watchdog task, and optionally user data.
#>

Add-Type -AssemblyName PresentationFramework

$projectDir = $PSScriptRoot
if (-not $projectDir) { $projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path }

# --- 1. Stop running QuickPaths ---

$mutexName = 'Global\QuickPaths_Singleton'
$found = $false
try {
    $m = [System.Threading.Mutex]::OpenExisting($mutexName)
    $found = $true
    $m.Close()
} catch {}

if ($found) {
    # Find and stop QuickPaths powershell processes
    Get-Process powershell -ErrorAction SilentlyContinue | Where-Object {
        try {
            $wmi = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue
            $wmi.CommandLine -like '*QuickPaths.ps1*'
        } catch { $false }
    } | ForEach-Object {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
}

# --- 2. Remove Startup VBS ---

$vbsPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'QuickPaths.vbs'
if (Test-Path $vbsPath) {
    Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue
}

# --- 3. Remove Task Scheduler watchdog ---

$taskName = 'QuickPaths_Watchdog'
schtasks /Delete /TN $taskName /F 2>$null

# --- 4. Ask about user data ---

$r = [System.Windows.MessageBox]::Show(
    "Delete user data (saved paths and window position)?`n`n" +
    "  paths.json - your saved path list`n" +
    "  config.json - window position`n`n" +
    "Click Yes to delete, No to keep.",
    'QuickPaths Uninstall', 'YesNo', 'Question')

if ($r -eq 'Yes') {
    $dataFile = Join-Path $projectDir 'paths.json'
    $configFile = Join-Path $projectDir 'config.json'
    if (Test-Path $dataFile) { Remove-Item $dataFile -Force -ErrorAction SilentlyContinue }
    if (Test-Path $configFile) { Remove-Item $configFile -Force -ErrorAction SilentlyContinue }
}

# --- Done ---

[System.Windows.MessageBox]::Show(
    "QuickPaths uninstalled.`n`n" +
    "  Auto-start: Removed`n" +
    "  Watchdog task: Removed`n" +
    "  Program files: Kept (delete manually if desired)",
    'QuickPaths Uninstall', 'OK', 'Information')
