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