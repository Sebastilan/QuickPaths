$mutexName = 'Global\QuickPaths_Singleton'
$m = $null
try {
    $m = [System.Threading.Mutex]::OpenExisting($mutexName)
} catch {
    Start-Process powershell.exe -ArgumentList '-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\ligon\CCA\QuickPaths\QuickPaths.ps1"' -WindowStyle Hidden
}