<#
    QuickPaths Installer
    Sets up auto-start with crash-recovery VBS wrapper and launches QuickPaths.
#>

Add-Type -AssemblyName PresentationFramework

$projectDir = $PSScriptRoot
if (-not $projectDir) { $projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path }

$mainScript = Join-Path $projectDir 'QuickPaths.ps1'

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

# --- 1. Clean up old installation (watchdog files + scheduled task) ---

$taskName = 'QuickPaths_Watchdog'
schtasks /Delete /TN $taskName /F 2>$null

$oldWatchdogPs1 = Join-Path $projectDir 'watchdog.ps1'
$oldWatchdogVbs = Join-Path $projectDir 'watchdog.vbs'
if (Test-Path $oldWatchdogPs1) { Remove-Item $oldWatchdogPs1 -Force -ErrorAction SilentlyContinue }
if (Test-Path $oldWatchdogVbs) { Remove-Item $oldWatchdogVbs -Force -ErrorAction SilentlyContinue }

# --- 2. Generate Startup VBS with crash-recovery loop ---

$startupDir = [System.IO.Path]::Combine(
    [Environment]::GetFolderPath('Startup'))  # shell:startup
$vbsPath = Join-Path $startupDir 'QuickPaths.vbs'
$lockFile = Join-Path $projectDir 'wrapper.lock'

# VBS wraps PowerShell in a Do...Loop:
#   exit 0 → user quit → stop loop
#   exit !0 → crash → 2s delay → restart
#   5 crashes in 60s → back off 30s
#   wrapper.lock prevents double-wrapping
$vbsContent = @"
Set fso = CreateObject("Scripting.FileSystemObject")
lockFile = "$lockFile"

' Prevent double wrapper
If fso.FileExists(lockFile) Then
    Set lockStream = fso.OpenTextFile(lockFile, 1)
    existingPid = Trim(lockStream.ReadLine())
    lockStream.Close
    On Error Resume Next
    Set wmi = GetObject("winmgmts:")
    Set proc = wmi.Get("Win32_Process.Handle='" & existingPid & "'")
    If Err.Number = 0 Then
        If InStr(LCase(proc.CommandLine), LCase("QuickPaths.vbs")) > 0 Then
            WScript.Quit 0
        End If
    End If
    On Error GoTo 0
End If

' Write our PID to lock
Set lockStream = fso.CreateTextFile(lockFile, True)
lockStream.WriteLine WScript.ProcessID
lockStream.Close

Set ws = CreateObject("WScript.Shell")
cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$mainScript"""

crashCount = 0
lastCrashTime = Now

Do
    exitCode = ws.Run(cmd, 0, True)
    If exitCode = 0 Then Exit Do

    ' Track crash rate
    crashCount = crashCount + 1
    If DateDiff("s", lastCrashTime, Now) > 60 Then
        crashCount = 1
        lastCrashTime = Now
    End If

    ' Back off if crashing too fast
    If crashCount >= 5 Then
        WScript.Sleep 30000
        crashCount = 0
        lastCrashTime = Now
    Else
        WScript.Sleep 2000
    End If
Loop

' Clean up lock file
On Error Resume Next
fso.DeleteFile lockFile
On Error GoTo 0
"@
[System.IO.File]::WriteAllText($vbsPath, $vbsContent, [System.Text.Encoding]::ASCII)

# --- 3. Ensure QuickPaths.ps1 has UTF-8 BOM (required for PS 5.1 + CJK) ---

$bom = New-Object System.Text.UTF8Encoding($true)
$content = [System.IO.File]::ReadAllText($mainScript, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($mainScript, $content, $bom)

# --- 4. Launch QuickPaths via VBS wrapper ---

Start-Process wscript.exe -ArgumentList "`"$vbsPath`"" -WindowStyle Hidden

# --- Done ---

[System.Windows.MessageBox]::Show(
    "QuickPaths installed successfully!`n`n" +
    "  Auto-start: Enabled (with crash recovery)`n" +
    "  Location: $projectDir`n`n" +
    "The floating dot should appear on your desktop.",
    'QuickPaths Install', 'OK', 'Information')
