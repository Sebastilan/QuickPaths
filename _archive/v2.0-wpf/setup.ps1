<#
    QuickPaths Online Installer
    Usage: irm https://raw.githubusercontent.com/Sebastilan/QuickPaths/master/setup.ps1 | iex
#>

Add-Type -AssemblyName PresentationFramework

$installDir = Join-Path $env:LOCALAPPDATA 'QuickPaths'
$exeUrl = 'https://github.com/Sebastilan/QuickPaths/releases/latest/download/QuickPaths.exe'
$exePath = Join-Path $installDir 'QuickPaths.exe'

try {
    # Create install dir
    if (-not (Test-Path $installDir)) {
        New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    }

    # Download exe
    Write-Host 'Downloading QuickPaths.exe...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $exeUrl -OutFile $exePath -UseBasicParsing
    Write-Host "Downloaded to: $installDir"

    # Run install
    & $exePath --install

} catch {
    [System.Windows.MessageBox]::Show(
        "Installation failed:`n$($_.Exception.Message)",
        'QuickPaths Setup', 'OK', 'Error')
}
