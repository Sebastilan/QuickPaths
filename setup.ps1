<#
    QuickPaths Online Installer
    Usage: irm https://raw.githubusercontent.com/Sebastilan/QuickPaths/master/setup.ps1 | iex
#>

Add-Type -AssemblyName PresentationFramework

$installDir = Join-Path $env:LOCALAPPDATA 'QuickPaths'
$zipUrl = 'https://github.com/Sebastilan/QuickPaths/archive/refs/heads/master.zip'
$zipPath = Join-Path $env:TEMP 'QuickPaths_download.zip'
$extractDir = Join-Path $env:TEMP 'QuickPaths_extract'

try {
    # Download
    Write-Host 'Downloading QuickPaths...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing

    # Extract
    Write-Host 'Extracting...'
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $extractDir

    # GitHub ZIP contains a folder named QuickPaths-master
    $source = Get-ChildItem $extractDir -Directory | Select-Object -First 1

    # Copy to install location
    if (Test-Path $installDir) { Remove-Item $installDir -Recurse -Force }
    Copy-Item $source.FullName $installDir -Recurse
    Write-Host "Installed to: $installDir"

    # Run local installer
    & (Join-Path $installDir 'install.ps1')

} catch {
    [System.Windows.MessageBox]::Show(
        "Installation failed:`n$($_.Exception.Message)",
        'QuickPaths Setup', 'OK', 'Error')
} finally {
    # Cleanup temp files
    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
}
