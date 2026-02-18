@echo off
setlocal

set FW=C:\Windows\Microsoft.NET\Framework64\v4.0.30319
set CSC=%FW%\csc.exe
if not exist "%CSC%" (
    echo ERROR: csc.exe not found at %CSC%
    exit /b 1
)

echo Building QuickPaths.exe ...
"%CSC%" /nologo /target:winexe /out:QuickPaths.exe ^
    /win32manifest:app.manifest ^
    /r:System.dll ^
    /r:System.Core.dll ^
    /r:System.Drawing.dll ^
    /r:System.Windows.Forms.dll ^
    /r:System.Runtime.Serialization.dll ^
    QuickPaths.cs

if %ERRORLEVEL% neq 0 (
    echo Build FAILED.
    exit /b 1
)
echo Build OK: QuickPaths.exe
