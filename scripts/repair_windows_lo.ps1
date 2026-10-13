# Run this file from an elevated PowerShell (right-click -> Run with PowerShell,
# or open PowerShell as Administrator and execute it from the project folder).
$ErrorActionPreference = 'Continue'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error 'Run this script from an elevated PowerShell window.'
    exit 1
}
$projectRoot = Split-Path -Parent $PSScriptRoot
$logPath = Join-Path $projectRoot 'tmp\windows-repair.log'
New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null
Start-Transcript -Path $logPath -Force

Write-Host 'Stopping and removing the legacy XLGuard driver...'
sc.exe stop XLGuard
sc.exe config XLGuard start= disabled
sc.exe delete XLGuard
$driver = Join-Path $env:WINDIR 'System32\drivers\XLGuard.sys'
if (Test-Path -LiteralPath $driver) {
    takeown.exe /F $driver
    icacls.exe $driver /grant Administrators:F
    Remove-Item -LiteralPath $driver -Force -ErrorAction Stop
    Write-Host 'Driver file removed.'
}

Write-Host 'Repairing the Windows component store (DISM)...'
DISM.exe /Online /Cleanup-Image /RestoreHealth

Write-Host 'Checking protected system files (SFC)...'
sfc.exe /scannow

Write-Host 'Completed. Restart Windows before testing LibreOffice again.'
Stop-Transcript
