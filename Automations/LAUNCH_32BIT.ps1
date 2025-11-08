# Simple launcher to run Access automation in 32-bit PowerShell
# This script should be run from any PowerShell (it will relaunch if needed)

$scriptPath = Join-Path $PSScriptRoot "Access_automaatio.ps1"
$ps32Path = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if ([System.IntPtr]::Size -eq 8) {
    # We're in 64-bit, need to relaunch in 32-bit
    Write-Host "Relaunching in 32-bit PowerShell..." -ForegroundColor Yellow
    & $ps32Path -NoProfile -ExecutionPolicy Bypass -File $scriptPath
} else {
    # Already in 32-bit, run the script
    Write-Host "Running in 32-bit PowerShell..." -ForegroundColor Green
    & $scriptPath
}
