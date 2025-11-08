# Minimal test script - just check bitness and Access detection
Write-Host "=== BITNESS TEST ===" -ForegroundColor Cyan

# Check Access
$accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
$accessIs32Bit = $false

if ($accessPath -and ($accessPath -match 'x86' -or $accessPath -match 'Program Files \(x86\)')) {
    $accessIs32Bit = $true
}

if (-not $accessPath) {
    $accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
    if ($accessPath) {
        $accessIs32Bit = $true
    }
}

# Check PowerShell
$psIs64Bit = [System.IntPtr]::Size -eq 8

# Display
Write-Host "Access Path: $accessPath"
Write-Host "Access is 32-bit: $accessIs32Bit"
Write-Host "PowerShell IntPtr.Size: $([System.IntPtr]::Size)"
Write-Host "PowerShell is 64-bit: $psIs64Bit"
Write-Host ""

if ($accessIs32Bit -and $psIs64Bit) {
    Write-Host "MISMATCH: 32-bit Access with 64-bit PowerShell" -ForegroundColor Red
} elseif (-not $accessIs32Bit -and -not $psIs64Bit) {
    Write-Host "MISMATCH: 64-bit Access with 32-bit PowerShell" -ForegroundColor Red
} else {
    Write-Host "MATCH: Bitness is compatible!" -ForegroundColor Green
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
