# Access VBA Automation - Quick Start Guide

## Problem Summary

- Your Access is **32-bit** (Program Files x86)
- Default PowerShell is **64-bit**
- 64-bit PowerShell **cannot** create COM objects for 32-bit Access

## Solution: Use the Batch File

### Easiest Method ⭐

**Double-click this file:**

```
c:\database_migration\Automations\RUN_ACCESS_AUTOMATION.bat
```

This automatically:

1. Launches 32-bit PowerShell
2. Bypasses execution policy
3. Runs the automation script

### What It Will Ask

1. **Database path** - Path to your .accdb file
2. **Component path** - `c:\database_migration\Access\DOCUMENTS\`

### Manual Method (If Batch Fails)

1. Open Start Menu
2. Type: `PowerShell (x86)` or `Windows PowerShell (x86)`
3. Right-click and "Run as Administrator" (if needed)
4. Run these commands:

   ```powershell
   cd c:\database_migration\Automations
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   .\Access_automaatio.ps1
   ```

## Testing Bitness

Run this to verify your setup:

```powershell
.\TEST_BITNESS.ps1
```

Should show:

- Access: 32-bit ✅
- PowerShell: 32-bit ✅  
- MATCH ✅

## Common Errors

### "Script execution is disabled"

**Fix:** The batch file uses `-ExecutionPolicy Bypass` automatically

### "Output stream already redirected"  

**Fix:** Updated Nov 9, 2025 - replaced `>` with `-` in error messages

### "Bitness mismatch"

**Fix:** Use the batch file or manually launch PowerShell (x86)

## Files Created Nov 9, 2025

- `RUN_ACCESS_AUTOMATION.bat` - Main launcher (USE THIS)
- `TEST_BITNESS.ps1` - Test your bitness setup
- `LAUNCH_32BIT.ps1` - Alternative PowerShell launcher
- `Access_automaatio.ps1` - Main automation script (updated with fixes)

## What the Script Does

1. ✅ Detects Access and PowerShell architecture
2. ✅ Opens your Access database
3. ✅ Imports all 26 VBA/CLS modules from DOCUMENTS folder
4. ✅ Strips headers automatically (already done in workspace)
5. ✅ Saves and closes safely

## Need Help?

Check the detailed documentation:

```
c:\database_migration\Logs\ACCESS_AUTOMATION_TRUST_FIX.md
```
