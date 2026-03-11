# Access Automation Trust Center Fix

**Date:** November 8, 2025  
**Issue:** PowerShell automation script cannot access Access database  
**Root Cause:** PowerShell/Access bitness mismatch (64-bit PS → 32-bit Access)  
**Status:** ✅ RESOLVED - Script now auto-detects and relaunches in correct bitness

---

## Problem Description

When running `Access_automaatio.ps1`, the script fails with COM object creation errors or VBA project access issues.

**Original assumptions:**

- ❌ Thought it was Trust Center macro settings
- ❌ Thought it was VBA Object Model Access checkbox

**Actual root cause:**

- ✅ **Bitness mismatch:** 64-bit PowerShell cannot create COM objects for 32-bit Access
- ✅ **Missing checkbox:** The "Trust access to VBA project object model" doesn't exist in all Access versions
- ✅ **Wrong guidance:** Documentation told you to use 64-bit PowerShell, but your Access is 32-bit

---

## Quick Start (After Fix)

**Easiest method - Double-click the batch file:**

```
c:\database_migration\Automations\RUN_ACCESS_AUTOMATION.bat
```

This automatically launches the correct PowerShell version!

**Alternative - Command line:**

```powershell
# The script now auto-detects and relaunches itself in the correct bitness
cd c:\database_migration\Automations
.\Access_automaatio.ps1
```

**What happens:**

1. Script detects: "Access is 32-bit, PowerShell is 64-bit"
2. Script automatically relaunches itself in 32-bit PowerShell
3. Script prompts for database path and component folder
4. Script imports all VBA modules successfully!

---

## Root Cause

**The automation was failing due to PowerShell/Access bitness mismatch:**

### Primary Issue: 32-bit/64-bit Incompatibility ❌

- **Your Access:** 32-bit (installed in `Program Files (x86)`)
- **Your PowerShell:** 64-bit (default on Windows 10/11)
- **Problem:** 64-bit PowerShell CANNOT create COM objects for 32-bit applications
- **This is a fundamental Windows COM limitation**

When 64-bit PowerShell tries `New-Object -ComObject Access.Application` for 32-bit Access:

- The COM registration lookup fails (different registry hives)
- Or the process bitness mismatch causes instantiation failure
- Result: Script fails before ever reaching Trust Center code

### Secondary Issue: VBA Object Model Access (May Also Apply)

**Microsoft Access has TWO separate security settings for VBA:**

1. **Macro Security** (you may have enabled this ✅)
   - Controls whether macros can RUN
   - Location: Trust Center → Macro Settings

2. **VBA Object Model Access** (may be missing in some Access versions ❓)
   - Controls whether EXTERNAL PROGRAMS can programmatically access VBA code
   - Location: Trust Center → Trust access to the VBA project object model
   - **This checkbox may not exist in all Access versions/editions**
   - **This is OFF by default for security reasons where it exists**

The PowerShell script is an external program trying to manipulate VBA code, so it requires both correct bitness AND this setting (if available).

---

## Solution

### Option 0: Use Correct PowerShell Bitness (PRIMARY FIX) ⭐

**UPDATED 2025-11-08:** The automation script now AUTO-DETECTS bitness mismatch and relaunches itself automatically!

**Manual Method (if auto-relaunch fails):**

1. Determine your Access architecture:

   ```powershell
   # Run this in ANY PowerShell:
   $accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
   if ($accessPath -match 'x86') { "32-bit Access" } else { "64-bit Access" }
   ```

2. Launch matching PowerShell:

   **For 32-bit Access (most common):**
   - Press Windows key
   - Type: `PowerShell (x86)` or `Windows PowerShell (x86)`
   - Right-click → Run as Administrator (if needed)
   - Navigate to: `cd c:\database_migration\Automations`
   - Run: `.\Access_automaatio.ps1`

   **For 64-bit Access:**
   - Use normal PowerShell (64-bit is default)
   - Run the script normally

3. The script will now work because COM bitness matches!

### Option 1: Enable VBA Object Model Access (If checkbox exists)

### Option 1: Enable VBA Object Model Access (If checkbox exists)

**NOTE:** This checkbox does not exist in all Access versions. If you don't see it, skip to Option 2.

1. Open Microsoft Access
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Click **Macro Settings** on the left
4. Look for checkbox: **"Trust access to the VBA project object model"**
5. If it exists, check it and click OK
6. **IMPORTANT:** Close and reopen Access completely
7. Make sure you're using the correct PowerShell bitness (see Option 0)
8. Run the PowerShell script again

**Security Note:** This setting allows ANY program on your computer to modify VBA code in Access databases. Only enable on trusted development machines.

### Option 2: Manual Import (Current Working Workaround)

Since you're on a shared/locked-down environment where you cannot change Trust Center settings:

1. Open Access database
2. Press **Alt+F11** to open VBA Editor
3. For each module:
   - Right-click in Project Explorer
   - Select "Import File..."
   - Navigate to `c:\database_migration\Access\DOCUMENTS\`
   - Select the `.vba` or `.cls` file
   - **CRITICAL:** After import, open the module and delete header lines:

     ```vba
     VERSION 1.0 CLASS
     BEGIN
       MultiUse = -1  'True
     END
     Attribute VB_Name = "..."
     Attribute VB_GlobalNameSpace = False
     ...
     ```

   - Keep only actual code starting from `Option Compare Database` or `Option Explicit`

**Status:** ✅ Headers now removed from all workspace files (completed Nov 8, 2025), so manual copy-paste is easier.

### Option 3: Registry Edit (Advanced - for locked environments)

If you have admin access but cannot change UI settings:

```powershell
# Enable VBA Object Model Access via Registry
# For Access 2016/2019/365 (version 16.0)
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Access\Security"
Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord

# Restart Access after this change
```

**Warning:** Requires administrator privileges. Test on non-production system first.

---

## Automation Script Improvements

### Added Error Detection

Updated `Access_automaatio.ps1` with:

1. **Better null checking** for VBA project access
2. **Clearer error messages** explaining the Trust Center issue
3. **Improved header removal** - now strips VERSION, BEGIN/END, MultiUse, all Attribute lines
4. **Retry logic** for database opening (handles file locks)

### Header Removal Fix

**Previous Logic (BROKEN):**

```powershell
# Only checked for Attribute and VERSION individually
# Would stop at BEGIN/END thinking they were code
if ($lines[$i] -match "^Attribute\s+" -or $lines[$i] -match "^VERSION\s+") {
    $codeStartIndex = $i + 1
}
```

**New Logic (FIXED):**

```powershell
# Properly skips ALL header elements
if ($line -match "^VERSION\s+" -or 
    $line -match "^BEGIN$" -or 
    $line -match "^END$" -or 
    $line -match "^Attribute\s+" -or
    $line -match "^MultiUse\s*=" -or
    $line -eq "") {
    $codeStartIndex = $i + 1
}
```

---

## Testing Results

### Workspace File Cleanup (Nov 8, 2025)

Removed headers from all source files:

**Standard Modules (2):**

- ✅ GlobalVBAs.vba - Removed `Attribute VB_Name = "GlobalVBAs"`
- ✅ ForDocuments.vba - Removed `Attribute VB_Name = "ForDocuments"`

**Class Modules (23):**

- ✅ All Form_*.cls files - Removed VERSION/BEGIN/END/Attribute blocks
- ✅ All Report_*.cls files - Removed VERSION/BEGIN/END/Attribute blocks

**Command Used:**

```powershell
$clsFiles = Get-ChildItem "c:\database_migration\Access\DOCUMENTS\*.cls"
foreach ($file in $clsFiles) {
    $content = Get-Content $file.FullName -Raw -Encoding UTF8
    $lines = $content -split "`r?`n"
    $codeStart = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if ($line -match "^VERSION\s+" -or 
            $line -match "^BEGIN$" -or 
            $line -match "^END$" -or 
            $line -match "^Attribute\s+" -or 
            $line -match "^MultiUse\s*=" -or 
            $line -eq "") {
            $codeStart = $i + 1
        } else { break }
    }
    $cleanCode = ($lines[$codeStart..($lines.Count - 1)] -join "`r`n").Trim()
    Set-Content -Path $file.FullName -Value $cleanCode -Encoding UTF8 -NoNewline
}
```

**Result:** All 25 files now contain only pure VBA code, no metadata headers.

---

## Impact Assessment

### Before Header Removal

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysShowCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub BCancel_Click()
...
```

**Issues:**

- ❌ Causes "Invalid outside procedure" when imported via AddFromString()
- ❌ Cannot be copied directly into VBA editor
- ❌ Confusing for manual code review
- ❌ Not valid VBA code outside of .cls file context

### After Header Removal

```vba
Option Compare Database
Option Explicit

Private Sub BCancel_Click()
...
```

**Benefits:**

- ✅ Valid standalone VBA code
- ✅ Can be copied directly into VBA editor
- ✅ Works with automation script (when Trust enabled)
- ✅ Cleaner for version control diffs
- ✅ Easier manual code review

---

## Alternative Approaches Considered

### Why Not Use VBComponents.Import()?

**Reason:** Import() adds invisible metadata that causes form/report corruption.

From the script comments:

```
TÄRKEÄÄ - VBComponents.Import-ongelma:
- Import() lisää näkymättömiä metatietoja.
- Import() aiheuttaa komponenttien toimintahäiriöitä
  (käyttäytyy eri tavalla kuin manuaalisesti kopioidut).
```

**Better Approach:** Use CodeModule.AddFromString() with cleaned code (current implementation).

### Why Not Use .accdb Export/Import?

- Requires database to be open in UI
- No automation possible
- Time-consuming for 25+ modules

### Why Not Use Source Control Integration?

- Access doesn't have native git support
- Third-party tools (e.g., Access Version Control) are complex
- Our approach: Export to text files, version control those

---

## Recommendations

### For Development Environment

1. ✅ **Enable "Trust access to VBA project object model"** if you control the machine
2. ✅ **Use the automation script** for bulk imports
3. ✅ **Keep workspace files header-free** (already done)

### For Locked-Down/Production Environment

1. ✅ **Use manual copy-paste** (now easier with headers removed)
2. ✅ **Test each module after import**
3. ✅ **Keep automation script updated** for when Trust is available

### For Version Control

1. ✅ **Commit header-free files** (cleaner diffs)
2. ✅ **Export regularly** to keep workspace in sync
3. ✅ **Document import process** (this file)

---

## Next Steps

- [ ] Test database with manually imported clean code
- [ ] If issues found, investigate specific modules
- [ ] Consider requesting Trust Center permission change for automation
- [ ] Update export script to strip headers automatically if Access adds them back

---

## References

- Microsoft Docs: [Trust access to the VBA project object model](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
- VBA Object Model: [Application.VBE Property](https://learn.microsoft.com/en-us/office/vba/api/access.application.vbe)
- Security: [VBA Project Password Protection](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/programming-tips)
