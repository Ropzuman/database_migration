<#
Export-AccessVBA.ps1
Exports all VBA components from Microsoft Access database files (.accdb, .mdb) using COM automation.

Requirements and caveats:
- "Trust access to the VBA project object model" must be enabled in Access Trust Center (Macro Settings).
- The PowerShell bitness (32/64-bit) must match the installed Microsoft Office/Access bitness.
  Run `powershell -Version` or check $env:PROCESSOR_ARCHITECTURE; if Access is 32-bit, run the 32-bit PowerShell.
- You may need to run PowerShell as a user who can interact with desktop COM objects.
- Exporting programmatically requires the host to allow programmatic access to the VBIDE (see Trust Center).
- This script attempts to gracefully release COM objects; if Access processes remain, check Task Manager and kill lingering MSACCESS.EXE.

Usage examples:
# Export a single file
.
# powershell -ExecutionPolicy Bypass -File .\scripts\export_access_vba.ps1 -InputPath "C:\Data\MyDB.accdb" -OutDir "C:\temp\access_vba_exports"

# Export all Access files in a folder
# powershell -ExecutionPolicy Bypass -File .\scripts\export_access_vba.ps1 -InputPath "C:\Data\AccessDBs" -OutDir "C:\temp\access_vba_exports"

Parameters:
-InputPath : Path to a single .accdb/.mdb file or a folder containing Access files.
-OutDir    : Destination directory for exported components (default: ./access_vba_exports)
-Verbose   : Provide verbose output

Behavior:
- For each database file, creates a subfolder named using the DB filename under OutDir.
- Exports each VBComponent using the component name and a suitable extension (.bas, .cls, .frm, .txt).

# VBA alternative (to run inside Access VBA Immediate window):
# Note: Requires reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
#
# Sub ExportAllComponents(outFolder As String)
#   Dim vbProj As VBIDE.VBProject
#   Dim vbComp As VBIDE.VBComponent
#   Dim ext As String
#   If Right(outFolder, 1) <> "\" Then outFolder = outFolder & "\"
#   Set vbProj = Application.VBE.ActiveVBProject
#   For Each vbComp In vbProj.VBComponents
#     Select Case vbComp.Type
#       Case vbext_ct_StdModule: ext = ".bas"
#       Case vbext_ct_ClassModule: ext = ".cls"
#       Case vbext_ct_MSForm: ext = ".frm"
#       Case Else: ext = ".txt"
#     End Select
#     vbComp.Export outFolder & vbComp.Name & ext
#   Next vbComp
# End Sub
#
# End of VBA snippet
#>

param(
    [Parameter(Position=0)]
    [string]$InputPath,

    [Parameter(Mandatory=$false)]
    [string]$OutDir = "./access_vba_exports"

    # Use PowerShell common parameter -Verbose (Write-Verbose) instead of redefining it here
)

# --- Convenience: set a default path here if you prefer editing the script rather than passing -InputPath
# Example: $DefaultInputPath = 'C:\Data\AccessDBs' or 'C:\Data\MyDB.accdb'
$DefaultInputPath = ''

# If InputPath not provided via parameters, try DefaultInputPath; otherwise prompt interactively.
if ([string]::IsNullOrWhiteSpace($InputPath)) {
    if (-not [string]::IsNullOrWhiteSpace($DefaultInputPath)) {
        Write-Host "Using DefaultInputPath: $DefaultInputPath" -ForegroundColor Cyan
        $InputPath = $DefaultInputPath
    } else {
        $prompt = Read-Host -Prompt 'Enter path to Access file or folder (leave empty to cancel)'
        if ([string]::IsNullOrWhiteSpace($prompt)) {
            Write-Host 'No InputPath provided. Exiting.' -ForegroundColor Yellow
            exit 1
        }
        $InputPath = $prompt
    }
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
}

function Get-AccessFilesFromPath {
    param([string]$Path)
    if (Test-Path $Path -PathType Leaf) {
        return @(Get-Item $Path)
    } elseif (Test-Path $Path -PathType Container) {
        return Get-ChildItem -Path $Path -Include *.accdb, *.mdb -File -Recurse
    } else {
        throw "InputPath '$Path' does not exist."
    }
}

# Map VBComponent.Type to file extension
function Get-ExtensionForComponentType {
    param([int]$Type)
    switch ($Type) {
        1 { return '.bas' }   # vbext_ct_StdModule
        2 { return '.cls' }   # vbext_ct_ClassModule
        3 { return '.frm' }   # vbext_ct_MSForm
        100 { return '.cls' } # vbext_ct_Document (form/report modules in host projects)
        Default { return '.txt' }
    }
}

# Sanitize filename
function Safe-FileName {
    param([string]$Name)
    # replace invalid path characters with underscore
    return ($Name -replace '[\\/:*?"<>|]', '_')
}

# Main
Ensure-Directory -Path $OutDir
$files = Get-AccessFilesFromPath -Path $InputPath
if ($files.Count -eq 0) { Write-Host "No Access files found at '$InputPath'"; exit 1 }

Write-Host "Found $($files.Count) file(s) to process." -ForegroundColor Cyan

foreach ($file in $files) {
    $dbPath = $file.FullName
    $dbName = [IO.Path]::GetFileNameWithoutExtension($dbPath)
    $dbOut = Join-Path -Path (Resolve-Path $OutDir) -ChildPath $dbName
    Ensure-Directory -Path $dbOut

    Write-Host "Processing: $dbPath" -ForegroundColor Green

    $access = $null
    try {
        $access = New-Object -ComObject Access.Application
        # Open without making visible; some environments require Visible = $true
        $access.Visible = $false
        $access.OpenCurrentDatabase($dbPath)

        $vbe = $access.VBE
        if ($null -eq $vbe) {
            throw "VBE object not available. Make sure 'Trust access to the VBA project object model' is enabled."
        }

        # Try to find the VBProject that belongs to the opened database. If not found, fall back to exporting all visible projects.
        $projectList = @()
        try {
            $currentProjName = $access.CurrentProject.Name -as [string]
        } catch {
            $currentProjName = $null
        }

        if ($currentProjName) {
            $projectList = @($vbe.VBProjects | Where-Object {
                ($_.Name -eq $currentProjName) -or ($_.FileName -and ([IO.Path]::GetFileName($_.FileName) -eq $currentProjName))
            })
        }

        if ($projectList.Count -eq 0) {
            Write-Warning "Couldn't find VBProject matching current DB ('$currentProjName'). Exporting all VBProjects visible to Access."
            $projectList = $vbe.VBProjects
        }

        foreach ($vbProj in $projectList) {
            foreach ($comp in $vbProj.VBComponents) {
                $ext = Get-ExtensionForComponentType -Type $comp.Type
                $safeName = Safe-FileName -Name $comp.Name
                $dest = Join-Path -Path $dbOut -ChildPath ("$safeName$ext")
                try {
                    $comp.Export($dest)
                    Write-Verbose "  Exported $($comp.Name) -> $dest"
                } catch {
                    Write-Warning "  Failed to export $($comp.Name): $($_.Exception.Message)"
                } finally {
                    # Try to release component COM object to avoid lingering handles
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comp) | Out-Null } catch {}
                }
            }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProj) | Out-Null } catch {}
        }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbe) | Out-Null } catch {}

        # Close database
        $access.CloseCurrentDatabase()
        $access.Quit()
    } catch {
        Write-Warning "Error processing $dbPath : $($_.Exception.Message)"
    } finally {
        if ($access -ne $null) {
            try {
                # Make sure the DB is closed and Access quits
                try { $access.CloseCurrentDatabase() } catch {}
                try { $access.Quit() } catch {}
            } finally {
                # Release the COM object and force GC
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null } catch {}
                Remove-Variable access -ErrorAction SilentlyContinue
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }
    }

    Write-Host "Finished: $dbPath" -ForegroundColor Cyan
}

Write-Host "All done. Exports saved under: $(Resolve-Path $OutDir)" -ForegroundColor Yellow
