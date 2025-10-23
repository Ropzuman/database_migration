#<
# Export-AccessVBA.ps1
# TARKOITUS: Viedä Access-tietokantojen (*.accdb, *.mdb) VBA-komponentit tiedostojärjestelmään tarkastelua,
# versionointia tai migraatiota varten. Skripti avaa jokaisen tietokannan ja yrittää viedä sen VBProjectin
# moduulit (standard, class, dokumentti/form-moduulit) erillisiksi .bas/.cls/.frm-tiedostoiksi.
#
# KÄYTTÄYTYMINEN:
# - Skripti pyytää syötehakemistoa, jos -InputPath-parametriä ei anneta. Se käy hakemiston läpi rekursiivisesti.
# - Kaikki *.accdb ja *.mdb tiedostot käsitellään, paitsi ne joiden nimessä esiintyy "_Backup" (case-insensitive).
# - Jokaiselle tietokannalle luodaan alihakemisto -OutDir-polun alle (DB:n nimen mukaan) ja viedään komponentit sinne.
# - Lokitus ja debug-viestit käytävät Write-Verbose; käynnistä skripti PowerShellin -Verbose -kytkimellä nähdäksesi lisätiedot.
#
# VAATIMUKSET JA HUOMAUTUKSET:
# - Ota Accessissa käyttöön "Trust access to the VBA project object model" (Trust Center -> Macro Settings).
# - Käytä PowerShell-bitnessiä, joka vastaa Office/Access -asennuksen bittisyyttä (32/64-bit).
# - Skripti vaatii työpöytäistunnon (COM-automaatio, Access.Application), eikä sovellu palvelin-CRON-tyyppiseen käyttöön.
# - Suojatut/Salasanalla suojatut VBProjectit eivät salli vientiä.
# - Jos käytät verkon polkuja (UNC, jaetut asemat), muista että verkon polkujen käsittely ja oikeudet voivat poiketa
#   paikallisista poluista; käytä UNC-polkuja (\\server\share\path) tai liitettyä asemaa ja varmista, että käyttäjällä
#   on riittävät oikeudet ja että luke/skriptaus sallitaan palvelimella.
#
# KÄYTTÖESIMERKIT:
#   Interaktiivinen (kysyy hakemiston):
#     pwsh -ExecutionPolicy Bypass -File .\Automations\export_access_vba.ps1
#
#   Ei-interaktiivinen (parametrit):
#     pwsh -ExecutionPolicy Bypass -File .\Automations\export_access_vba.ps1 -InputPath "C:\Data\AccessDBs" -OutDir "C:\temp\access_exports" -Verbose
#
# Lokit ja lisätiedot:
# - Automaatioloki: Logs/AUTOMATIONS_LOG.md
#
# Luotu: 2025-10-22
# Päivitetty: 2025-10-22 (suomi-käännös, verkko-polkuvaroitus, export-käyttäytymisen selkeytys)
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

# If InputPath not provided via parameters, try DefaultInputPath; otherwise prompt interactively for a directory.
if ([string]::IsNullOrWhiteSpace($InputPath)) {
    if (-not [string]::IsNullOrWhiteSpace($DefaultInputPath)) {
        Write-Host "Using DefaultInputPath: $DefaultInputPath" -ForegroundColor Cyan
        $InputPath = $DefaultInputPath
    } else {
        $prompt = Read-Host -Prompt 'Enter directory path to scan for Access files (subfolders included). Leave empty to cancel.'
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
        $item = Get-Item $Path
        # Skip files that include _Backup in the filename (case-insensitive)
        if ($item.Name -match '_Backup') { return @() }
        return @($item)
    } elseif (Test-Path $Path -PathType Container) {
        # Get all .accdb/.mdb files recursively and exclude ones with '_Backup' in the filename
        return Get-ChildItem -Path $Path -Include *.accdb, *.mdb -File -Recurse | Where-Object { $_.Name -notmatch '_Backup' }
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
