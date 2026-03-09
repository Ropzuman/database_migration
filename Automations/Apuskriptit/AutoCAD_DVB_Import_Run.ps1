# UTF-8 BOM - pakollinen PS5.1:lle
# AutoCAD DVB Re-Import skripti - RunMacro/SaveDVB -menetelmae
# Tallentaa jokaisen projektin kayttamalla valiaikaista helper-moduulia

$dvbSource = 'C:\database_migration\AutoCAD'
$exportRoot = 'C:\database_migration\AutoCAD\exported'
$migrRoot = 'C:\database_migration\AutoCAD\migrated'

if (-not (Test-Path $migrRoot)) {
    New-Item -ItemType Directory -Path $migrRoot | Out-Null
}

Write-Host ""
Write-Host "==================================================================="
Write-Host " AutoCAD DVB Re-Import (RunMacro-menetelma)"
Write-Host "==================================================================="
Write-Host ""

# Yhdistetaan kaynnissa olevaan AutoCADiin
Write-Host ($(Get-Date -Format 'HH:mm:ss') + " [IMPORT] Yhdistetaan AutoCADiin...")
try {
    $acad = [System.Runtime.InteropServices.Marshal]::GetActiveObject("AutoCAD.Application")
    Write-Host ($(Get-Date -Format 'HH:mm:ss') + " [IMPORT] AutoCAD loydetty: " + $acad.Version)
}
catch {
    Write-Host ($(Get-Date -Format 'HH:mm:ss') + " [ERROR] AutoCAD ei ole kaynnissa: " + $_.Exception.Message)
    exit 1
}

$acad.Visible = $true
$vbe = $acad.VBE

# Etsitaan globaali projekti (ensimmainen projekti VBE:ssa, yleensa "ACADProject")
# Global projekti = se johon voi tallentaa SaveDVB-kutsulla
$globalProj = $vbe.VBProjects.Item(1)
Write-Host ($(Get-Date -Format 'HH:mm:ss') + " [IMPORT] Globaali projekti: " + $globalProj.Name)

$ok = 0
$skip = 0
$fail = 0
$folders = Get-ChildItem $exportRoot -Directory | Sort-Object Name

foreach ($folder in $folders) {
    $projName = $folder.Name
    $dvbPath = Join-Path $dvbSource ($projName + '.dvb')
    $outPath = Join-Path $migrRoot  ($projName + '.dvb')
    $srcFiles = Get-ChildItem $folder.FullName -Include '*.bas', '*.cls', '*.frm' -Recurse

    Write-Host ""
    Write-Host ($(Get-Date -Format 'HH:mm:ss') + " [IMPORT] $projName ($($srcFiles.Count) tiedostoa)")

    if (-not (Test-Path $dvbPath)) {
        Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   [SKIP] DVB puuttuu")
        $skip++
        continue
    }

    if ($srcFiles.Count -eq 0) {
        Copy-Item $dvbPath $outPath -Force
        Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   Tyhja projekti - kopioitu")
        $skip++
        continue
    }

    try {
        # Ladataan DVB
        $acad.LoadDVB($dvbPath)
        Start-Sleep -Milliseconds 500

        # Etsitaan ladattu projekti nimella
        $proj = $null
        foreach ($p in $vbe.VBProjects) {
            if ($p.Name -eq $projName) { $proj = $p; break }
        }
        if ($proj -eq $null) {
            $proj = $vbe.VBProjects.Item($vbe.VBProjects.Count)
            Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   [HUOM] Projektin nimi: " + $proj.Name)
        }

        # Poistetaan vanhat komponentit (ei Document-tyyppia)
        $toRemove = @()
        foreach ($comp in $proj.VBComponents) {
            if ($comp.Type -ne 100) { $toRemove += $comp }
        }
        foreach ($comp in $toRemove) {
            Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   Poistetaan: " + $comp.Name)
            $proj.VBComponents.Remove($comp)
        }

        # Tuodaan korjatut tiedostot
        foreach ($srcFile in $srcFiles) {
            Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   Tuodaan: " + $srcFile.Name)
            $null = $proj.VBComponents.Import($srcFile.FullName)
        }

        # --- TALLENNUS RUNMACRO-MENETELMALLA ---
        # Luodaan valiaikainen helper-moduuli jossa polku on kovakoodattuna
        $helperName = "DvbSH_" + ([System.IO.Path]::GetRandomFileName() -replace '[^a-z]', '')
        $outEsc = $outPath.Replace('"', '""')
        $helperCode = "Attribute VB_Name = `"$helperName`"" + [char]13 + [char]10 +
        "Public Sub DoSave()" + [char]13 + [char]10 +
        "    Application.SaveDVB `"$outEsc`"" + [char]13 + [char]10 +
        "End Sub"
        $helperPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $helperName + ".bas")
        [System.IO.File]::WriteAllText($helperPath, $helperCode, [System.Text.Encoding]::ASCII)

        # Tuodaan helper globaaliin projektiin
        $null = $globalProj.VBComponents.Import($helperPath)
        try { Remove-Item $helperPath -Force -ErrorAction SilentlyContinue } catch {}

        # Asetetaan ladattu projekti aktiiviseksi (SaveDVB tallentaa ActiveVBProjectin)
        try { $vbe.ActiveVBProject = $proj } catch {}
        Start-Sleep -Milliseconds 300

        # Kutsutaan helper RunMacrolla (AutoCADin sisalla SaveDVB toimii)
        # Oikea muoto: "ProjektiNimi!ModuuliNimi.Aliohjelma"
        Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   Tallennetaan RunMacrolla...")
        $macroPath = $globalProj.Name + "!" + $helperName + ".DoSave"
        $acad.RunMacro($macroPath)
        Start-Sleep -Milliseconds 1200

        # Poistetaan helper globaalista projektista
        try {
            foreach ($c in @($globalProj.VBComponents)) {
                if ($c.Name -eq $helperName) { $globalProj.VBComponents.Remove($c); break }
            }
        }
        catch {}

        # Varmistetaan tulos
        if (Test-Path $outPath) {
            $sz = (Get-Item $outPath).Length
            Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   OK -> $outPath ($sz tavua)")
            $ok++
        }
        else {
            Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   [ERROR] Tiedostoa ei syntynyt") -ForegroundColor Red
            $fail++
        }

        try { $acad.UnloadDVB($dvbPath) } catch {}

    }
    catch {
        Write-Host ($(Get-Date -Format 'HH:mm:ss') + "   [ERROR] $projName : " + $_.Exception.Message) -ForegroundColor Red
        $fail++
        try { $acad.UnloadDVB($dvbPath) } catch {}
    }
}

Write-Host ""
Write-Host "==================================================================="
Write-Host " VALMIS: $ok onnistui  |  $skip ohitettu  |  $fail epaonnistui"
Write-Host " Migratoidut DVB:t: $migrRoot"
Write-Host "==================================================================="
