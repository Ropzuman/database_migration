# Excel työkalujen migraation automaatio

# Skripti vaihtaa kohdeprojektin listojen Excel-kyselyiden VBA koodin moduulit ja tekee niistä yhteensopivat 64-bittisen Officen kanssa.
# Määritä polut ylempään kohtaa tulee vaihtaa kohdeprojektin Excel-työkalujen kansion polku, yleensä muotoa \Z\tools\Projektin listojen excel-kyselyt.
# Määritä polut alempaan kohtaan tulee vaihtaa 64-bittisten moduulien polku. Tästä tullee kiinteä sijainti Y-asemalle.
# Module Names kohtaan 64-bittisten moduulien nimet.

# Muokatun tiedoston voi tallentaa muodossa .ps1 haluamaansa sijaintiin ja suorittaa seuraavasti:
# Avaa PowerShell Administratorina ja suorita komento Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass 
# Suorita "polku tiedostoon"\"tiedoston_nimi".ps1

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Määritä polut
$excelFilesPath = C:\Users\roope.vaha-aho\OneDrive - Proense Oy\Documents\Projektit\Fortum\24PRO260 Vermo Lämmönsiirrinasema\Z\tools\Projektin listojen excel-kyselyt 64bit WORK IN PROGRESS
$modulePath = C:\database_migration\Excel\Kytkentälista\Moduulit

# Module names
$moduleNames = @("Module1", "Module2", "Module3") 

Get-ChildItem -Path $excelFilesPath -Filter "*.xlsm" | ForEach-Object {
    $workbookPath = $_.FullName
    Write-Host "Käsitellään: $workbookPath"

    try {
        $workbook = $excel.Workbooks.Open($workbookPath)
        $vbaProject = $workbook.VBProject

        # 1. Poista vanhat moduulit
        foreach ($name in $moduleNames) {
            try {
                $module = $vbaProject.VBComponents.Item($name)
                $vbaProject.VBComponents.Remove($module)
                Write-Host "  - Poistettiin vanha $name"
            } catch {
                # Moduulia ei löytynyt, ei hätää
            }
        }

        # 2. Tuo uudet moduulit
        foreach ($name in $moduleNames) {
            $fullModulePath = Join-Path $modulePath "$($name).bas"
            $vbaProject.VBComponents.Import($fullModulePath)
            Write-Host "  - Tuotiin uusi $name"
        }

        # 3. Tallenna ja sulje
        $workbook.Save()
        $workbook.Close()

    } catch {
        Write-Error "Virhe käsiteltäessä $_.FullName: $($_.Exception.Message)"
        if ($workbook -ne $null) {
            $workbook.Close($false) # Sulje tallentamatta virhetilanteessa
        }
    }
}

# 4. Sulje Excel
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 
Remove-Variable excel