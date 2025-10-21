# Excel työkalujen migraation automaatio

# Skripti vaihtaa kohdeprojektin listojen Excel-kyselyiden VBA koodin moduulit ja tekee niistä yhteensopivat 64-bittisen Officen kanssa.
# Määritä polut ylempään kohtaa tulee vaihtaa kohdeprojektin Excel-työkalujen kansion polku, yleensä muotoa \Z\tools\Projektin listojen excel-kyselyt.
# Määritä polut alempaan kohtaan tulee vaihtaa 64-bittisten moduulien polku. Tästä tullee kiinteä sijainti Y-asemalle.
# Module Names kohtaan 64-bittisten moduulien nimet.

# Muokatun tiedoston voi tallentaa muodossa .ps1 haluamaansa sijaintiin ja suorittaa seuraavasti:
# Avaa PowerShell Administratorina ja suorita komento Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass 
# Suorita "polku tiedostoon"\"tiedoston_nimi".ps1

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Määritä polut (Polut on suojattu lainausmerkeillä välilyöntien takia)
$excelFilesPath = "\\proense01\projektit\24PRO260 Vermo Lämmönsiirrinasema\Z\tools\Projektin listojen excel-kyselyt 64bit WORK IN PROGRESS"
$modulePath = "C:\database_migration\Excel\Kytkentälista\Moduulit"

# Module names
$moduleNames = @("Module1", "Module2", "Module3") 

# Retry-asetukset OneDrive-lukkojen kiertämiseksi
$maxRetries = 3
$retryDelaySeconds = 1

# Alusta työkirja-muuttuja
$workbook = $null

# Käsittele kaikki .xlsm tiedostot kohdekansiossa
Get-ChildItem -Path $excelFilesPath -Filter "*.xlsm" | ForEach-Object {
    $workbookPath = $_.FullName
    Write-Host "--- KÄSITTELLÄÄN: $workbookPath ---"

    $retryCount = 0
    $isOpened = $false
    $workbook = $null
    
    # 1. Poistaa tiedoston Vain luku -attribuutin (IsReadOnly = $false).
    Set-ItemProperty -Path $workbookPath -Name IsReadOnly -Value $false -Force
    Write-Host "  ✅ Poistettiin Vain luku -attribuutti tiedostojärjestelmästä."

    # 2. Avaa työkirja (Retry-logiikalla lukkojen kiertämiseksi)
    do {
        try {
            # Yksinkertaistettu kutsu COM-virheen välttämiseksi
            # Parametrit: (FileName, UpdateLinks, ReadOnly)
            $workbook = $excel.Workbooks.Open(
                $workbookPath, 
                $false,       # 2: UpdateLinks (False, ei null)
                $false        # 3: ReadOnly (Pakotetaan avaamaan kirjoitusoikeudella)
            )
            $isOpened = $true
            Write-Host "  - Työkirja avattu onnistuneesti."
        } catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "  ⚠️ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen $retryDelaySeconds sekunnin kuluttua (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            } else {
                Write-Error "  ❌ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen. Jätetään käsittelemättä."
                $isOpened = $false
                # Nosta virhe, jotta päästään pää-catch-lohkoon
                throw $_
            }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)
    
    # Käsittely jatkuu vain, jos tiedosto avattiin onnistuneesti
    if ($isOpened) {
        
        try {
            # Tarkista, onko VBA-projekti käytettävissä (riippuu Trust Center -asetuksista)
            $vbaProject = $workbook.VBProject
            Write-Host "  - VBA-projekti avattu."

            # 3. Poista vanhat moduulit
            foreach ($name in $moduleNames) {
                try {
                    $module = $vbaProject.VBComponents.Item($name)
                    $vbaProject.VBComponents.Remove($module)
                    Write-Host "  ✅ Poistettiin vanha $name"
                } catch {
                    # Hiljainen virhe, jos moduulia ei löydy
                }
            }

            # 4. Tuo uudet moduulit
            foreach ($name in $moduleNames) {
                $fullModulePath = Join-Path $modulePath "$($name).bas"
                
                if (-not (Test-Path $fullModulePath)) {
                    Write-Error "  ❌ VIRHE: Uutta moduulitiedostoa $fullModulePath ei löydy. Ohitetaan tuonti."
                    continue
                }
                
                try {
                    $vbaProject.VBComponents.Import($fullModulePath)
                    Write-Host "  ✅ Tuotiin uusi $name"
                } catch {
                    Write-Error "  ❌ VIRHE: Moduulin $name tuonti epäonnistui: $($_.Exception.Message)"
                }
            }

            # 5. Tallenna samaan polkuun eri nimellä, poista vanha ja nimeä uusi uudelleen
            $tempSuffix = "_MIGRATED"
            $tempWorkbookPath = $workbookPath.Replace(".xlsm", "$tempSuffix.xlsm")
            
            Write-Host "  - Tallennetaan väliaikaiseen tiedostoon: $tempWorkbookPath"
            
            # Tallenna muutettu työkirja väliaikaiseen tiedostoon
            # TÄMÄ ON NYT ALKUPERÄINEN TALLENNUS, joka käyttää uutta nimeä
            $workbook.SaveAs($tempWorkbookPath) 
            Write-Host "  ✅ Väliaikainen tallennus onnistui."

            # Sulje työkirja (TÄRKEÄÄ: tiedosto on suljettava ennen tiedostojärjestelmän operaatioita)
            $workbook.Close()
            $workbook = $null # Nollaa muuttuja

            # Poista alkuperäinen tiedosto ja nimeä uusi uudelleen
            Write-Host "  - Poistetaan alkuperäinen tiedosto..."
            Remove-Item -Path $workbookPath -Force -ErrorAction Stop
            Write-Host "  ✅ Alkuperäinen poistettu."

            Write-Host "  - Nimetään uusi tiedosto alkuperäiseksi..."
            # Käytetään Split-Path -Leaf varmistaaksemme, että NewName on vain tiedoston nimi
            Rename-Item -Path $tempWorkbookPath -NewName (Split-Path $workbookPath -Leaf) -Force -ErrorAction Stop
            Write-Host "  ✅ Tiedosto $workbookPath päivitetty onnistuneesti korvausmenetelmällä."

        } catch {
            # Virheenkäsittely
            Write-Error "❌ Virhe VBA-käsittelyssä tai tallennuksessa/korvauksessa: $($_.Exception.Message)"
            
            # HUOM: Jos $workbook on null, se tarkoittaa, että se on jo suljettu/nollattu onnistuneen Save/Close-syklin aikana.
            if ($workbook -ne $null) {
                Write-Host "  ⚠️ Suljetaan työkirja tallentamatta virhetilanteen vuoksi."
                $workbook.Close($false) 
            }
        }
    } # end if ($isOpened)
} # end ForEach-Object

# 6. Sulje Excel
Write-Host "--- Suljetaan Excel-prosessi ---"
$excel.Quit()

# Vapauta COM-objekti muistista
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 
Remove-Variable excel -ErrorAction SilentlyContinue


#--------------------------------- Skripti päättyy ----------------------------------