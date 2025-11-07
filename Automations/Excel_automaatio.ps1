# Excel_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit listojen kysely työkaluissa 64-bittisillä moduuleilla.

# KÄYTTÄYTYMINEN:
# - Kysyy polut työkirjojen hakemistoon ja moduulitiedostojen hakemistoon (ellei oletuksia ole asetettu).
# - Avaa jokaisen .xlsm-työkirjan, päivittää moduulien sisällön suoraan, tallentaa väliaikaisesti ja korvaa alkuperäisen.
# - Käyttää retry-logiikkaa lukkojen kiertämiseksi (OneDrive ym.).

# TÄRKEÄÄ - VBComponents.Import-ongelma:
# - Tämä skripti KORVAA moduulien sisällön suoraan CodeModule-rajapinnan kautta.
# - EI käytetä VBComponents.Import()-funktiota, koska se lisää näkymättömiä metatietoja moduuleihin.
# - Import() aiheuttaa moduulien toimintahäiriöitä (käyttäytyy eri tavalla kuin manuaalisesti kopioidut).
# - Nykyinen toteutus: lue .bas-tiedosto → poista headerit → kirjoita puhdas koodi AddFromString()-funktiolla.
# - Tämä vastaa manuaalista kopioi-liitä -toimintoa VBA-editorissa.

# - HUOM: Kun muutoksia ajetaan verkkosijaintiin, pitää käyttää verkkosijainnin nimeä \\proense01\projektit\ 
# - esim. "\\proense01\projektit\24PRO260 Vermo Lämmönsiirrinasema\Z\tools\Projektin listojen excel-kyselyt 64bit WORK IN PROGRESS"

# Muokatun tiedoston voi tallentaa muodossa .ps1 haluamaansa sijaintiin ja suorittaa seuraavasti:
# - Avaa PowerShell Administratorina ja suorita komento Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# - Suorita "polku tiedostoon"\"tiedoston_nimi".ps1

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# --- Polut ja asetukset (Päivitä oletusarvot tähän halutessasi. Varsinkin uusien moduulien sijainti yleensä kiinteä.) ---
$DefaultExcelFilesPath = ''
$DefaultModulePath = ''

Write-Host "Excel files folder: $DefaultExcelFilesPath" -ForegroundColor Cyan
$inputExcel = Read-Host -Prompt 'Lisää polku Excel-tiedostoille (paina Enter käyttääksesi oletusta)'
if ([string]::IsNullOrWhiteSpace($inputExcel)) { $excelFilesPath = $DefaultExcelFilesPath } else { $excelFilesPath = $inputExcel }

Write-Host "Module files folder: $DefaultModulePath" -ForegroundColor Cyan
$inputModule = Read-Host -Prompt 'Lisää polku moduulitiedostoille (paina Enter käyttääksesi oletusta)'
if ([string]::IsNullOrWhiteSpace($inputModule)) { $modulePath = $DefaultModulePath } else { $modulePath = $inputModule }

# Validate paths
if (-not (Test-Path $excelFilesPath -PathType Container)) {
    Write-Error "Excel files folder does not exist: $excelFilesPath"
    exit 1
}
if (-not (Test-Path $modulePath -PathType Container)) {
    Write-Error "Module files folder does not exist: $modulePath"
    exit 1
}

# Moduulien nimet.
# Määrittele kaikki ne moduulit, jotka poistetaan ja tuodaan uudelleen.
# Älä sisällytä tiedostopäätettä (.bas) nimiin.
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
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "  ⚠️ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen $retryDelaySeconds sekunnin kuluttua (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
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

            # 3. Päivitä moduulien sisältö suoraan (välttää Import-metatietojen ongelman)
            foreach ($name in $moduleNames) {
                $fullModulePath = Join-Path $modulePath "$($name).bas"
                
                if (-not (Test-Path $fullModulePath)) {
                    Write-Error "  ❌ VIRHE: Uutta moduulitiedostoa $fullModulePath ei löydy. Ohitetaan päivitys."
                    continue
                }
                
                try {
                    # Lue .bas-tiedoston sisältö
                    $moduleContent = Get-Content -Path $fullModulePath -Raw -Encoding UTF8
                    
                    # Poista VBA-tiedoston header-rivit (Attribute VB_Name jne.)
                    # Säilytetään vain varsinainen VBA-koodi
                    $lines = $moduleContent -split "`r?`n"
                    $codeStartIndex = 0
                    for ($i = 0; $i -lt $lines.Count; $i++) {
                        if ($lines[$i] -match "^Attribute\s+" -or $lines[$i] -match "^VERSION\s+") {
                            $codeStartIndex = $i + 1
                        }
                        elseif ($lines[$i].Trim() -eq "") {
                            continue
                        }
                        else {
                            break
                        }
                    }
                    
                    # Ota vain VBA-koodi (ilman header-rivejä)
                    $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n").Trim()
                    
                    # Etsi tai luo moduuli
                    $module = $null
                    try {
                        $module = $vbaProject.VBComponents.Item($name)
                        Write-Host "  - Moduuli $name löytyi, päivitetään sisältö..."
                    }
                    catch {
                        # Jos moduulia ei ole, luo se
                        Write-Host "  - Moduulia $name ei löytynyt, luodaan uusi..."
                        $module = $vbaProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
                        $module.Name = $name
                    }
                    
                    # Tyhjennä vanha koodi ja aseta uusi
                    $codeModule = $module.CodeModule
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    $codeModule.AddFromString($cleanCode)
                    
                    Write-Host "  ✅ Päivitettiin $name ($(($cleanCode -split "`n").Count) riviä koodia)"
                    
                }
                catch {
                    Write-Error "  ❌ VIRHE: Moduulin $name päivitys epäonnistui: $($_.Exception.Message)"
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

        }
        catch {
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