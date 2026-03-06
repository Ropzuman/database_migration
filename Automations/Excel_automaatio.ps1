# Excel_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit listojen kysely työkaluissa 64-bittisillä moduuleilla.

# KÄYTTÄYTYMINEN:
# - Varmistaa, että skripti ajetaan 64-bittisessä (x64) PowerShellissä.
# - Kysyy polut työkirjojen hakemistoon ja moduulitiedostojen hakemistoon (ellei oletuksia ole asetettu).
# - Avaa jokaisen .xlsm-työkirjan, päivittää moduulien sisällön suoraan, tallentaa väliaikaisesti ja korvaa alkuperäisen.
# - Käyttää retry-logiikkaa lukkojen kiertämiseksi (OneDrive ym.).
# - Käyttää try...finally-lohkoa varmistaakseen, että Excel-prosessi suljetaan aina.

# TÄRKEÄÄ - VBComponents.Import-ongelma:
# - Tämä skripti KORVAA moduulien sisällön suoraan CodeModule-rajapinnan kautta.
# - EI käytetä VBComponents.Import()-funktiota, koska se lisää näkymättömiä metatietoja moduuleihin.
# - Import() aiheuttaa moduulien toimintahäiriöitä (käyttäytyy eri tavalla kuin manuaalisesti kopioidut).
# - Nykyinen toteutus: lue .bas-tiedosto → poista headerit → kirjoita puhdas koodi AddFromString()-funktiolla.
# - Tämä vastaa manuaalista kopioi-liitä -toimintoa VBA-editorissa.

# - HUOM: Kun muutoksia ajetaan verkkosijaintiin, pitää käyttää verkkosijainnin nimeä \\proense01\projektit\ 
# - esim. "\\proense01\projektit\24PRO260 Vermo Lämmönsiirrinasema\Z\tools\Projektin listojen excel-kyselyt 64bit WORK IN PROGRESS"

# Muokatun tiedoston voi tallentaa muodossa .ps1 haluamaansa sijaintiin ja suorittaa  seuraavasti:
# - Avaa PowerShell Administratorina ja suorita komento Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# - Suorita "polku tiedostoon"\"tiedoston_nimi".ps1

# --- KRIITTINEN TARKISTUS: Bittisyys ---
if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    Write-Error "Sulje tämä (x86) ikkuna ja käynnistä normaali 'Windows PowerShell'."
    # return on ISE-yhteensopiva; exit 1 lopettaisi koko ISE-istunnon
    return
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä." -ForegroundColor Green

# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$excel = $null

try {
    # --- 1. Alustus ---
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [ALUSTUS] Luodaan Excel COM-objekti..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Excel-objekti luotu." -ForegroundColor Green

    # --- 2. Polkujen kysely (UUSI JÄRJESTYS: Ensin moduulit, sitten kohteet) ---
    # --- Polut ja asetukset (Päivitä oletusarvot tähän halutessasi. Varsinkin uusien moduulien sijainti yleensä kiinteä.) ---
    $DefaultModulePath = ''
    $DefaultExcelFilesPath = ''

    Write-Host "`nVAIHE 1: Moduulien lähde" -ForegroundColor Magenta
    $defaultModuleDisplay = if ([string]::IsNullOrWhiteSpace($DefaultModulePath)) { "(ei oletusta asetettu)" } else { $DefaultModulePath }
    Write-Host "Oletuspolku moduuleille: $defaultModuleDisplay" -ForegroundColor Cyan
    $inputModule = Read-Host -Prompt 'Lisää polku moduulitiedostoille (.bas) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputModule)) {
        if ([string]::IsNullOrWhiteSpace($DefaultModulePath)) {
            Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultModulePath skriptin alussa."
            throw "No module path provided"
        }
        $modulePath = $DefaultModulePath
    }
    else { $modulePath = $inputModule }

    if (-not (Test-Path $modulePath -PathType Container)) {
        Write-Error "Moduulikansiota ei löydy: $modulePath"
        throw "Invalid module path"
    }

    Write-Host "`nVAIHE 2: Päivitettävät Excel-tiedostot" -ForegroundColor Magenta
    $defaultExcelDisplay = if ([string]::IsNullOrWhiteSpace($DefaultExcelFilesPath)) { "(ei oletusta asetettu)" } else { $DefaultExcelFilesPath }
    Write-Host "Oletuspolku Excel-tiedostoille: $defaultExcelDisplay" -ForegroundColor Cyan
    $inputExcel = Read-Host -Prompt 'Lisää polku Excel-tiedostoille (.xlsm) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputExcel)) {
        if ([string]::IsNullOrWhiteSpace($DefaultExcelFilesPath)) {
            Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultExcelFilesPath skriptin alussa."
            throw "No Excel files path provided"
        }
        $excelFilesPath = $DefaultExcelFilesPath
    }
    else { $excelFilesPath = $inputExcel }

    if (-not (Test-Path $excelFilesPath -PathType Container)) {
        Write-Error "Excel-tiedostojen kansiota ei löydy: $excelFilesPath"
        throw "Invalid Excel files path"
    }

    # --- 3. Skannaa moduulit automaattisesti ---
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [MODUULIT] Skannataan .bas-tiedostot kansiosta: $modulePath" -ForegroundColor Cyan
    $basFiles = Get-ChildItem -Path $modulePath -Filter "*.bas"  # vain ylätaso, ei alihakemistoja
    
    if ($basFiles.Count -eq 0) {
        Write-Error "Ei löytynyt yhtään .bas-tiedostoa kansiosta: $modulePath"
        throw "No module files found"
    }
    
    # Poimii tiedostonimet ilman .bas-päätettä
    $moduleNames = $basFiles | ForEach-Object { $_.BaseName }
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $($moduleNames.Count) moduulia:" -ForegroundColor Green
    $moduleNames | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray } 

    # Retry-asetukset OneDrive-lukkojen kiertämiseksi
    $maxRetries = 3
    $retryDelaySeconds = 1

    # --- 4. Käsittele kaikki .xlsm tiedostot ---
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [TYÖKIRJAT] Haetaan .xlsm-tiedostot kohteesta: $excelFilesPath" -ForegroundColor Cyan
    $xlsmFiles = Get-ChildItem -Path $excelFilesPath -Filter "*.xlsm"
    $totalFiles = $xlsmFiles.Count
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $totalFiles työkirjaa käsiteltäväksi." -ForegroundColor Green
    
    $currentFileIndex = 0
    $wbSuccess = 0; $wbSkipped = 0; $wbFailed = 0
    $modSuccess = 0; $modFailed = 0

    # Käsittele kaikki .xlsm tiedostot kohdekansiossa
    $xlsmFiles | ForEach-Object {
        $currentFileIndex++
        $workbookPath = $_.FullName
        Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [TYÖKIRJA $currentFileIndex/$totalFiles] KÄSITELLÄÄN: $workbookPath" -ForegroundColor Yellow

        $retryCount = 0
        $isOpened = $false
        $workbook = $null
        
        # 1. Poistaa tiedoston Vain luku -attribuutin (IsReadOnly = $false).
        try {
            Set-ItemProperty -Path $workbookPath -Name IsReadOnly -Value $false -Force
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Poistettiin Vain luku -attribuutti tiedostojärjestelmästä."
        }
        catch {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Vain luku -attribuutin poisto epäonnistui. Jatketaan silti."
        }

        # 2. Avaa työkirja (Retry-logiikalla lukkojen kiertämiseksi)
        do {
            try {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    [AVAUS] Avataan työkirja..."
                # Parametrit: (FileName, UpdateLinks, ReadOnly)
                $workbook = $excel.Workbooks.Open(
                    $workbookPath, 
                    $false,       # 2: UpdateLinks (False, ei null)
                    $false        # 3: ReadOnly (Pakotetaan avaamaan kirjoitusoikeudella)
                )
                $isOpened = $true
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Työkirja avattu onnistuneesti."
            }
            catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen $retryDelaySeconds sekunnin kuluttua (Yritys $retryCount / $maxRetries)."
                    Start-Sleep -Seconds $retryDelaySeconds
                }
                else {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss')    ✗ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen. Jätetään käsittelemättä."
                    $isOpened = $false
                    $wbSkipped++
                    break  # Ei throw — ForEach-Object jatkaa seuraavaan tiedostoon
                }
            }
        } while (-not $isOpened -and $retryCount -lt $maxRetries)
    
        # Käsittely jatkuu vain, jos tiedosto avattiin onnistuneesti
        if ($isOpened) {
            
            try {
                # Tarkista, onko VBA-projekti käytettävissä (riippuu Trust Center -asetuksista)
                $vbaProject = $workbook.VBProject
                
                # KRIITTINEN TARKISTUS: Trust Center -asetukset
                if ($null -eq $vbaProject) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✗ KRIITTINEN VIRHE: VBA-projektiin ei päästy käsiksi (palautti null)." -ForegroundColor Red
                    Write-Host "     SYY: Excelin turva-asetukset estävät tämän. Tarkista Trust Center -asetukset." -ForegroundColor Yellow
                    throw "VBA Project is null. Check Excel Trust Center settings."
                }
                
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ VBA-projekti avattu." -ForegroundColor Green

                # 3. Päivitä moduulien sisältö suoraan (välttää Import-metatietojen ongelman)
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    [MODUULIT] Aloitetaan päivitys..."
                foreach ($name in $moduleNames) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')       [KÄSITTELY] $name" -ForegroundColor Cyan
                    
                    $fullModulePath = Join-Path $modulePath "$($name).bas"
                    
                    if (-not (Test-Path $fullModulePath)) {
                        Write-Error "$(Get-Date -Format 'HH:mm:ss')  ✗ VIRHE: Uutta moduulitiedostoa $fullModulePath ei löydy. Ohitetaan päivitys."
                        $modFailed++ 
                        continue
                    }
                    
                    try {
                        # Lue .bas-tiedoston sisältö StreamReaderilla — käsittelee UTF-8 BOM:n automaattisesti
                        # Get-Content -Encoding UTF8 voi PS 5.1:ssä palauttaa BOM:n merkkijonon ensimmäisenä merkkinä
                        # try-finally takaa Close()-kutsun myös ReadToEnd()-poikkeuksen sattuessa (tiedostokahva ei jää auki)
                        $reader = $null
                        try {
                            $reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
                            $moduleContent = $reader.ReadToEnd()
                        }
                        finally {
                            if ($null -ne $reader) { $reader.Close(); $reader = $null }
                        }
                        # Poistetaan BOM varmuuden vuoksi (U+FEFF), jos StreamReader ei sitä poistanut
                        if ($moduleContent.Length -gt 0 -and [int][char]$moduleContent[0] -eq 0xFEFF) {
                            $moduleContent = $moduleContent.Substring(1)
                        }
                        
                        # PARANNETTU HEADER-PARSAUS:
                        # Poista VBA-tiedoston header-rivit (Attribute VB_Name jne.)
                        # Säilytetään varsinainen VBA-koodi (Option Explicit, Declare, Public, Private, jne.)
                        $lines = $moduleContent -split "`r?`n"
                        $codeStartIndex = 0
                        $inHeader = $true
                        
                        # Käy läpi rivejä ja tunnista header-lohkon loppu
                        for ($i = 0; $i -lt $lines.Count; $i++) {
                            $line = $lines[$i].Trim()
                            
                            if ($inHeader) {
                                # Header-rivit (poistetaan):
                                if ($line -match "^Attribute\s+VB_(Name|GlobalNameSpace|Creatable|PredeclaredId|Exposed)" -or 
                                    $line -match "^VERSION\s+" -or
                                    $line -eq "") {
                                    # Jatka header-lohkossa
                                    $codeStartIndex = $i + 1
                                }
                                else {
                                    # Ensimmäinen varsinainen koodirivi löytyi (Option, Declare, Public, Private, jne.)
                                    $inHeader = $false
                                    break
                                }
                            }
                        }
                        
                        # Ota vain VBA-koodi (ilman header-rivejä)
                        $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n").Trim()
                        
                        if ([string]::IsNullOrWhiteSpace($cleanCode)) {
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ VAROITUS: Tiedosto $name on tyhjä tai sisältää vain headerit. Ohitetaan." -ForegroundColor Yellow
                            continue
                        }
                        
                        # Etsi tai luo moduuli
                        $module = $null
                        try {
                            $module = $vbaProject.VBComponents.Item($name)
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ Moduuli löytyi, päivitetään sisältö..."
                        }
                        catch {
                            # Jos moduulia ei ole, luo se
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ! Moduulia ei löytynyt, luodaan uusi..."
                            $module = $vbaProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
                            $module.Name = $name
                        }
                        
                        # Tyhjennä vanha koodi ja aseta uusi
                        $codeModule = $module.CodeModule
                        $oldLineCount = $codeModule.CountOfLines
                        
                        if ($oldLineCount -gt 0) {
                            $codeModule.DeleteLines(1, $oldLineCount)
                        }
                        $codeModule.AddFromString($cleanCode)
                        
                        $newLineCount = $codeModule.CountOfLines
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ VALMIS: $name ($oldLineCount → $newLineCount riviä)" -ForegroundColor Green
                        $modSuccess++
                        
                    }
                    catch {
                        Write-Error "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Moduulin $name päivitys epäonnistui: $($_.Exception.Message)"
                        $modFailed++
                    }
                }
                
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    [MODUULIT] Kaikki moduulit käsitelty."

                # 4. Tallenna väliaikaiseen tiedostoon, korvaa atomisesti
                $tempSuffix = "_MIGRATED"
                # Käytetään Path-metodeja estääksemme ".xlsm"-korvauksen kansionimiin (HIGH: path corruption)
                $wbDir = [System.IO.Path]::GetDirectoryName($workbookPath)
                $wbStem = [System.IO.Path]::GetFileNameWithoutExtension($workbookPath)
                $tempWorkbookPath = [System.IO.Path]::Combine($wbDir, $wbStem + $tempSuffix + ".xlsm")

                Write-Host "$(Get-Date -Format 'HH:mm:ss')    [TALLENNUS] Tallennetaan väliaikaiseen tiedostoon: $tempWorkbookPath"

                # Tarkistetaan, onko väliaikainen tiedosto jäänyt edellisestä epäonnistuneesta ajosta
                if (Test-Path $tempWorkbookPath) {
                    Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Väliaikainen tiedosto löytyi jäänteenä edellisestä ajosta: $tempWorkbookPath"
                    Write-Warning "    Poistetaan ennen tallennusta..."
                    Remove-Item -Path $tempWorkbookPath -Force -ErrorAction Stop
                }

                # KRIITTINEN KORJAUS: Käytä FileFormat-parametria (52 = xlOpenXMLWorkbookMacroEnabled)
                # Ilman tätä eri Office-versiot voivat tulkita formaatin eri tavalla
                $xlOpenXMLWorkbookMacroEnabled = 52
                $workbook.SaveAs($tempWorkbookPath, $xlOpenXMLWorkbookMacroEnabled)
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Väliaikainen tallennus onnistui."

                # Sulje työkirja ja vapauta COM-viite ennen tiedostojärjestelmäoperaatioita
                $workbook.Close($false)
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
                $workbook = $null
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Työkirja suljettu ja COM-viite vapautettu."

                # ATOMINEN KORVAUS: alkuperäinen → .bak, temp → lopullinen, .bak poistetaan
                # Jos Rename-Item epäonnistuu, .bak palautetaan alkuperäiseksi — tiedosto ei häviä
                $backupPath = $workbookPath + ".bak"
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    [KORVAUS] Siirretään alkuperäinen varmuuskopioksi..."
                Move-Item -Path $workbookPath -Destination $backupPath -Force -ErrorAction Stop
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Alkuperäinen siirretty varmuuskopioksi."

                try {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    [KORVAUS] Nimetään väliaikainen tiedosto lopulliseksi..."
                    Rename-Item -Path $tempWorkbookPath -NewName (Split-Path $workbookPath -Leaf) -Force -ErrorAction Stop
                    # Onnistui — varmuuskopio poistetaan
                    Remove-Item -Path $backupPath -Force -ErrorAction SilentlyContinue
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tiedosto $workbookPath päivitetty onnistuneesti!" -ForegroundColor Green
                    $wbSuccess++
                }
                catch {
                    # Tallennetaan alkuperäinen virheviesti ennen palautusyritystä — muuten palautusvirhe korvaa sen
                    $renameError = $_.Exception.Message
                    Write-Error "$(Get-Date -Format 'HH:mm:ss')    ✗ Uudelleennimeäminen epäonnistui: $renameError"
                    Write-Warning "   Yritetään palauttaa alkuperäinen varmuuskopiosta: $backupPath"
                    try {
                        Move-Item -Path $backupPath -Destination $workbookPath -Force -ErrorAction Stop
                        Write-Host "   ✓ Alkuperäinen palautettu onnistuneesti." -ForegroundColor Green
                    }
                    catch {
                        # Palautuskin epäonnistui — kerrotaan operaattorille tiedostojen tila selkeästi
                        Write-Error "   ✗ KRIITTINEN: Palautus epäonnistui myös: $($_.Exception.Message)"
                        Write-Error "   Tiedostot levyllä:"
                        Write-Error "     Varmuuskopio (alkuperäinen): $backupPath"
                        Write-Error "     Päivitetty (nimeämätön):     $tempWorkbookPath"
                        Write-Error "   Nimeä päivitetty tiedosto manuaalisesti alkuperäiseksi tai palauta varmuuskopio."
                    }
                    # Uudelleenheitä ALKUPERÄINEN nimivirhe, ei palautusvirhe
                    throw [System.Exception]::new("Rename failed: $renameError", $_.Exception)
                }

            }
            catch {
                # Virheenkäsittely
                $wbFailed++
                Write-Error "$(Get-Date -Format 'HH:mm:ss') ✗ VIRHE VBA-käsittelyssä tai tallennuksessa/korvauksessa: $($_.Exception.Message)"
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    Virhetyyppi: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
                
                # Jos $workbook ei ole null, se tarkoittaa että tallennus/sulku ei ehtinyt ajua
                if ($workbook -ne $null) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ⚠ Suljetaan työkirja tallentamatta virhetilanteen vuoksi."
                    try {
                        $workbook.Close($false)
                        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
                        $workbook = $null
                    }
                    catch {
                        Write-Warning "$(Get-Date -Format 'HH:mm:ss')       Työkirjan sulkeminen epäonnistui."
                        $workbook = $null
                    }
                }
            }
        } # end if ($isOpened)
    } # end ForEach-Object
    
    Write-Host ""
    Write-Host "$(Get-Date -Format 'HH:mm:ss') === YHTEENVETO ===" -ForegroundColor Cyan
    Write-Host "  Työkirjat: $wbSuccess onnistui / $wbSkipped ohitettu / $wbFailed epäonnistui" -ForegroundColor $(if ($wbFailed -gt 0 -or $wbSkipped -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "  Moduulit:  $modSuccess onnistui / $modFailed epäonnistui" -ForegroundColor $(if ($modFailed -gt 0) { 'Yellow' } else { 'Green' })

}
catch {
    # Päätason virheenkäsittely
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [FATAL] KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA" -ForegroundColor Red
    Write-Error $_.Exception.Message
    Write-Host "$(Get-Date -Format 'HH:mm:ss')    Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
}
finally {
    # --- PAKOTETTU SIIVOUS ---
    # Tämä lohko suoritetaan AINA, vaikka skripti kaatuisi tai onnistuisi.
    # Tämä estää "zombie" (jumittuneiden) Excel-prosessien syntymisen.

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Siivotaan ja suljetaan Excel-prosessi..." -ForegroundColor Magenta
    
    # Sulje Excel-sovellus
    if ($null -ne $excel) {
        try {
            $excel.Quit()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Excel.Quit() suoritettu."
        }
        catch {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Excel.Quit() epäonnistui (prosessi oli ehkä jo kaatunut)."
        }
        
        Start-Sleep -Milliseconds 500
        
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Excel COM-objekti vapautettu."
        }
        catch { <# Hiljainen #> }
    }
    
    Remove-Variable excel -ErrorAction SilentlyContinue

    # Pakotetaan roskienkeruu COM-viitteiden välittömäksi vapauttamiseksi
    # Ilman tätä Excel.exe voi jäädä prosessilistalle kunnes GC ajaa automaattisesti
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Siivous valmis." -ForegroundColor Green
}

#--------------------------------- Skripti päättyy ----------------------------------