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

# Script parameters: allow non-interactive runs when paths/selections are supplied.
param(
    [string]$ModulePath,
    [string]$ExcelFilesPath,
    [string]$Selection,
    [switch]$AutoRun
)

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

    # If parameters supplied, prefer them over defaults
    if (-not [string]::IsNullOrWhiteSpace($ModulePath)) { $DefaultModulePath = $ModulePath }
    if (-not [string]::IsNullOrWhiteSpace($ExcelFilesPath)) { $DefaultExcelFilesPath = $ExcelFilesPath }

    Write-Host "`nVAIHE 1: Moduulien lähde" -ForegroundColor Magenta
    $defaultModuleDisplay = if ([string]::IsNullOrWhiteSpace($DefaultModulePath)) { "(ei oletusta asetettu)" } else { $DefaultModulePath }
    Write-Host "Oletuspolku moduuleille: $defaultModuleDisplay" -ForegroundColor Cyan
    # If ModulePath parameter/default exists, don't prompt
    if (-not [string]::IsNullOrWhiteSpace($DefaultModulePath)) {
        $modulePath = $DefaultModulePath
        Write-Host "Käytetään moduulipolkua: $modulePath" -ForegroundColor Cyan
    }
    else {
        $inputModule = Read-Host -Prompt 'Lisää polku moduulitiedostoille (.bas) (paina Enter käyttääksesi oletusta)'
        if ([string]::IsNullOrWhiteSpace($inputModule)) {
            if ([string]::IsNullOrWhiteSpace($DefaultModulePath)) {
                Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultModulePath skriptin alussa."
                throw "No module path provided"
            }
            $modulePath = $DefaultModulePath
        }
        else { $modulePath = $inputModule }
    }

    if (-not (Test-Path $modulePath -PathType Container)) {
        Write-Error "Moduulikansiota ei löydy: $modulePath"
        throw "Invalid module path"
    }

    Write-Host "`nVAIHE 2: Päivitettävät Excel-tiedostot" -ForegroundColor Magenta
    $defaultExcelDisplay = if ([string]::IsNullOrWhiteSpace($DefaultExcelFilesPath)) { "(ei oletusta asetettu)" } else { $DefaultExcelFilesPath }
    Write-Host "Oletuspolku Excel-tiedostoille: $defaultExcelDisplay" -ForegroundColor Cyan
    # If ExcelFilesPath parameter/default exists, don't prompt
    if (-not [string]::IsNullOrWhiteSpace($DefaultExcelFilesPath)) {
        $excelFilesPath = $DefaultExcelFilesPath
        Write-Host "Käytetään Excel-kansiota: $excelFilesPath" -ForegroundColor Cyan
    }
    else {
        $inputExcel = Read-Host -Prompt 'Lisää polku Excel-tiedostoille (.xlsm) (paina Enter käyttääksesi oletusta)'
        if ([string]::IsNullOrWhiteSpace($inputExcel)) {
            if ([string]::IsNullOrWhiteSpace($DefaultExcelFilesPath)) {
                Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultExcelFilesPath skriptin alussa."
                throw "No Excel files path provided"
            }
            $excelFilesPath = $DefaultExcelFilesPath
        }
        else { $excelFilesPath = $inputExcel }
    }

    if (-not (Test-Path $excelFilesPath -PathType Container)) {
        Write-Error "Excel-tiedostojen kansiota ei löydy: $excelFilesPath"
        throw "Invalid Excel files path"
    }

    # --- 3. Skannaa moduulit automaattisesti ---
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [MODUULIT] Skannataan .bas- ja .cls-tiedostot kansiosta: $modulePath" -ForegroundColor Cyan
    $basFiles = Get-ChildItem -Path $modulePath -Filter "*.bas"  # vain ylätaso, ei alihakemistoja
    $clsFiles = Get-ChildItem -Path $modulePath -Filter "*.cls"  # luokkamoduulit
    $allModuleFiles = @($basFiles) + @($clsFiles)

    if ($allModuleFiles.Count -eq 0) {
        Write-Error "Ei löytynyt yhtään .bas- tai .cls-tiedostoa kansiosta: $modulePath"
        throw "No module files found"
    }

    # Rakennetaan moduulitieto-objektit (nimi, polku, tiedostopääte)
    $moduleInfos = $allModuleFiles | ForEach-Object {
        [PSCustomObject]@{ Name = $_.BaseName; Path = $_.FullName; Extension = $_.Extension.ToLower() }
    }

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $($moduleInfos.Count) moduulia:" -ForegroundColor Green
    $moduleInfos | ForEach-Object { Write-Host "  - $($_.Name) ($($_.Extension))" -ForegroundColor Gray }

    # Retry-asetukset OneDrive-lukkojen kiertämiseksi
    $maxRetries = 3
    $retryDelaySeconds = 1

    # --- 4. Käsittele kaikki .xlsm tiedostot ---
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [TYÖKIRJAT] Haetaan .xlsm-tiedostot kohteesta: $excelFilesPath" -ForegroundColor Cyan

    # Selection logic: prefer explicit $Selection param; if -AutoRun provided and no selection, process all files
    if (-not [string]::IsNullOrWhiteSpace($Selection)) {
        $selection = $Selection
        Write-Host "Käytetään parametrina annettua valintaa: $selection" -ForegroundColor Cyan
    }
    elseif ($AutoRun) {
        # AutoRun without selection => process all .xlsm files
        $selection = "*.xlsm"
        Write-Host "AutoRun käytössä: käsitellään kaikki .xlsm-tiedostot kansiossa." -ForegroundColor Cyan
    }
    else {
        # Kysytään käyttäjältä suoraan tiedostoa tai wildcardia. Tyhjä vastaus peruuttaa suorittamisen.
        $selection = Read-Host -Prompt 'Anna tiedostonimi tai wildcard (esim. MyFile.xlsm tai *kysely*.xlsm). Tyhjä = peruuta'
        if ([string]::IsNullOrWhiteSpace($selection)) {
            Write-Host "Peruutetaan: ei valittua tiedostoa." -ForegroundColor Yellow
            return
        }
    }

    if ($selection -match '[\\/:]') {
        # Käyttäjä antoi täyden polun
        if (Test-Path $selection) {
            $xlsmFiles = @((Get-Item -Path $selection))
        }
        else {
            Write-Error "Tiedostoa ei löydy: $selection"
            return
        }
    }
    elseif ($selection -match '[*?]') {
        $xlsmFiles = @(Get-ChildItem -Path $excelFilesPath -Filter $selection)
    }
    else {
        $candidate = Join-Path $excelFilesPath $selection
        if (Test-Path $candidate) {
            $xlsmFiles = @((Get-Item -Path $candidate))
        }
        else {
            # Yritetään myös osittaisella nimellä jos suora nimi ei löytynyt
            $xlsmFiles = @(Get-ChildItem -Path $excelFilesPath -Filter "*$selection*")
        }
    }

    if ($xlsmFiles.Count -eq 0) {
        Write-Error "Ei löydetty yhtään tiedostoa hakuehdoilla: $selection"
        return
    }

    $totalFiles = $xlsmFiles.Count
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $totalFiles työkirjaa käsiteltäväksi." -ForegroundColor Green
    Write-Host "HUOM: Moduulien kirjoittaminen vaatii työkirjojen avaamisen VBProject-rajapinnan kautta. Avataan tiedostot yksi kerrallaan, näkyvyyttä ei aseteta." -ForegroundColor Yellow
    
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

        try {
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [AVAUS] Avataan työkirja..."
            $workbook = $excel.Workbooks.Open($workbookPath, $false, $false)
            $isOpened = $true
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Työkirja avattu onnistuneesti."
        }
        catch {
            Write-Error "$(Get-Date -Format 'HH:mm:ss')    ✗ VIRHE: Tiedostoa ei voitu avata: $($_.Exception.Message). Jätetään käsittelemättä."
            $isOpened = $false
            $wbSkipped++
        }
    
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
                foreach ($modInfo in $moduleInfos) {
                    $name = $modInfo.Name
                    $fullModulePath = $modInfo.Path
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')       [KÄSITTELY] $name ($($modInfo.Extension))" -ForegroundColor Cyan
                    
                    if (-not (Test-Path $fullModulePath)) {
                        Write-Error "$(Get-Date -Format 'HH:mm:ss')  ✗ VIRHE: Uutta moduulitiedostoa $fullModulePath ei löydy. Ohitetaan päivitys."
                        $modFailed++ 
                        continue
                    }
                    
                    try {
                        # Lue .bas- tai .cls-tiedoston sisältö StreamReaderilla — käsittelee UTF-8 BOM:n automaattisesti
                        # Get-Content -Encoding UTF8 voi PS 5.1:ssä palauttaa BOM:n merkkijonon ensimmäisenä merkkinä
                        # try-finally takaa Close()-kutsun myös ReadToEnd()-poikkeuksen sattuessa (tiedostokahva ei jää auki)
                        $reader = $null
                        try {
                            # Use New-Object for broader PowerShell compatibility instead of ::new()
                            $reader = New-Object System.IO.StreamReader ($fullModulePath, [System.Text.Encoding]::UTF8, $true)
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
                        # .cls-tiedostoissa poistetaan myös VERSION...CLASS ja BEGIN...END-lohko
                        # Säilytetään varsinainen VBA-koodi (Option Explicit, Declare, Public, Private, jne.)
                        $lines = $moduleContent -split "`r?`n"
                        $codeStartIndex = 0
                        $inHeader = $true
                        $inBeginBlock = $false  # .cls-tiedoston BEGIN...END-lohkon seuranta
                        
                        # Käy läpi rivejä ja tunnista header-lohkon loppu
                        for ($i = 0; $i -lt $lines.Count; $i++) {
                            $line = $lines[$i].Trim()
                            
                            if ($inHeader) {
                                if ($inBeginBlock) {
                                    # Ollaan BEGIN...END-lohkossa — ohitetaan rivit kunnes END löytyy
                                    if ($line -eq "END") { $inBeginBlock = $false }
                                    $codeStartIndex = $i + 1
                                }
                                # Header-rivit (poistetaan):
                                elseif ($line -match "^Attribute\s+VB_(Name|GlobalNameSpace|Creatable|PredeclaredId|Exposed)" -or 
                                    $line -match "^VERSION\s+" -or
                                    $line -eq "") {
                                    # Jatka header-lohkossa
                                    $codeStartIndex = $i + 1
                                }
                                elseif ($line -eq "BEGIN") {
                                    # .cls-tiedoston BEGIN...END-lohko alkaa
                                    $inBeginBlock = $true
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
                        if ($codeStartIndex -lt $lines.Count) {
                            if ($codeStartIndex -eq ($lines.Count - 1)) {
                                $cleanCode = $lines[$codeStartIndex].Trim()
                            }
                            else {
                                $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n").Trim()
                            }
                        }
                        else {
                            $cleanCode = ''
                        }
                        
                        if ([string]::IsNullOrWhiteSpace($cleanCode)) {
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ VAROITUS: Tiedosto $name on tyhjä tai sisältää vain headerit. Ohitetaan." -ForegroundColor Yellow
                            continue
                        }
                        
                        # Etsi tai luo moduuli (tyyppi määräytyy tiedostopäätteen mukaan)
                        $module = $null
                        try {
                            $module = $vbaProject.VBComponents.Item($name)
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ Moduuli löytyi, päivitetään sisältö..."
                        }
                        catch {
                            # Jos moduulia ei ole, luo se oikealla tyypillä
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ! Moduulia ei löytynyt, luodaan uusi..."
                            # 1 = vbext_ct_StdModule (.bas), 2 = vbext_ct_ClassModule (.cls)
                            $moduleType = if ($modInfo.Extension -eq ".cls") { 2 } else { 1 }
                            $module = $vbaProject.VBComponents.Add($moduleType)
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
                    Remove-Item -Path $backupPath -Force -ErrorAction SilentlyContinue
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tiedosto $workbookPath päivitetty onnistuneesti!" -ForegroundColor Green
                    $wbSuccess++
                }
                catch {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss')    ✗ Uudelleennimeäminen epäonnistui: $($_.Exception.Message)"
                    Write-Warning "   Yritetään palauttaa alkuperäinen varmuuskopiosta: $backupPath"
                    Move-Item -Path $backupPath -Destination $workbookPath -Force -ErrorAction SilentlyContinue
                    throw [System.Exception]::new("Rename failed: $($_.Exception.Message)", $_.Exception)
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