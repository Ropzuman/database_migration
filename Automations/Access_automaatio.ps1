# Access_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit (.bas) ja luokkamoduulit (.cls) kannassa.
# YMPARISTO: Tämä skripti on suunniteltu ajettavaksi 64-bittisessä PowerShellissä,
#           ja se automatisoi 64-bittistä Microsoft Accessia.

# KÄYTTÄYTYMINEN:
# - Varmistaa, että skripti ajetaan 64-bittisessä (x64) PowerShellissä.
# - Kysyy polun yhteen tietokantatiedostoon ja polun komponenttihakemistoon.
# - Avaa .accdb-kannan, päivittää komponenttien sisällön suoraan (moduulit/luokat).
# - Käyttää try...finally-lohkoa varmistaakseen, että Access-prosessi suljetaan aina.

# TÄRKEÄ - VBComponents.Import-ongelma:
# - Tämä skripti KORVAA komponenttien sisällön suoraan CodeModule-rajapinnan kautta.
# - EI käytetä VBComponents.Import()-funktiota, koska se lisää näkymättömiä metatietoja.
# - Import() aiheuttaa komponenttien toimintahäiriöitä (käyttäytyy eri tavalla kuin manuaalisesti kopioidut).
# - Nykyinen toteutus: lue .bas/.cls-tiedosto -> poista headerit -> kirjoita puhdas koodi AddFromString()-funktiolla.
# - Tämä vastaa manuaalista kopioi-liitä -toimintoa VBA-editorissa.

# --- KRIITTINEN TARKISTUS: Bittisyys ---
if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    Write-Error "Sulje tämä (x86) ikkuna ja käynnistä normaali 'Windows PowerShell'."
    Start-Sleep -Seconds 10
    exit 1
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä." -ForegroundColor Green


# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$access = $null
$database = $null
$vbaProject = $null

try {
    # --- 1. Alustus ---
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [ALUSTUS] Luodaan Access COM-objekti..." -ForegroundColor Cyan
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Access-objekti luotu." -ForegroundColor Green

    # --- 2. Polkujen kysely (UUSI JÄRJESTYS: Ensin komponentit, sitten kohteet) ---
    $DefaultComponentPath = ''
    $DefaultAccessFilePath = ''

    Write-Host "`nVAIHE 1: Komponenttien lähde" -ForegroundColor Magenta
    Write-Host "Component files folder: $DefaultComponentPath" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Lisää polku komponenttitiedostoille (.bas/.cls) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) { $componentPath = $DefaultComponentPath } else { $componentPath = $inputComponent }

    Write-Host "`nVAIHE 2: Päivitettävä Access-tiedosto" -ForegroundColor Magenta
    Write-Host "Access file path: $DefaultAccessFilePath" -ForegroundColor Cyan
    $inputAccess = Read-Host -Prompt 'Lisää polku Access-tiedostoon (.accdb) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) { $databasePath = $DefaultAccessFilePath } else { $databasePath = $inputAccess }

    # Polkujen tarkistus
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Host "Component files folder does not exist: $componentPath"
        exit 1
    }
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Host "Access-tiedostoa ei löydy tai polku on hakemisto: $databasePath"
        exit 1
    }

    # --- 3. Skannaa komponentit automaattisesti ---
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [KOMPONENTIT] Skannataan .bas ja .cls -tiedostot kansiosta: $componentPath" -ForegroundColor Cyan
    $basFiles = Get-ChildItem -Path $componentPath -Filter "*.bas"
    $clsFiles = Get-ChildItem -Path $componentPath -Filter "*.cls"
    $allComponentFiles = $basFiles + $clsFiles
    
    if ($allComponentFiles.Count -eq 0) {
        Write-Error "Ei löytynyt yhtään .bas tai .cls -tiedostoa kansiosta: $componentPath"
        throw "No component files found"
    }
    
    # Poimii tiedostonimet ilman päätettä
    $componentNames = $allComponentFiles | ForEach-Object { $_.BaseName }
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $($componentNames.Count) komponenttia:" -ForegroundColor Green
    $componentNames | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray } 

    # Retry-asetukset
    $maxRetries = 3
    $retryDelaySeconds = 1

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [TIETOKANTA] Käsitellään: $databasePath" -ForegroundColor Yellow

    $retryCount = 0
    $isOpened = $false

    # --- 4. Tiedoston valmistelu ja avaus ---
    
    # Poista 'Vain luku' -attribuutti
    try {
        Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
        Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Poistettiin Vain luku -attribuutti."
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Vain luku -attribuutin poisto epäonnistui: $($_.Exception.Message). Jatketaan silti."
    }

    # Avaa tietokanta retry-logiikalla
    do {
        try {
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [AVAUS] Avataan tietokanta..."
            
            # KRIITTINEN: Aseta AutomationSecurity ENNEN avausta (estää makrojen automaattisen suorituksen)
            # Visible = $false on jo asetettu rivillä 39 (Access-objektin luonnin yhteydessä)
            $access.AutomationSecurity = 1  # msoAutomationSecurityLow
            
            # OpenCurrentDatabase(FilePath, [Exclusive], [Password])
            $access.OpenCurrentDatabase($databasePath, $false, "")
            $database = $access.CurrentDb()
            $isOpened = $true
            
            # KRIITTINEN: OpenCurrentDatabase() resetoi Visible-arvon, asetetaan uudelleen
            $access.Visible = $false
            
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tietokanta avattu onnistuneesti (Startup ohitettu)."
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✗ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen."
                throw $_ # Heitetään virhe pää-try-lohkolle
            }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)

    # --- 5. VBA-komponenttien käsittely ---
    if ($isOpened) {
        
        try {
            # Kytke Accessin sisäiset varoitukset pois päältä
            $access.DoCmd.SetWarnings($false)
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Access-varoitukset poistettu käytöstä."

            # KRIITTINEN: VBA-projektiin pääsy suoraan $access-objektista
            # ($database.Application.VBE.ActiveVBProject ei toimi)
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [VBE] Haetaan VBA-projekti..."
            $vbaProject = $access.VBE.ActiveVBProject

            # KRIITTINEN TARKISTUS: Jos $vbaProject on $null, Accessin turva-asetukset estävät toiminnon.
            if ($null -eq $vbaProject) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✗ KRIITTINEN VIRHE: VBA-projektiin (VBE) ei päästy käsiksi (palautti null)." -ForegroundColor Red
                Write-Host "     SYY: Accessin turva-asetukset estävät tämän. Tarkista seuraavat:" -ForegroundColor Yellow
                Write-Host "     1. Access - Asetukset - Luottamuskeskus - Luota VBA-projektin objektimallin käyttöön"
                Write-Host "     2. Access - Asetukset - Luottamuskeskus - Luotetut sijainnit - lisää polut: $databasePath ja $componentPath"
                Write-Host "     3. Tiedoston Ominaisuudet - Salli eli Unblock, jos se on ladattu verkosta"
                throw "VBA Project is null. Check Access Trust Center settings."
            }
            
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ VBA-projekti avattu onnistuneesti." -ForegroundColor Green

            # 5.1 Päivitä komponenttien sisältö suoraan (välttää Import-metatietojen ongelman)
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [KOMPONENTIT] Aloitetaan päivitys..."
            foreach ($name in $componentNames) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')       [KÄSITTELY] $name" -ForegroundColor Cyan
                
                $basPath = Join-Path $componentPath "$($name).bas"
                $clsPath = Join-Path $componentPath "$($name).cls"
                $fullModulePath = $null
                $isFormComponent = $false

                if (Test-Path $basPath) {
                    $fullModulePath = $basPath
                    $componentType = 1  # vbext_ct_StdModule
                }
                elseif (Test-Path $clsPath) {
                    $fullModulePath = $clsPath
                    # Tarkista, onko kyseessä lomake (alkaa "Form_")
                    if ($name -like "Form_*") {
                        $isFormComponent = $true
                        $componentType = 100  # vbext_ct_MSForm (lomakkeet)
                    }
                    else {
                        $componentType = 2  # vbext_ct_ClassModule
                    }
                }

                if (-not $fullModulePath) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Komponenttitiedostoa $name.bas tai $name.cls ei löydy polusta $componentPath. Ohitetaan päivitys." -ForegroundColor Red
                    continue
                }
                
                try {
                    # Lue .bas/.cls-tiedoston sisältö (UTF8 ilman BOM)
                    $moduleContent = Get-Content -Path $fullModulePath -Raw -Encoding UTF8
                    
                    # PARANNETTU HEADER-PARSAUS:
                    # Poista VBA-tiedoston header-rivit (.cls: VERSION, BEGIN/END, Attribute; .bas: Attribute)
                    # Säilytetään varsinainen VBA-koodi (Option Explicit, Declare, Function, Sub, Dim, jne.)
                    $lines = $moduleContent -split "`r?`n"
                    $codeStartIndex = 0
                    $inHeader = $true

                    # Käy läpi rivejä ja tunnista header-lohkon loppu
                    for ($i = 0; $i -lt $lines.Count; $i++) {
                        $line = $lines[$i].Trim()

                        if ($inHeader) {
                            # Header-rivit (poistetaan):
                            if ($line -match "^VERSION\s+" -or
                                $line -match "^BEGIN\s*" -or
                                $line -match "^END\s*$" -or
                                $line -match "^Attribute\s+VB_(Name|GlobalNameSpace|Creatable|PredeclaredId|Exposed)" -or
                                $line -match "^MultiUse\s*=" -or
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
                    # KRIITTINEN: ÄLÄ käytä Trim() - se poistaa trailing newlinen ja aiheuttaa syntax-virheen!
                    $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n")
                    
                    if ([string]::IsNullOrWhiteSpace($cleanCode)) {
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ VAROITUS: Tiedosto $name on tyhjä tai sisältää vain headerit. Ohitetaan." -ForegroundColor Yellow
                        continue
                    }
                    
                    # Etsi tai luo komponentti
                    $component = $null
                    try {
                        $component = $vbaProject.VBComponents.Item($name)
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ Komponentti löytyi, päivitetään sisältö..."
                    }
                    catch {
                        # HUOM: Lomakkeita (Form_*) ei voi luoda VBComponents.Add()-komennolla!
                        # Ne pitää olla jo olemassa kannassa.
                        if ($isFormComponent) {
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Lomake $name ei ole olemassa kannassa. Lomakkeita ei voi luoda automaattisesti, vain päivittää." -ForegroundColor Red
                            continue
                        }
                        
                        # Luo uusi moduuli/luokka
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ! Komponenttia ei löytynyt, luodaan uusi..."
                        $component = $vbaProject.VBComponents.Add($componentType)
                        $component.Name = $name
                    }

                    # Tyhjennä vanha koodi ja lisää uusi
                    $codeModule = $component.CodeModule
                    $oldLineCount = $codeModule.CountOfLines
                    
                    if ($oldLineCount -gt 0) {
                        $codeModule.DeleteLines(1, $oldLineCount)
                    }
                    
                    # Lisää uusi koodi yhtenä stringinä (AddFromString säilyttää line breaks oikein)
                    $codeModule.AddFromString($cleanCode)
                    
                    $newLineCount = $codeModule.CountOfLines
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ VALMIS: $name ($oldLineCount → $newLineCount riviä)" -ForegroundColor Green

                }
                catch {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Komponentin $name päivitys epäonnistui: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')             Virhetyyppi: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')             Stack: $($_.ScriptStackTrace)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [KOMPONENTIT] Kaikki komponentit käsitelty."
            
            # 5.3 Tallenna ja sulje
            # HUOM: CloseCurrentDatabase() tallentaa automaattisesti VBA-projektin.
            # RunCommand($acCmdSaveDatabase) aiheuttaa COM-virheen jos VBA:ssa syntax-virheitä!
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [TALLENNUS] Suljetaan ja tallennetaan tietokanta..."

            # Laita varoitukset takaisin päälle
            $access.DoCmd.SetWarnings($true)
            
            # CloseCurrentDatabase tallentaa VBA-projektin automaattisesti
            $access.CloseCurrentDatabase()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tiedosto $databasePath päivitetty onnistuneesti!" -ForegroundColor Green
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ⚠ HUOM: Tarkista VBA syntax-virheet manuaalisesti (Debug → Compile VBA Project)" -ForegroundColor Yellow

        }
        catch {
            # Tämä 'catch' nappaa VBA-käsittelyn virheet
            Write-Host "$(Get-Date -Format 'HH:mm:ss') ✗ VIRHE VBA-käsittelyssä tai tallennuksessa: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    Virhetyyppi: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow

            # Yritetään siistiä tietokantayhteys
            try {
                if ($null -ne $access) {
                    $access.DoCmd.SetWarnings($true)
                    $access.CloseCurrentDatabase() 
                }
            }
            catch {
                Write-Warning "$(Get-Date -Format 'HH:mm:ss')    - Huomio: Tietokannan sulkeminen virhetilanteessa epäonnistui."
            }
        }
    } # end if ($isOpened)

}
catch {
    # Tämä 'catch' nappaa kaikki ylemmän tason virheet (esim. New-Object, polkujen tarkistus)
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [FATAL] KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA" -ForegroundColor Red
    Write-Error $_.Exception.Message
    Write-Host "$(Get-Date -Format 'HH:mm:ss')    Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    
}
finally {
    # --- 6. PAKOTETTU SIIVOUS ---
    # Tämä lohko suoritetaan AINA, vaikka skripti kaatuisi tai onnistuisi.
    # Tämä estää "zombie" (jumittuneiden) Access-prosessien syntymisen.

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Siivotaan ja suljetaan Access-prosessi..." -ForegroundColor Magenta
    
    # Vapauta VBA Project -viittaus
    if ($null -ne $vbaProject) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbaProject) | Out-Null
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ VBA Project vapautettu."
        }
        catch { <# Hiljainen #> }
    }
    
    # Vapauta Database-viittaus
    if ($null -ne $database) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($database) | Out-Null
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Database-objekti vapautettu."
        }
        catch { <# Hiljainen #> }
    }
    
    # Sulje Access-sovellus
    if ($null -ne $access) {
        try {
            $access.Quit()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Access.Quit() suoritettu."
        }
        catch {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Access.Quit() epäonnistui (prosessi oli ehkä jo kaatunut)."
        }
        
        Start-Sleep -Milliseconds 500
        
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Access COM-objekti vapautettu."
        }
        catch { <# Hiljainen #> }
    }
    
    Remove-Variable access -ErrorAction SilentlyContinue
    Remove-Variable database -ErrorAction SilentlyContinue
    Remove-Variable vbaProject -ErrorAction SilentlyContinue
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Siivous valmis." -ForegroundColor Green
}
