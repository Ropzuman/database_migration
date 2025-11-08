# Access_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit (.bas) ja luokkamoduulit (.cls) kannassa.
# YMPÄRISTÖ: Tämä skripti tukee sekä 32-bittistä että 64-bittistä Microsoft Accessia.
#           Skripti tunnistaa automaattisesti Accessin arkkitehtuurin ja käynnistyy
#           uudelleen oikeassa PowerShell-ympäristössä tarvittaessa.

# KÄYTTÄYTYMINEN:
# - Tunnistaa Accessin bittisyyden (32-bit tai 64-bit)
# - Jos PowerShell-bittisyys ei vastaa Accessia, käynnistyy automaattisesti uudelleen
# - Kysyy polun yhteen tietokantatiedostoon ja polun komponenttihakemistoon
# - Avaa .accdb-kannan, päivittää komponenttien sisällön suoraan (moduulit/luokat)
# - Käyttää try...finally-lohkoa varmistaakseen, että Access-prosessi suljetaan aina

# TÄRKEÄÄ - VBComponents.Import-ongelma:
# - Tämä skripti KORVAA komponenttien sisällön suoraan CodeModule-rajapinnan kautta.
# - EI käytetä VBComponents.Import()-funktiota, koska se lisää näkymättömiä metatietoja.
# - Import() aiheuttaa komponenttien toimintahäiriöitä (käyttäytyy eri tavalla kuin manuaalisesti kopioidut).
# - Nykyinen toteutus: lue .bas/.cls-tiedosto → poista headerit → kirjoita puhdas koodi AddFromString()-funktiolla.
# - Tämä vastaa manuaalista kopioi-liitä -toimintoa VBA-editorissa.

# --- KRIITTINEN TARKISTUS: Bittisyys ---
# Tunnista Accessin bittisyys rekisteristä
$accessPath = $null
$accessIs32Bit = $false

# Tarkista ensin 64-bit sijainti
$accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
if ($accessPath -and ($accessPath -match 'x86' -or $accessPath -match 'Program Files \(x86\)')) {
    $accessIs32Bit = $true
}

# Jos ei löytynyt, tarkista 32-bit sijainti (WOW6432Node)
if (-not $accessPath) {
    $accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
    if ($accessPath) {
        $accessIs32Bit = $true
    }
}

if (-not $accessPath) {
    Write-Error "VIRHE: Microsoft Accessia ei löydy järjestelmästä."
    Write-Error "Asenna Microsoft Access ennen tämän skriptin ajamista."
    Start-Sleep -Seconds 10
    exit 1
}

# Tarkista PowerShellin bittisyys
$psIs64Bit = [System.IntPtr]::Size -eq 8

# Näytä tunnistetut arkkitehtuurit
Write-Host "=== BITTISYYS-TARKISTUS ===" -ForegroundColor Cyan
Write-Host "Access: $(if($accessIs32Bit){'32-bit'}else{'64-bit'}) ($accessPath)"
Write-Host "PowerShell: $(if($psIs64Bit){'64-bit'}else{'32-bit'})"

# Jos bittisyydet eivät täsmää, näytä virhe ja lopeta
if ($accessIs32Bit -and $psIs64Bit) {
    Write-Host ""
    Write-Host "❌ VIRHE: Bittisyydet eivät täsmää!" -ForegroundColor Red
    Write-Host "   Access on 32-bittinen, mutta PowerShell on 64-bittinen." -ForegroundColor Yellow
    Write-Host "   64-bit PowerShell ei voi luoda COM-yhteyttä 32-bit Accessiin." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "RATKAISU: Aja skripti batch-tiedostolla:" -ForegroundColor Cyan
    Write-Host "   Kaksoisklikkaa: RUN_ACCESS_AUTOMATION.bat" -ForegroundColor Green
    Write-Host ""
    Write-Host "TAI käynnistä manuaalisesti 32-bit PowerShellissä:" -ForegroundColor Cyan
    Write-Host "   1. Avaa Käynnistä-valikko" -ForegroundColor White
    Write-Host "   2. Etsi 'PowerShell (x86)'" -ForegroundColor White
    Write-Host "   3. Suorita siellä:" -ForegroundColor White
    Write-Host "      cd c:\database_migration\Automations" -ForegroundColor Gray
    Write-Host "      .\Access_automaatio.ps1" -ForegroundColor Gray
    Write-Host ""
    Start-Sleep -Seconds 15
    exit 1
}
elseif (-not $accessIs32Bit -and -not $psIs64Bit) {
    Write-Host ""
    Write-Host "❌ VIRHE: Bittisyydet eivät täsmää!" -ForegroundColor Red
    Write-Host "   Access on 64-bittinen, mutta PowerShell on 32-bittinen." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "RATKAISU: Käynnistä 64-bit PowerShellissä:" -ForegroundColor Cyan
    Write-Host "   1. Avaa Käynnistä-valikko" -ForegroundColor White
    Write-Host "   2. Etsi 'PowerShell' (ei x86)" -ForegroundColor White
    Write-Host "   3. Suorita siellä:" -ForegroundColor White
    Write-Host "      cd c:\database_migration\Automations" -ForegroundColor Gray
    Write-Host "      .\Access_automaatio.ps1" -ForegroundColor Gray
    Write-Host ""
    Start-Sleep -Seconds 15
    exit 1
}
else {
    Write-Host "✅ Bittisyydet täsmäävät - jatketaan..." -ForegroundColor Green
}

Write-Host ""


# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$access = $null
$database = $null

try {
    # --- 1. Alustus ---
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false

    # --- 2. Polkujen kysely ---
    $DefaultAccessFilePath = ''
    $DefaultComponentPath = ''

    Write-Host "Access file path: $DefaultAccessFilePath" -ForegroundColor Cyan
    $inputAccess = Read-Host -Prompt 'Lisää polku Access-tiedostoon (.accdb) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) { $databasePath = $DefaultAccessFilePath } else { $databasePath = $inputAccess }

    Write-Host "Component files folder: $DefaultComponentPath" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Lisää polku komponenttitiedostoille (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) { $componentPath = $DefaultComponentPath } else { $componentPath = $inputComponent }

    # Polkujen tarkistus
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Error "Access-tiedostoa ei löydy tai polku on hakemisto: $databasePath"
        exit 1
    }
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Error "Component files folder does not exist: $componentPath"
        exit 1
    }

    # --- 3. Komponenttien määrittely ---
    # Määrittele kaikki ne moduulien ja luokkamoduulien nimet, jotka poistetaan ja tuodaan.
    # Älä käytä tiedostopäätteitä (.bas/.cls) nimissä.
    # 
    # PÄIVITETTY 2025-11-08: Phase 1 cleanup (dead code removal, Replace() optimization)
    # - Poistettu: Form_USysRevText_OLD (dead code, korvattu Form_USysRevText:llä)
    # - Muokattu: GlobalVBAs (custom Replace() poistettu, Pituus-muuttujat siivottu)
    # - Muokattu: ForDocuments (VBA71.dll API-deklaraatiot poistettu)
    $componentNames = @(
        # Standard modules (.vba)
        "GlobalVBAs",
        "ForDocuments",
        
        # Form modules (.cls)
        "Form_DBUsers",
        "Form_DISTRIBUTION",
        "Form_DOCUMENTS",
        "Form_SETTINGS",
        "Form_USysAddDocument",
        "Form_USysAddedDistr",
        "Form_USysAddToDistr",
        "Form_USysDISTRIB",
        "Form_USysDocs",
        "Form_USysEditDistribution",
        "Form_USysExcelReport",
        "Form_USysNewDistribution",
        "Form_USysNewRecipient",
        "Form_USysOpenFile",
        "Form_USysRecipientsFrm",
        "Form_USysReserve",
        "Form_USysRevText",
        "Form_USysShowCommon",
        "Form_USysStart",
        
        # Report modules (.cls)
        "Report_Copy of TRANSMITTAL",
        "Report_TRANSMITTAL Copy",
        "Report_TRANSMITTAL",
        "Report_USYSTRANSMITTALFP"
    ) 

    # Retry-asetukset
    $maxRetries = 3
    $retryDelaySeconds = 1

    Write-Host "--- KÄSITTELLÄÄN: $databasePath ---"

    $retryCount = 0
    $isOpened = $false

    # --- 4. Tiedoston valmistelu ja avaus ---
    
    # Poista 'Vain luku' -attribuutti
    try {
        Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
        Write-Host "   ✅ Poistettiin Vain luku -attribuutti."
    }
    catch {
        Write-Warning "   ⚠️ Vain luku -attribuutin poisto epäonnistui: $($_.Exception.Message). Jatketaan silti."
    }

    # Avaa tietokanta retry-logiikalla
    do {
        try {
            $access.OpenCurrentDatabase($databasePath) 
            $database = $access.CurrentDb()
            $isOpened = $true
            Write-Host "   - Tietokanta avattu onnistuneesti."
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "   ⚠️ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
                Write-Error "   ❌ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen."
                throw $_ # Heitetään virhe pää-try-lohkolle
            }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)

    # --- 5. VBA-komponenttien käsittely ---
    if ($isOpened) {
        
        try {
            # Kytke Accessin sisäiset varoitukset pois päältä
            $access.DoCmd.SetWarnings($false)
            Write-Host "   - Access-varoitukset poistettu käytöstä."
            
            # Yritä saada yhteys VBA-projektiin
            $vbaProject = $database.Application.VBE.ActiveVBProject
            
            # KRIITTINEN TARKISTUS: Jos $vbaProject on $null, Accessin turva-asetukset estävät toiminnon.
            if ($null -eq $vbaProject) {
                Write-Error "   ❌ KRIITTINEN VIRHE: VBA-projektiin (VBE) ei päästy käsiksi (palautti null)."
                Write-Error "     SYY: Accessin turva-asetukset estävät tämän. Tarkista seuraavat:"
                Write-Error "     1. Access - Asetukset - Luottamuskeskus - Luota VBA-projektin objektimallin käyttöön"
                Write-Error "     2. Access - Asetukset - Luottamuskeskus - Luotetut sijainnit - lisää polut: $databasePath ja $componentPath"
                Write-Error "     3. Tiedoston Ominaisuudet - Salli eli Unblock, jos se on ladattu verkosta"
                throw "VBA Project is null. Check Access Trust Center settings."
            }
            
            Write-Host "   - VBA-projekti avattu onnistuneesti."

            # 5.1 Päivitä komponenttien sisältö suoraan (välttää Import-metatietojen ongelman)
            foreach ($name in $componentNames) {
                $basPath = Join-Path $componentPath "$($name).bas"
                $clsPath = Join-Path $componentPath "$($name).cls"
                $fullModulePath = $null

                if (Test-Path $basPath) {
                    $fullModulePath = $basPath
                    $componentType = 1  # vbext_ct_StdModule
                }
                elseif (Test-Path $clsPath) {
                    $fullModulePath = $clsPath
                    $componentType = 2  # vbext_ct_ClassModule
                }

                if (-not $fullModulePath) {
                    Write-Error "   ❌ VIRHE: Komponenttitiedostoa $name.bas tai $name.cls ei löydy polusta $componentPath. Ohitetaan päivitys."
                    continue
                }
                
                try {
                    # Lue .bas/.cls-tiedoston sisältö
                    $moduleContent = Get-Content -Path $fullModulePath -Raw -Encoding UTF8
                    
                    # Poista VBA-tiedoston header-rivit (.cls: VERSION, BEGIN/END, Attribute; .bas: Attribute)
                    # Säilytetään vain varsinainen VBA-koodi (Option/Declare/Function/Sub/Dim jne.)
                    $lines = $moduleContent -split "`r?`n"
                    $codeStartIndex = 0
                    
                    # Käy läpi rivejä ja ohita kaikki header-rivit
                    for ($i = 0; $i -lt $lines.Count; $i++) {
                        $line = $lines[$i].Trim()
                        
                        # Ohita VERSION, BEGIN, END, Attribute, MultiUse ja tyhjät rivit
                        if ($line -match "^VERSION\s+" -or 
                            $line -match "^BEGIN$" -or 
                            $line -match "^END$" -or 
                            $line -match "^Attribute\s+" -or
                            $line -match "^MultiUse\s*=" -or
                            $line -eq "") {
                            $codeStartIndex = $i + 1
                        }
                        else {
                            # Kun törmätään ensimmäiseen varsinaiseen koodiriviin, lopeta
                            break
                        }
                    }
                    
                    # Ota vain VBA-koodi (ilman header-rivejä)
                    $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n").Trim()
                    
                    # Etsi tai luo komponentti
                    $component = $null
                    try {
                        $component = $vbaProject.VBComponents.Item($name)
                        Write-Host "   - Komponentti $name löytyi, päivitetään sisältö..."
                    }
                    catch {
                        # Jos komponenttia ei ole, luo se
                        Write-Host "   - Komponenttia $name ei löytynyt, luodaan uusi..."
                        $component = $vbaProject.VBComponents.Add($componentType)
                        $component.Name = $name
                    }
                    
                    # Tyhjennä vanha koodi ja aseta uusi
                    $codeModule = $component.CodeModule
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    $codeModule.AddFromString($cleanCode)
                    
                    Write-Host "   ✅ Paivitettiin $name - $(($cleanCode -split "`n").Count) rivia koodia"
                    
                }
                catch {
                    Write-Error "   ❌ VIRHE: Komponentin $name päivitys epäonnistui: $($_.Exception.Message)"
                }
            }
            
            # 5.3 Tallenna ja sulje
            $acCmdSaveDatabase = 19
            $database.Application.DoCmd.RunCommand($acCmdSaveDatabase) 
            Write-Host "   ✅ Tietokanta tallennettiin paikoilleen."
            
            # Laita varoitukset takaisin päälle
            $access.DoCmd.SetWarnings($true)
            
            $access.CloseCurrentDatabase()
            Write-Host "   ✅ Tiedosto $databasePath päivitetty onnistuneesti."

        }
        catch {
            # Tämä 'catch' nappaa VBA-käsittelyn virheet
            Write-Error "❌ Virhe VBA-käsittelyssä tai tallennuksessa: $($_.Exception.Message)"
            
            # Yritetään siistiä tietokantayhteys
            try {
                if ($null -ne $access) {
                    $access.DoCmd.SetWarnings($true)
                    $access.CloseCurrentDatabase() 
                }
            }
            catch {
                Write-Warning "   - Huomio: Tietokannan sulkeminen virhetilanteessa epäonnistui."
            }
        }
    } # end if ($isOpened)

}
catch {
    # Tämä 'catch' nappaa kaikki ylemmän tason virheet (esim. New-Object, polkujen tarkistus)
    Write-Error "--- KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA ---"
    Write-Error $_.Exception.Message
    
}
finally {
    # --- 6. PAKOTETTU SIIVOUS ---
    # Tämä lohko suoritetaan AINA, vaikka skripti kaatuisi tai onnistuisi.
    # Tämä estää "zombie" (jumittuneiden) Access-prosessien syntymisen.
    
    Write-Host "--- Siivotaan ja suljetaan Access-prosessi... ---"
    
    if ($null -ne $access) {
        try {
            $access.Quit()
            Write-Host "   - Access.Quit() kutsuttu."
        }
        catch {
            Write-Warning "   - Access.Quit() epäonnistui (prosessi oli ehkä jo kaatunut)."
        }
        
        Start-Sleep -Seconds 1 
        
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
            Write-Host "   - COM-objekti vapautettu."
        }
        catch {
            # Hiljainen virhe, jos COM-objektia ei koskaan luotu kunnolla
        }
    }
    
    Remove-Variable access -ErrorAction SilentlyContinue
    Remove-Variable database -ErrorAction SilentlyContinue
    
    Write-Host "--- Siivous valmis. ---"
}