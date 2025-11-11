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
Write-Host "Tarkistus OK: Ajetaan 64-bittisessä PowerShellissä." -ForegroundColor Green


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
        Write-Host "Access-tiedostoa ei löydy tai polku on hakemisto: $databasePath"
        exit 1
    }
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Host "Component files folder does not exist: $componentPath"
        exit 1
    }

    # --- 3. Komponenttien määrittely ---
    # Määrittele kaikki ne moduulien ja luokkamoduulien nimet, jotka poistetaan ja tuodaan.
    # Älä käytä tiedostopäätteitä (.bas/.cls) nimissä.
    $componentNames = @(
        "Module1",
        "General",
        "For ACAD Utility",
        "USysCheck",
        "Form_DBUsers", 
        "Form_Linkkien vaihto",
        "Form_Tee Kuvat"
    ) 

    # Retry-asetukset
    $maxRetries = 3
    $retryDelaySeconds = 1

    Write-Host "--- KÄSITELLÄÄN: $databasePath ---"

    $retryCount = 0
    $isOpened = $false

    # --- 4. Tiedoston valmistelu ja avaus ---
    
    # Poista 'Vain luku' -attribuutti
    try {
        Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
        Write-Host "   ✓ Poistettiin Vain luku -attribuutti."
    }
    catch {
        Write-Warning "   ⚠ Vain luku -attribuutin poisto epäonnistui: $($_.Exception.Message). Jatketaan silti."
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
                Write-Warning "   ⚠ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
                Write-Host "   ✗ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen."
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
                Write-Host "   ✗ KRIITTINEN VIRHE: VBA-projektiin (VBE) ei päästy käsiksi (palautti null)."
                Write-Host "     SYY: Accessin turva-asetukset estävät tämän. Tarkista seuraavat:"
                Write-Host "     1. Access - Asetukset - Luottamuskeskus - Luota VBA-projektin objektimallin käyttöön"
                Write-Host "     2. Access - Asetukset - Luottamuskeskus - Luotetut sijainnit - lisää polut: $databasePath ja $componentPath"
                Write-Host "     3. Tiedoston Ominaisuudet - Salli eli Unblock, jos se on ladattu verkosta"
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
                    Write-Host "   ✗ VIRHE: Komponenttitiedostoa $name.bas tai $name.cls ei löydy polusta $componentPath. Ohitetaan päivitys."
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

                    Write-Host "   ✓ päivitettiin $name - $(($cleanCode -split "`n").Count) riviä koodia"

                }
                catch {
                    Write-Host "   ✗ VIRHE: Komponentin $name päivitys epäonnistui: $($_.Exception.Message)"
                }
            }
            
            # 5.3 Tallenna ja sulje
            $acCmdSaveDatabase = 19
            $database.Application.DoCmd.RunCommand($acCmdSaveDatabase) 
            Write-Host "   ✓ Tietokanta tallennettiin paikoilleen."

            # Laita varoitukset takaisin päälle
            $access.DoCmd.SetWarnings($true)
            
            $access.CloseCurrentDatabase()
            Write-Host "   ✓ Tiedosto $databasePath päivitetty onnistuneesti."

        }
        catch {
            # Tämä 'catch' nappaa VBA-käsittelyn virheet
            Write-Host "✗ Virhe VBA-käsittelyssä tai tallennuksessa: $($_.Exception.Message)"

            # Yritetään siistiä tietokantayhteys
            try {
                if ($null -ne $access) {
                    $access.DoCmd.SetWarnings($true)
                    $access.CloseCurrentDatabase() 
                }
            }
            catch {
                Write-Warning "   - Huomio: Tietokannan sulkeminen virhetilanteessa failed."
            }
        }
    } # end if ($isOpened)

}
catch {
    # Tämä 'catch' nappaa kaikki ylemmän tason virheet (esim. New-Object, polkujen tarkistus)
    Write-Host "--- KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA ---"
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
            Write-Host "   - Access.Quit() called."
        }
        catch {
            Write-Warning "   - Access.Quit() failed (prosessi oli ehkä jo kaatunut)."
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
