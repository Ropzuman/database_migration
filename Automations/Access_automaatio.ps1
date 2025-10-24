# Access_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit (.bas) ja luokkamoduulit (.cls) kannassa.
# YMPÄRISTÖ: Tämä skripti on suunniteltu ajettavaksi 64-bittisessä PowerShellissä, 
#           ja se automatisoi 64-bittistä Microsoft Accessia.

# KÄYTTÄYTYMINEN:
# - Varmistaa, että skripti ajetaan 64-bittisessä (x64) PowerShellissä.
# - Kysyy polun yhteen tietokantatiedostoon ja polun komponenttihakemistoon.
# - Avaa .accdb-kannan, poistaa määritellyt komponentit (moduulit/luokat) ja tuo uudet.
# - Käyttää try...finally-lohkoa varmistaakseen, että Access-prosessi suljetaan aina.

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
        Write-Error "Access-tiedostoa ei löydy (tai polku on hakemisto): $databasePath"
        exit 1
    }
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Error "Component files folder does not exist: $componentPath"
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
                Write-Error "     1. Access > Asetukset > Luottamuskeskus > 'Luota VBA-projektin objektimallin käyttöön'."
                Write-Error "     2. Access > Asetukset > Luottamuskeskus > 'Luotetut sijainnit' (lisää $databasePath ja $componentPath)."
                Write-Error "     3. Tiedoston Ominaisuudet > 'Salli' (Unblock), jos se on ladattu verkosta."
                throw "VBA Project is null. Check Access Trust Center settings."
            }
            
            Write-Host "   - VBA-projekti avattu onnistuneesti."

            # 5.1 Poista vanhat komponentit
            foreach ($name in $componentNames) {
                try {
                    $component = $vbaProject.VBComponents.Item($name)
                    $vbaProject.VBComponents.Remove($component)
                    Write-Host "   ✅ Poistettiin vanha komponentti: $name"
                }
                catch {
                    # Tämä on ok, jos komponenttia ei ollut olemassa
                    Write-Host "   - Info: Komponenttia $name ei löytynyt poistettavaksi (tämä on ok)."
                }
            }

            # 5.2 Tuo uudet komponentit
            foreach ($name in $componentNames) {
                $basPath = Join-Path $componentPath "$($name).bas"
                $clsPath = Join-Path $componentPath "$($name).cls"
                $fullModulePath = $null

                if (Test-Path $basPath) {
                    $fullModulePath = $basPath
                }
                elseif (Test-Path $clsPath) {
                    $fullModulePath = $clsPath
                }

                if (-not $fullModulePath) {
                    Write-Error "   ❌ VIRHE: Uutta komponenttitiedostoa ($name.bas/.cls) ei löydy polusta $componentPath. Ohitetaan tuonti."
                    continue
                }
                
                try {
                    $vbaProject.VBComponents.Import($fullModulePath)
                    Write-Host "   ✅ Tuotiin uusi komponentti: $name"
                }
                catch {
                    Write-Error "   ❌ VIRHE: Komponentin $name tuonti epäonnistui: $($_.Exception.Message)"
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