# Access_automaatio.ps1

# --- KRIITTINEN TARKISTUS: Bittisyys ---
# Kohde-Office on 32-bittinen, joten tämän skriptin on AJAUDUTTAVA 32-bittisessä (x86) PowerShellissä.
if ([System.IntPtr]::Size -ne 4) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 32-bittisessä (x86) PowerShellissä, koska kohde-Office on 32-bittinen."
    Write-Error "Sulje tämä (64-bittinen) ikkuna ja käynnistä 'Windows PowerShell (x86)'."
    Start-Sleep -Seconds 10
    exit 1
}
Write-Host "Tarkistus OK: Ajetaan 32-bittisessä (x86) PowerShellissä." -ForegroundColor Green


# Määritellään $access-muuttuja ennalta nulliksi, jotta 'finally'-lohko toimii
$access = $null
$database = $null

try {
    # Koko automaatioprosessi ajetaan try-lohkossa
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false

    # --- Polut ja asetukset ---
    $DefaultAccessFilePath = ''
    $DefaultComponentPath = ''

    Write-Host "Access file path: $DefaultAccessFilePath" -ForegroundColor Cyan
    $inputAccess = Read-Host -Prompt 'Lisää polku Access-tiedostoon (.accdb) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) { $databasePath = $DefaultAccessFilePath } else { $databasePath = $inputAccess }

    Write-Host "Component files folder: $DefaultComponentPath" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Lisää polku komponenttitiedostoille (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) { $componentPath = $DefaultComponentPath } else { $componentPath = $inputComponent }

    # Validate paths
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Error "Access-tiedostoa ei löydy (tai polku on hakemisto): $databasePath"
        exit 1
    }
    if ($databasePath -notlike "*.accdb") {
        Write-Error "Tiedoston pitää olla .accdb-päätteinen: $databasePath"
        exit 1
    }
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Error "Component files folder does not exist: $componentPath"
        exit 1
    }

    # --- Komponenttien nimet ---
    # HUOM! TÄMÄ SKRIPTI EI OSALLA TUODA LOMAKKEITA (Form_) TAI RAPORTTEJA.
    # Se osaa tuoda vain standardimoduuleja (.bas) ja luokkamoduuleja (.cls).
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

    # 1. Poista Vain luku -attribuutti
    try {
        Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
        Write-Host "   ✅ Poistettiin Vain luku -attribuutti tiedostojärjestelmästä."
    }
    catch {
        Write-Warning "   ⚠️ Vain luku -attribuutin poisto epäonnistui: $($_.Exception.Message). Jatketaan yrittämistä."
    }

    # 2. Avaa tietokanta
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

    # 3. Käsittely (jos avaus onnistui)
    if ($isOpened) {
        
        try {
            $access.DoCmd.SetWarnings($false)
            Write-Host "   - Access-varoitukset poistettu käytöstä."
            
            $vbaProject = $database.Application.VBE.ActiveVBProject
            
            if ($null -eq $vbaProject) {
                Write-Error "   ❌ KRIITTINEN VIRHE: VBA-projektiin (VBE) ei päästy käsiksi (palautti null)."
                Write-Error "     Vaikka bittisyys on oikein, varmista UUDELLEEN, että 'Trust access to the VBA project object model'"
                Write-Error "     on päällä Accessin asetuksissa (Tiedosto > Asetukset > Luottamuskeskus > Makroasetukset)."
                Write-Error "     Tarkista myös 'Luotetut sijainnit' (Trusted Locations) ja lisää sinne tiedoston ja komponenttien polut."
                throw "VBA Project is null. Check Access Trust Center settings."
            }
            
            Write-Host "   - VBA-projekti avattu onnistuneesti."

            # 3.1 Poista vanhat
            foreach ($name in $componentNames) {
                try {
                    $component = $vbaProject.VBComponents.Item($name)
                    $vbaProject.VBComponents.Remove($component)
                    Write-Host "   ✅ Poistettiin vanha komponentti: $name"
                }
                catch {
                    Write-Host "   - Info: Komponenttia $name ei löytynyt poistettavaksi (tämä on ok)."
                }
            }

            # 3.2 Tuo uudet
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
            
            # 3.3 Tallenna ja sulje
            $acCmdSaveDatabase = 19
            $database.Application.DoCmd.RunCommand($acCmdSaveDatabase) 
            Write-Host "   ✅ Tietokanta tallennettiin paikoilleen."
            
            $access.DoCmd.SetWarnings($true)
            
            $access.CloseCurrentDatabase()
            Write-Host "   ✅ Tiedosto $databasePath päivitetty onnistuneesti."

        }
        catch {
            # Tämä 'catch' nappaa VBA-käsittelyn virheet
            Write-Error "   ❌ Virhe VBA-käsittelyssä tai tallennuksessa: $($_.Exception.Message)"
            
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
    # Tämä 'catch' nappaa kaikki ylemmän tason virheet (esim. New-Object, polkujen tarkistus, tiedoston avaus)
    Write-Error "--- KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA ---"
    Write-Error $_.Exception.Message
    
}
finally {
    # --- PAKOTETTU SIIVOUS ---
    # Tämä lohko suoritetaan AINA, vaikka skripti kaatuisi tai onnistuisi.
    # Tämä estää zombie-prosessien syntymisen.
    
    Write-Host "--- Siivotaan ja suljetaan Access-prosessi... ---"
    
    if ($null -ne $access) {
        try {
            $access.Quit()
            Write-Host "   - Access.Quit() kutsuttu."
        }
        catch {
            Write-Warning "   - Access.Quit() epäonnistui (prosessi oli ehkä jo kaatunut)."
        }
        
        # Odotetaan hetki, jotta prosessi ehtii sulkeutua ennen COM-objektin tuhoamista
        Start-Sleep -Seconds 1 
        
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
            Write-Host "   - COM-objekti vapautettu."
        }
        catch {
            # Tämä voi epäonnistua, jos $access-muuttujaa ei koskaan luotu kunnolla
        }
    }
    
    Remove-Variable access -ErrorAction SilentlyContinue
    Remove-Variable database -ErrorAction SilentlyContinue
    
    Write-Host "--- Siivous valmis. ---"
}