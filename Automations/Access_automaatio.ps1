# Access_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit (.bas) ja luokkamoduulit (.cls) kannassa.
# YMP?RIST?: T?m? skripti tukee sek? 32-bittist? ett? 64-bittist? Microsoft Accessia.
#           Skripti tunnistaa automaattisesti Accessin arkkitehtuurin ja k?ynnistyy
#           uudelleen oikeassa PowerShell-ymp?rist?ss? tarvittaessa.

# K?YTT?YTYMINEN:
# - Tunnistaa Accessin bittisyyden (32-bit tai 64-bit)
# - Jos PowerShell-bittisyys ei vastaa Accessia, k?ynnistyy automaattisesti uudelleen
# - Kysyy polun yhteen tietokantatiedostoon ja polun komponenttihakemistoon
# - Avaa .accdb-kannan, p?ivitt?? komponenttien sis?ll?n suoraan (moduulit/luokat)
# - K?ytt?? try...finally-lohkoa varmistaakseen, ett? Access-prosessi suljetaan aina

# T?RKE?? - VBComponents.Import-ongelma:
# - T?m? skripti KORVAA komponenttien sis?ll?n suoraan CodeModule-rajapinnan kautta.
# - EI k?ytet? VBComponents.Import()-funktiota, koska se lis?? n?kym?tt?mi? metatietoja.
# - Import() aiheuttaa komponenttien toimintah?iri?it? (k?ytt?ytyy eri tavalla kuin manuaalisesti kopioidut).
# - Nykyinen toteutus: lue .bas/.cls-tiedosto ? poista headerit ? kirjoita puhdas koodi AddFromString()-funktiolla.
# - T?m? vastaa manuaalista kopioi-liit? -toimintoa VBA-editorissa.

# --- KRIITTINEN TARKISTUS: Bittisyys ---
# Tunnista Accessin bittisyys rekisterist?
$accessPath = $null
$accessIs32Bit = $false

# Tarkista ensin 64-bit sijainti
$accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
if ($accessPath -and ($accessPath -match 'x86' -or $accessPath -match 'Program Files \(x86\)')) {
    $accessIs32Bit = $true
}

# Jos ei l?ytynyt, tarkista 32-bit sijainti (WOW6432Node)
if (-not $accessPath) {
    $accessPath = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE" -ErrorAction SilentlyContinue).'(default)'
    if ($accessPath) {
        $accessIs32Bit = $true
    }
}

if (-not $accessPath) {
    Write-Host "VIRHE: Microsoft Accessia ei l?ydy j?rjestelm?st?."
    Write-Host "Asenna Microsoft Access ennen t?m?n skriptin ajamista."
    Start-Sleep -Seconds 10
    exit 1
}

# Tarkista PowerShellin bittisyys
$psIs64Bit = [System.IntPtr]::Size -eq 8

# N?yt? tunnistetut arkkitehtuurit
Write-Host "=== BITTISYYS-TARKISTUS ===" -ForegroundColor Cyan
Write-Host "Access: $(if($accessIs32Bit){'32-bit'}else{'64-bit'}) ($accessPath)"
Write-Host "PowerShell: $(if($psIs64Bit){'64-bit'}else{'32-bit'})"

# Jos bittisyydet eiv?t t?sm??, n?yt? virhe ja lopeta
if ($accessIs32Bit -and $psIs64Bit) {
    Write-Host ""
    Write-Host "[ERROR] VIRHE: Bittisyydet eiv?t t?sm??!" -ForegroundColor Red
    Write-Host "   Access on 32-bittinen, mutta PowerShell on 64-bittinen." -ForegroundColor Yellow
    Write-Host "   64-bit PowerShell ei voi luoda COM-yhteytt? 32-bit Accessiin." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "RATKAISU: Aja skripti batch-tiedostolla:" -ForegroundColor Cyan
    Write-Host "   Kaksoisklikkaa: RUN_ACCESS_AUTOMATION.bat" -ForegroundColor Green
    Write-Host ""
    Write-Host "TAI k?ynnist? manuaalisesti 32-bit PowerShelliss?:" -ForegroundColor Cyan
    Write-Host "   1. Avaa K?ynnist?-valikko" -ForegroundColor White
    Write-Host "   2. Etsi 'PowerShell (x86)'" -ForegroundColor White
    Write-Host "   3. Suorita siell?:" -ForegroundColor White
    Write-Host "      cd c:\database_migration\Automations" -ForegroundColor Gray
    Write-Host "      .\Access_automaatio.ps1" -ForegroundColor Gray
    Write-Host ""
    Start-Sleep -Seconds 15
    exit 1
}
elseif (-not $accessIs32Bit -and -not $psIs64Bit) {
    Write-Host ""
    Write-Host "[ERROR] VIRHE: Bittisyydet eiv?t t?sm??!" -ForegroundColor Red
    Write-Host "   Access on 64-bittinen, mutta PowerShell on 32-bittinen." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "RATKAISU: K?ynnist? 64-bit PowerShelliss?:" -ForegroundColor Cyan
    Write-Host "   1. Avaa K?ynnist?-valikko" -ForegroundColor White
    Write-Host "   2. Etsi 'PowerShell' (ei x86)" -ForegroundColor White
    Write-Host "   3. Suorita siell?:" -ForegroundColor White
    Write-Host "      cd c:\database_migration\Automations" -ForegroundColor Gray
    Write-Host "      .\Access_automaatio.ps1" -ForegroundColor Gray
    Write-Host ""
    Start-Sleep -Seconds 15
    exit 1
}
else {
    Write-Host "[OK] Bittisyydet t?sm??v?t - jatketaan..." -ForegroundColor Green
}

Write-Host ""


# M??ritell??n muuttujat ennalta 'finally'-lohkoa varten
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
    $inputAccess = Read-Host -Prompt 'Lis?? polku Access-tiedostoon (.accdb) (paina Enter k?ytt??ksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) { $databasePath = $DefaultAccessFilePath } else { $databasePath = $inputAccess }

    Write-Host "Component files folder: $DefaultComponentPath" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Lis?? polku komponenttitiedostoille (paina Enter k?ytt??ksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) { $componentPath = $DefaultComponentPath } else { $componentPath = $inputComponent }

    # Polkujen tarkistus
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Host "Access-tiedostoa ei l?ydy tai polku on hakemisto: $databasePath"
        exit 1
    }
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Host "Component files folder does not exist: $componentPath"
        exit 1
    }

    # --- 3. Komponenttien m??rittely ---
    # M??rittele kaikki ne moduulien ja luokkamoduulien nimet, jotka poistetaan ja tuodaan.
    # ?l? k?yt? tiedostop??tteit? (.bas/.cls) nimiss?.
    # 
    # P?IVITETTY 2025-11-08: Phase 1 cleanup (dead code removal, Replace() optimization)
    # - Poistettu: Form_USysRevText_OLD (dead code, korvattu Form_USysRevText:ll?)
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

    Write-Host "--- K?SITTELL??N: $databasePath ---"

    $retryCount = 0
    $isOpened = $false

    # --- 4. Tiedoston valmistelu ja avaus ---
    
    # Poista 'Vain luku' -attribuutti
    try {
        Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
        Write-Host "   [OK] Poistettiin Vain luku -attribuutti."
    }
    catch {
        Write-Warning "   [WARN] Vain luku -attribuutin poisto failed: $($_.Exception.Message). Jatketaan silti."
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
                Write-Warning "   [WARN] Avaus failed: $($_.Exception.Message). Yritet??n uudelleen (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
                Write-Host "   [ERROR] VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen j?lkeen."
                throw $_ # Heitet??n virhe p??-try-lohkolle
            }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)

    # --- 5. VBA-komponenttien k?sittely ---
    if ($isOpened) {
        
        try {
            # Kytke Accessin sis?iset varoitukset pois p??lt?
            $access.DoCmd.SetWarnings($false)
            Write-Host "   - Access-varoitukset poistettu k?yt?st?."
            
            # Yrit? saada yhteys VBA-projektiin
            $vbaProject = $database.Application.VBE.ActiveVBProject
            
            # KRIITTINEN TARKISTUS: Jos $vbaProject on $null, Accessin turva-asetukset est?v?t toiminnon.
            if ($null -eq $vbaProject) {
                Write-Host "   [ERROR] KRIITTINEN VIRHE: VBA-projektiin (VBE) ei p??sty k?siksi (palautti null)."
                Write-Host "     SYY: Accessin turva-asetukset est?v?t t?m?n. Tarkista seuraavat:"
                Write-Host "     1. Access - Asetukset - Luottamuskeskus - Luota VBA-projektin objektimallin k?ytt??n"
                Write-Host "     2. Access - Asetukset - Luottamuskeskus - Luotetut sijainnit - lis?? polut: $databasePath ja $componentPath"
                Write-Host "     3. Tiedoston Ominaisuudet - Salli eli Unblock, jos se on ladattu verkosta"
                throw "VBA Project is null. Check Access Trust Center settings."
            }
            
            Write-Host "   - VBA-projekti avattu onnistuneesti."

            # 5.1 P?ivit? komponenttien sis?lt? suoraan (v?ltt?? Import-metatietojen ongelman)
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
                    Write-Host "   [ERROR] VIRHE: Komponenttitiedostoa $name.bas tai $name.cls ei l?ydy polusta $componentPath. Ohitetaan p?ivitys."
                    continue
                }
                
                try {
                    # Lue .bas/.cls-tiedoston sis?lt?
                    $moduleContent = Get-Content -Path $fullModulePath -Raw -Encoding UTF8
                    
                    # Poista VBA-tiedoston header-rivit (.cls: VERSION, BEGIN/END, Attribute; .bas: Attribute)
                    # S?ilytet??n vain varsinainen VBA-koodi (Option/Declare/Function/Sub/Dim jne.)
                    $lines = $moduleContent -split "`r?`n"
                    $codeStartIndex = 0
                    
                    # K?y l?pi rivej? ja ohita kaikki header-rivit
                    for ($i = 0; $i -lt $lines.Count; $i++) {
                        $line = $lines[$i].Trim()
                        
                        # Ohita VERSION, BEGIN, END, Attribute, MultiUse ja tyhj?t rivit
                        if ($line -match "^VERSION\s+" -or 
                            $line -match "^BEGIN$" -or 
                            $line -match "^END$" -or 
                            $line -match "^Attribute\s+" -or
                            $line -match "^MultiUse\s*=" -or
                            $line -eq "") {
                            $codeStartIndex = $i + 1
                        }
                        else {
                            # Kun t?rm?t??n ensimm?iseen varsinaiseen koodiriviin, lopeta
                            break
                        }
                    }
                    
                    # Ota vain VBA-koodi (ilman header-rivej?)
                    $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n").Trim()
                    
                    # Etsi tai luo komponentti
                    $component = $null
                    try {
                        $component = $vbaProject.VBComponents.Item($name)
                        Write-Host "   - Komponentti $name l?ytyi, p?ivitet??n sis?lt?..."
                    }
                    catch {
                        # Jos komponenttia ei ole, luo se
                        Write-Host "   - Komponenttia $name ei l?ytynyt, luodaan uusi..."
                        $component = $vbaProject.VBComponents.Add($componentType)
                        $component.Name = $name
                    }
                    
                    # Tyhjenn? vanha koodi ja aseta uusi
                    $codeModule = $component.CodeModule
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    $codeModule.AddFromString($cleanCode)
                    
                    Write-Host "   [OK] Paivitettiin $name - $(($cleanCode -split "`n").Count) rivia koodia"
                    
                }
                catch {
                    Write-Host "   [ERROR] VIRHE: Komponentin $name p?ivitys failed: $($_.Exception.Message)"
                }
            }
            
            # 5.3 Tallenna ja sulje
            $acCmdSaveDatabase = 19
            $database.Application.DoCmd.RunCommand($acCmdSaveDatabase) 
            Write-Host "   [OK] Tietokanta tallennettiin paikoilleen."
            
            # Laita varoitukset takaisin p??lle
            $access.DoCmd.SetWarnings($true)
            
            $access.CloseCurrentDatabase()
            Write-Host "   [OK] Tiedosto $databasePath p?ivitetty onnistuneesti."

        }
        catch {
            # T?m? 'catch' nappaa VBA-k?sittelyn virheet
            Write-Host "[ERROR] Virhe VBA-k?sittelyss? tai tallennuksessa: $($_.Exception.Message)"
            
            # Yritet??n siisti? tietokantayhteys
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
    # T?m? 'catch' nappaa kaikki ylemm?n tason virheet (esim. New-Object, polkujen tarkistus)
    Write-Host "--- KRIITTINEN VIRHE SKRIPTIN SUORITUKSESSA ---"
    Write-Error $_.Exception.Message
    
}
finally {
    # --- 6. PAKOTETTU SIIVOUS ---
    # T?m? lohko suoritetaan AINA, vaikka skripti kaatuisi tai onnistuisi.
    # T?m? est?? "zombie" (jumittuneiden) Access-prosessien syntymisen.
    
    Write-Host "--- Siivotaan ja suljetaan Access-prosessi... ---"
    
    if ($null -ne $access) {
        try {
            $access.Quit()
            Write-Host "   - Access.Quit() called."
        }
        catch {
            Write-Warning "   - Access.Quit() failed (prosessi oli ehk? jo kaatunut)."
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
