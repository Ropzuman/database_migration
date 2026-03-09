# Access_automaatio_Batch_v4.ps1
# Eräajo, joka päivittää kaikki tietokannat kerralla alikansioista.
# Kysyy polut (Moduulit ensin, ilman oletuksia) ja lisää automaattisesti uudet moduulit.

$ErrorActionPreference = 'Stop'

if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    return
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä. Aloitetaan eräajo.`n" -ForegroundColor Green

# --- 1. Polkujen kysely (ILMAN OLETUKSIA) ---

Write-Host "VAIHE 1: Moduulien juurikansio" -ForegroundColor Magenta
$ModulesRoot = Read-Host -Prompt 'Syötä polku moduulikansioille'
if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { throw "Moduulien polkua ei annettu. Keskeytetään." }

Write-Host "`nVAIHE 2: Tietokantojen juurikansio" -ForegroundColor Magenta
$DatabasesRoot = Read-Host -Prompt 'Syötä polku tietokansioille (.accdb)'
if ([string]::IsNullOrWhiteSpace($DatabasesRoot)) { throw "Tietokantojen polkua ei annettu. Keskeytetään." }

# Validoidaan polut
if (-not (Test-Path $ModulesRoot -PathType Container)) { throw "Moduulikansiota ei löydy: $ModulesRoot" }
if (-not (Test-Path $DatabasesRoot -PathType Container)) { throw "Tietokantakansiota ei löydy: $DatabasesRoot" }

# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$access = $null
$raportti = @()

try {
    # --- 2. Alustus ---
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [ALUSTUS] Luodaan Access COM-objekti..." -ForegroundColor Cyan
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.AutomationSecurity = 1

    $dbFiles = Get-ChildItem -Path $DatabasesRoot -Filter "*.accdb"
    
    if ($dbFiles.Count -eq 0) {
        throw "Kansiosta $DatabasesRoot ei löytynyt yhtään .accdb -tiedostoa."
    }

    Write-Host "$(Get-Date -Format 'HH:mm:ss') Löydettiin $($dbFiles.Count) tietokantaa. Aloitetaan käsittely.`n" -ForegroundColor Cyan

    # === 3. MASTER LOOP (Käydään kannat läpi) ===
    foreach ($dbFile in $dbFiles) {
        $dbName = $dbFile.BaseName
        $dbPath = $dbFile.FullName
        $modulePath = Join-Path $ModulesRoot $dbName

        Write-Host ("=" * 60) -ForegroundColor Gray
        Write-Host "$(Get-Date -Format 'HH:mm:ss') [KANTA] $dbName" -ForegroundColor Magenta

        # Tarkistetaan onko kanta lukittu
        $lockFile = Join-Path $dbFile.DirectoryName ($dbName + ".laccdb")
        if (Test-Path $lockFile) {
            Write-Host "  ✗ VAROITUS: Kanta on todennäköisesti käytössä (.laccdb löytyy). Ohitetaan." -ForegroundColor Yellow
            $raportti += [PSCustomObject]@{ Kanta = $dbName; Tila = "Ohitettu (Lukittu)"; Onnistuneet = 0; Uudet = 0; Virheet = 0 }
            continue
        }

        # Tarkistetaan löytyykö moduulikansiota
        if (-not (Test-Path $modulePath)) {
            Write-Host "  ⚠ Moduulikansiota ei löydy polusta: $modulePath. Ohitetaan." -ForegroundColor Yellow
            $raportti += [PSCustomObject]@{ Kanta = $dbName; Tila = "Ohitettu (Ei moduuleja)"; Onnistuneet = 0; Uudet = 0; Virheet = 0 }
            continue
        }

        # Etsitään moduulit alikansioita myöten
        $componentMap = @{}
        @(Get-ChildItem -Path $modulePath -Filter "*.cls" -Recurse) | ForEach-Object { $componentMap[$_.BaseName] = $_ }
        @(Get-ChildItem -Path $modulePath -Filter "*.bas" -Recurse) | ForEach-Object { $componentMap[$_.BaseName] = $_ }

        if ($componentMap.Count -eq 0) {
            Write-Host "  ⚠ Kansiossa $modulePath (tai alikansioissa) ei ole VBA-tiedostoja. Ohitetaan." -ForegroundColor Yellow
            $raportti += [PSCustomObject]@{ Kanta = $dbName; Tila = "Ohitettu (Tyhjä kansio)"; Onnistuneet = 0; Uudet = 0; Virheet = 0 }
            continue
        }

        # Tietokantakohtaiset COM-muuttujat ja laskurit
        $vbe = $null
        $vbaProject = $null
        $dbSuccess = 0
        $dbNew = 0
        $dbFail = 0
        $kantaAvattu = $false

        try {
            try { Set-ItemProperty -Path $dbPath -Name IsReadOnly -Value $false -Force } catch { }

            # Avataan kanta retry-logiikalla
            $maxRetries = 3
            $retryDelaySeconds = 1
            $retryCount = 0

            do {
                try {
                    $access.OpenCurrentDatabase($dbPath, $false, "")
                    $kantaAvattu = $true
                    $access.Visible = $false
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tietokanta avattu."
                }
                catch {
                    $retryCount++
                    if ($retryCount -lt $maxRetries) { Start-Sleep -Seconds $retryDelaySeconds }
                    else { throw $_ }
                }
            } while (-not $kantaAvattu -and $retryCount -lt $maxRetries)

            if ($kantaAvattu) {
                $access.DoCmd.SetWarnings($false)

                $vbe = $access.VBE
                $vbaProject = $vbe.ActiveVBProject

                if ($null -eq $vbaProject) {
                    throw "VBA Project is null. Tarkista Accessin Trust Center -makroasetukset."
                }

                # === 4. INNER LOOP (Käydään komponentit läpi) ===
                foreach ($name in ($componentMap.Keys | Sort-Object)) {
                    $component = $null
                    $codeModule = $null
                    $isNewModule = $false

                    try {
                        $fullModulePath = $componentMap[$name].FullName
                        $isBoundComponent = $false
                        $ext = $componentMap[$name].Extension

                        if ($ext -eq ".bas") { $componentType = 1 }
                        elseif ($name -match "^(Form_|Report_)") { $isBoundComponent = $true; $componentType = 100 }
                        else { $componentType = 2 }

                        $reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
                        $moduleContent = $reader.ReadToEnd()
                        $reader.Dispose()

                        if ($moduleContent.Length -gt 0 -and [int][char]$moduleContent[0] -eq 0xFEFF) {
                            $moduleContent = $moduleContent.Substring(1)
                        }

                        $lines = $moduleContent -split "`r?`n"
                        $codeStartIndex = 0
                        $inHeader = $true

                        for ($i = 0; $i -lt $lines.Count; $i++) {
                            $line = $lines[$i].Trim()
                            if ($inHeader) {
                                if ($line -match "^VERSION\s+" -or $line -match "^BEGIN\s*" -or $line -match "^END\s*$" -or
                                    $line -match "^Attribute\s+VB_" -or $line -match "^MultiUse\s*=" -or $line -eq "") {
                                    $codeStartIndex = $i + 1
                                }
                                else {
                                    $inHeader = $false; break
                                }
                            }
                        }

                        if ($codeStartIndex -gt ($lines.Count - 1)) { $cleanCode = "" }
                        else { $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n") }

                        if ([string]::IsNullOrWhiteSpace($cleanCode)) {
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ $name on tyhjä tiedosto. Ohitetaan." -ForegroundColor Yellow
                            $dbFail++
                            continue
                        }
                        
                        $cleanCode = $cleanCode.TrimEnd([char]13, [char]10) + "`r`n"

                        # TARKISTUS: Löytyykö vanha vai luodaanko uusi
                        try {
                            $component = $vbaProject.VBComponents.Item($name)
                        }
                        catch {
                            if ($isBoundComponent) { throw "Lomakkeen/raportin koodimoduulia ei voi luoda skriptillä tyhjästä." }
                            $component = $vbaProject.VBComponents.Add($componentType)
                            $component.Name = $name
                            $isNewModule = $true
                        }

                        $codeModule = $component.CodeModule
                        $oldLineCount = $codeModule.CountOfLines
                        
                        if ($oldLineCount -gt 0) { $codeModule.DeleteLines(1, $oldLineCount) }
                        
                        $codeModule.InsertLines(1, $cleanCode)
                        
                        if ($isNewModule) {
                            $dbNew++
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          + Lisätty uusi: $name" -ForegroundColor DarkCyan
                        }
                        else {
                            $dbSuccess++
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ Päivitetty: $name" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE ($name): $($_.Exception.Message)" -ForegroundColor Red
                        $dbFail++
                    }
                    finally {
                        if ($null -ne $codeModule) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($codeModule) | Out-Null }
                        if ($null -ne $component) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($component) | Out-Null }
                    }
                }
                
                $raportti += [PSCustomObject]@{ Kanta = $dbName; Tila = "OK"; Onnistuneet = $dbSuccess; Uudet = $dbNew; Virheet = $dbFail }
            }
        }
        catch {
            Write-Host "  ✗ KANNAN KÄSITTELY EPÄONNISTUI: $($_.Exception.Message)" -ForegroundColor Red
            $raportti += [PSCustomObject]@{ Kanta = $dbName; Tila = "VIRHE"; Onnistuneet = $dbSuccess; Uudet = $dbNew; Virheet = $dbFail + 1 }
        }
        finally {
            if ($null -ne $vbaProject) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbaProject) | Out-Null }
            if ($null -ne $vbe) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbe) | Out-Null }
            
            if ($kantaAvattu) {
                try {
                    $access.DoCmd.SetWarnings($true)
                    $access.CloseCurrentDatabase()
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Kanta suljettu ja tallennettu." -ForegroundColor DarkGreen
                }
                catch { 
                    Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Kannan sulkeminen epäonnistui."
                }
            }
        }
    }

    # === 5. LOPPURAPORTTI ===
    Write-Host "`n"
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host " ERÄAJON YHTEENVETO" -ForegroundColor Cyan
    Write-Host ("=" * 65) -ForegroundColor Cyan
    $raportti | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Cyan

}
catch {
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [FATAL] KRIITTINEN VIRHE ERÄAJOSSA: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Siivotaan Access-prosessi..." -ForegroundColor Magenta
    
    if ($null -ne $access) {
        try { $access.Quit() } catch { }
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null } catch { }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Eräajo suoritettu ja siivous valmis." -ForegroundColor Green
}