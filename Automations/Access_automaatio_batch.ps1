# Access_automaatio_Batch.ps1
# Eräajo, joka päivittää kaikki tietokannat kerralla alikansioista.
# Sisältää Shift-ohituksen AutoExec- ja käynnistysmakrojen estämiseksi.
#
# KORJATTU (tarkistettu versio):
#   - Tietokantojen haku on nyt rekursiivinen (-Recurse), kuten kommentti aina väitti.
#   - VBA-lähdetiedostot luetaan oikealla merkistöllä: Access vie .cls/.bas-tiedostot
#     oletuksena ANSI (Windows-1252) -muodossa, ei UTF-8:na. Pelkkä UTF-8-oletus
#     rikkoo ääkköset (ä, ö, å) hiljaisesti ilman virheilmoitusta.
#   - Access-prosessin pakkosulkeminen ei enää perustu Process.MainWindowHandle-
#     osumaan (joka on lähes aina 0 kun Visible=$false) vaan PID:iin, joka haetaan
#     COM-ikkunakahvasta GetWindowThreadProcessId-kutsulla.
#   - Lisätty valinnainen kääntövaihe (Debug > Compile) ennen kannan sulkemista.
#   - Lisätty varoitus, jos moduulikansiossa on saman niminen tiedosto useaan kertaan.

$ErrorActionPreference = 'Stop'

if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    return
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä. Aloitetaan eräajo.`n" -ForegroundColor Green

# --- Win32 API: Shift-näppäimen simulointi + ikkovan omistavan prosessin PID ---
$signature = @'
[DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@
$win32 = Add-Type -MemberDefinition $signature -Name "Win32Helpers" -Namespace "Win32" -PassThru

# Palauttaa tiedoston sisällön oikealla merkistöllä.
# Access vie VBA-moduulit ANSI:na (Windows-1252) ellei tiedostossa ole BOM-merkkiä,
# joten pelkkä UTF-8-oletus rikkoo ääkköset hiljaisesti.
function Read-VbaSourceFile {
    param([Parameter(Mandatory)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    else {
        return [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
    }
}

# --- 1. Polkujen kysely (ILMAN OLETUKSIA) ---

Write-Host "VAIHE 1: Moduulien juurikansio" -ForegroundColor Magenta
$ModulesRoot = Read-Host -Prompt 'Syötä polku moduulikansioille'
if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { throw "Moduulien polkua ei annettu. Keskeytetään." }

Write-Host "`nVAIHE 2: Tietokantojen juurikansio" -ForegroundColor Magenta
$DatabasesRoot = Read-Host -Prompt 'Syötä polku tiedostokansioille'
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
    # 3 = msoAutomationSecurityForceDisable (estää kääntämisen avattaessa)
    $access.AutomationSecurity = 3

    # HUOM: haku on rekursiivinen, jotta alikansioissa olevat kannat löytyvät
    # (skriptin oma kuvaus lupaa tämän).
    $dbFiles = Get-ChildItem -Path $DatabasesRoot -Filter "*.accdb" -Recurse

    if ($dbFiles.Count -eq 0) {
        throw "Kansiosta $DatabasesRoot (tai sen alikansioista) ei löytynyt yhtään .accdb -tiedostoa."
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
        foreach ($file in @(Get-ChildItem -Path $modulePath -Filter "*.cls" -Recurse) + @(Get-ChildItem -Path $modulePath -Filter "*.bas" -Recurse)) {
            if ($componentMap.ContainsKey($file.BaseName)) {
                Write-Host "  ⚠ VAROITUS: Löytyi kaksi tiedostoa nimellä '$($file.BaseName)' - käytetään: $($file.FullName) (ohitettu: $($componentMap[$file.BaseName].FullName))" -ForegroundColor Yellow
            }
            $componentMap[$file.BaseName] = $file
        }

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
                    # Painetaan Shift pohjaan
                    $win32::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
                    Start-Sleep -Milliseconds 200

                    $access.OpenCurrentDatabase($dbPath, $false, "")
                    $kantaAvattu = $true
                    $access.Visible = $false
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tietokanta avattu Shift-ohituksella."
                }
                catch {
                    $retryCount++
                    if ($retryCount -lt $maxRetries) { Start-Sleep -Seconds $retryDelaySeconds }
                    else { throw $_ }
                }
                finally {
                    # Vapautetaan Shift aina
                    $win32::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
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

                        $moduleContent = Read-VbaSourceFile -Path $fullModulePath

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

                # Käännetään ja tallennetaan kaikki moduulit ennen sulkemista, jotta rikkinäinen
                # liitetty koodi näkyy raportissa eikä jää huomaamatta seuraavaan avaukseen asti.
                try {
                    $access.RunCommand(126) # acCmdCompileAndSaveAllModules
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Moduulit käännetty ja tallennettu." -ForegroundColor DarkGreen
                }
                catch {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ⚠ Kääntäminen epäonnistui - tarkista moduulit käsin: $($_.Exception.Message)" -ForegroundColor Yellow
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
        $accessProcess = $null
        try {
            # Haetaan prosessin PID ikkunakahvasta - toimii myös kun Visible=$false,
            # toisin kuin Process.MainWindowHandle-vertailu (joka on tällöin lähes aina 0).
            $accessPid = 0
            try {
                $accessWindowHandle = [IntPtr]$access.hWnd
                if ($accessWindowHandle -ne [IntPtr]::Zero) {
                    $pidOut = [uint32]0
                    [void]$win32::GetWindowThreadProcessId($accessWindowHandle, [ref]$pidOut)
                    $accessPid = $pidOut
                }
            } catch { $accessPid = 0 }

            if ($accessPid -ne 0) {
                $accessProcess = Get-Process -Id $accessPid -ErrorAction SilentlyContinue
            }

            Write-Host "$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Suljetaan Access..." -ForegroundColor Gray
            $access.Quit()

            # Odotetaan hetki (max 2 sekuntia), että prosessi sulkeutuu itse
            $counter = 0
            while ($accessProcess -and -not $accessProcess.HasExited -and $counter -lt 20) {
                Start-Sleep -Milliseconds 100
                $counter++
                $accessProcess.Refresh()
            }
        }
        catch {
            Write-Warning "Sulkeminen epäonnistui tai Access oli jo kiinni: $($_.Exception.Message)"
        }
        finally {
            # Jos prosessi roikkuu edelleen pystyssä, tapetaan se väkisin PID:n perusteella
            if ($null -ne $accessProcess -and -not $accessProcess.HasExited) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Access ei sulkeutunut. Pakkolopetetaan prosessi (PID: $($accessProcess.Id))..." -ForegroundColor Yellow
                Stop-Process -Id $accessProcess.Id -Force -ErrorAction SilentlyContinue
            }

            # Vapautetaan COM-objektin viittaukset muistista
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null } catch {}
            $access = $null
        }
    }

    # Pakotetaan .NET-roskienkeruu ajoon
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Eräajo suoritettu ja siivous valmis." -ForegroundColor Green
}
