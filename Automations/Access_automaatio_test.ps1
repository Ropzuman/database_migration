# Access_automaatio.ps1
# ... (Alkuperäiset kommentit säilytetty) ...

# Asetetaan virheiden tiukka hallinta
$ErrorActionPreference = 'Stop'

if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    return
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä." -ForegroundColor Green

# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$access = $null
$vbe = $null
$vbaProject = $null

try {
    # --- 1. Alustus ---
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [ALUSTUS] Luodaan Access COM-objekti..." -ForegroundColor Cyan
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    
    # --- 2. Polkujen kysely ---
    # Aseta oletukset tähän nopeuttaaksesi testausta, tai jätä tyhjäksi jolloin skripti kysyy
    $DefaultComponentPath = ''   # esim. 'C:\database_migration\Access\MAINEQ'
    $DefaultAccessFilePath = ''  # esim. 'C:\polku\tietokanta.accdb'

    Write-Host "`nVAIHE 1: Komponenttien lähde" -ForegroundColor Magenta
    $defaultCompDisplay = if ([string]::IsNullOrWhiteSpace($DefaultComponentPath)) { "(ei oletusta)" } else { $DefaultComponentPath }
    Write-Host "Oletuspolku: $defaultCompDisplay" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Polku komponenttitiedostoille (.bas/.cls) [Enter = oletus]'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) {
        if ([string]::IsNullOrWhiteSpace($DefaultComponentPath)) { throw "Komponenttipolkua ei annettu." }
        $componentPath = $DefaultComponentPath
    }
    else { $componentPath = $inputComponent }

    Write-Host "`nVAIHE 2: Päivitettävä Access-tiedosto" -ForegroundColor Magenta
    $defaultAccessDisplay = if ([string]::IsNullOrWhiteSpace($DefaultAccessFilePath)) { "(ei oletusta)" } else { $DefaultAccessFilePath }
    Write-Host "Oletuspolku: $defaultAccessDisplay" -ForegroundColor Cyan
    $inputAccess = Read-Host -Prompt 'Polku Access-tiedostoon (.accdb) [Enter = oletus]'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) {
        if ([string]::IsNullOrWhiteSpace($DefaultAccessFilePath)) { throw "Access-tiedostopolkua ei annettu." }
        $databasePath = $DefaultAccessFilePath
    }
    else { $databasePath = $inputAccess }

    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Error "Komponenttikansio ei löydy: '$componentPath'"
        throw "Invalid component path: $componentPath"
    }
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Error "Access-tiedostoa ei löydy: '$databasePath'"
        throw "Invalid database path: $databasePath"
    }

    $componentMap = @{}
    @(Get-ChildItem -Path $componentPath -Filter "*.cls") | ForEach-Object { $componentMap[$_.BaseName] = $_ }
    @(Get-ChildItem -Path $componentPath -Filter "*.bas") | ForEach-Object { $componentMap[$_.BaseName] = $_ }

    if ($componentMap.Count -eq 0) { throw "No component files found" }

    # --- 4. Tiedoston valmistelu ja avaus ---
    try { Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force } catch { }

    $maxRetries = 3
    $retryDelaySeconds = 1
    $retryCount = 0
    $isOpened = $false

    do {
        try {
            $access.AutomationSecurity = 1
            $access.OpenCurrentDatabase($databasePath, $false, "")
            $isOpened = $true
            $access.Visible = $false
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tietokanta avattu onnistuneesti."
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) { Start-Sleep -Seconds $retryDelaySeconds }
            else { throw $_ }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)

    # --- 5. VBA-komponenttien käsittely ---
    if ($isOpened) {
        try {
            $access.DoCmd.SetWarnings($false)

            # KORJAUS: Vältetään Double-Dot COM -ongelma (implisiittiset objektit)
            $vbe = $access.VBE
            $vbaProject = $vbe.ActiveVBProject

            if ($null -eq $vbaProject) {
                throw "VBA Project is null. Check Access Trust Center settings."
            }

            $successCount = 0
            $failureCount = 0

            foreach ($name in ($componentMap.Keys | Sort-Object)) {
                # Määritellään iterointikohtaiset COM-muuttujat
                $component = $null
                $codeModule = $null

                try {
                    $fullModulePath = $componentMap[$name].FullName
                    $isBoundComponent = $false
                    $ext = $componentMap[$name].Extension

                    if ($ext -eq ".bas") { $componentType = 1 }
                    elseif ($name -match "^(Form_|Report_)") { $isBoundComponent = $true; $componentType = 100 }
                    else { $componentType = 2 }

                    # ... (Tiedoston luku ja Header-parsaus - säilyvät ennallaan, logiikkasi tässä oli erinomainen) ...
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
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ VAROITUS: Tiedosto $name tyhjä. Ohitetaan." -ForegroundColor Yellow
                        $failureCount++
                        continue
                    }
                    
                    $cleanCode = $cleanCode.TrimEnd([char]13, [char]10) + "`r`n"

                    # Etsi tai luo komponentti
                    try {
                        $component = $vbaProject.VBComponents.Item($name)
                    }
                    catch {
                        if ($isBoundComponent) { throw "Sidottua komponenttia ei voi luoda." }
                        $component = $vbaProject.VBComponents.Add($componentType)
                        $component.Name = $name
                    }

                    $codeModule = $component.CodeModule
                    $oldLineCount = $codeModule.CountOfLines
                    
                    if ($oldLineCount -gt 0) { $codeModule.DeleteLines(1, $oldLineCount) }
                    
                    $codeModule.InsertLines(1, $cleanCode)
                    $successCount++
                }
                catch {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: $name - $($_.Exception.Message)" -ForegroundColor Red
                    $failureCount++
                }
                finally {
                    # KORJAUS: Vapautetaan iterointikohtaiset COM-objektit heti, 
                    # jotta Garbage Collector ei tukehdu ja prosessi pääsee sulkeutumaan.
                    if ($null -ne $codeModule) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($codeModule) | Out-Null }
                    if ($null -ne $component) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($component) | Out-Null }
                }
            }
            
            # Tallenna ja sulje
            $access.DoCmd.SetWarnings($true)
            $access.CloseCurrentDatabase()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tiedosto päivitetty onnistuneesti!" -ForegroundColor Green

        }
        catch {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') ✗ VIRHE VBA-käsittelyssä: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [FATAL] KRIITTINEN VIRHE: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [CLEANUP] Siivotaan..." -ForegroundColor Magenta
    
    if ($null -ne $access) {
        try { $access.Quit() } catch { }
    }

    # KORJAUS: Vapautetaan nyt myös $vbe, jotta koko ketju puhdistuu
    foreach ($obj in @($vbaProject, $vbe, $access)) {
        if ($null -ne $obj) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Siivous valmis." -ForegroundColor Green
}