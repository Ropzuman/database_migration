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
# - Nykyinen toteutus: lue .bas/.cls-tiedosto -> poista headerit -> kirjoita puhdas koodi InsertLines(1, ...)-funktiolla.
# - EI käytetä AddFromString() - se tuottaa tyhjiä sulkeita () moduulin loppuun DeleteLines-ajon jälkeen.
# - Tämä vastaa manuaalista kopioi-liitä -toimintoa VBA-editorissa.

# KORJATTU (tarkistettu versio):
#   - AutomationSecurity asetetaan nyt ForceDisable (3) Low (1) sijaan. VBE/VBOM-käyttöoikeus
#     tulee Access-asetuksen "Trust access to the VBA project object model" kautta, EI
#     AutomationSecurity-arvosta - se säätelee vain makrojen automaattista suoritusta avattaessa.
#     Low sallisi myös AutoExec-makron/lomakkeen Form_Open-koodin suorituksen näkymättömässä
#     Access-istunnossa, mikä voi jumittaa skriptin näkymättömään valintaikkunaan.
#   - Lisätty Shift-ohitus (kuten eräajoversiossa) lisävarmistukseksi AutoExecia vastaan.
#   - VBA-lähdetiedostot luetaan nyt oikealla merkistöllä: Access vie .cls/.bas-tiedostot
#     oletuksena ANSI (Windows-1252) -muodossa, ei UTF-8:na.
#   - $component ja $codeModule vapautetaan nyt ReleaseComObject:lla jokaisen komponentin
#     jälkeen (aiemmin vapauttamatta - COM-viitteiden vuoto).
#   - Lisätty valinnainen kääntövaihe (Debug > Compile) ennen kannan sulkemista.

# --- KRIITTINEN TARKISTUS: Bittisyys ---
if ([System.IntPtr]::Size -ne 8) {
    Write-Error "VIRHE: Tämä skripti on suoritettava 64-bittisessä (x64) PowerShellissä."
    Write-Error "Sulje tämä (x86) ikkuna ja käynnistä normaali 'Windows PowerShell'."
    # return on ISE-yhteensopiva; exit 1 lopettaisi koko ISE-istunnon
    return
}
Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Ajetaan 64-bittisessä PowerShellissä." -ForegroundColor Green

# --- Win32 API: Shift-näppäimen simulointi AutoExec-makron ohittamiseksi ---
$signature = @'
[DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
'@
$win32 = Add-Type -MemberDefinition $signature -Name "Win32Kb" -Namespace "Win32" -PassThru

# Palauttaa tiedoston sisällön oikealla merkistöllä.
# Access vie VBA-moduulit ANSI:na (Windows-1252) ellei tiedostossa ole BOM-merkkiä,
# joten pelkkä UTF-8-oletus rikkoo ääkköset (ä, ö, å) hiljaisesti.
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

# Määritellään muuttujat ennalta 'finally'-lohkoa varten
$access = $null
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
    $defaultCompDisplay = if ([string]::IsNullOrWhiteSpace($DefaultComponentPath)) { "(ei oletusta asetettu)" } else { $DefaultComponentPath }
    Write-Host "Oletuspolku komponenteille: $defaultCompDisplay" -ForegroundColor Cyan
    $inputComponent = Read-Host -Prompt 'Lisää polku komponenttitiedostoille (.bas/.cls) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputComponent)) {
        if ([string]::IsNullOrWhiteSpace($DefaultComponentPath)) {
            Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultComponentPath skriptin alussa."
            throw "No component path provided"
        }
        $componentPath = $DefaultComponentPath
    }
    else { $componentPath = $inputComponent }

    Write-Host "`nVAIHE 2: Päivitettävä Access-tiedosto" -ForegroundColor Magenta
    $defaultAccessDisplay = if ([string]::IsNullOrWhiteSpace($DefaultAccessFilePath)) { "(ei oletusta asetettu)" } else { $DefaultAccessFilePath }
    Write-Host "Oletuspolku Access-tiedostolle: $defaultAccessDisplay" -ForegroundColor Cyan
    $inputAccess = Read-Host -Prompt 'Lisää polku Access-tiedostoon (.accdb) (paina Enter käyttääksesi oletusta)'
    if ([string]::IsNullOrWhiteSpace($inputAccess)) {
        if ([string]::IsNullOrWhiteSpace($DefaultAccessFilePath)) {
            Write-Error "Polkua ei annettu eikä oletusta ole asetettu. Aseta `$DefaultAccessFilePath skriptin alussa."
            throw "No Access file path provided"
        }
        $databasePath = $DefaultAccessFilePath
    }
    else { $databasePath = $inputAccess }

    # Polkujen tarkistus — throw käytetään exit 1:n sijaan, jotta finally-lohko ajetaan aina
    if (-not (Test-Path $componentPath -PathType Container)) {
        Write-Error "Komponenttikansio ei löydy: '$componentPath'"
        throw "Invalid component path: $componentPath"
    }
    if (-not (Test-Path $databasePath -PathType Leaf)) {
        Write-Error "Access-tiedostoa ei löydy: '$databasePath'"
        throw "Invalid database path: $databasePath"
    }

    # --- 3. Skannaa komponentit automaattisesti ---
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') [KOMPONENTIT] Skannataan .bas ja .cls -tiedostot kansiosta: $componentPath" -ForegroundColor Cyan
    # Rakennetaan hajautustaulu (nimi → tiedosto) deduplikointia varten.
    # Jos samalla nimellä on sekä .bas että .cls, .bas saa etusijan (se lisätään viimeisenä).
    $componentMap = @{}
    @(Get-ChildItem -Path $componentPath -Filter "*.cls") | ForEach-Object { $componentMap[$_.BaseName] = $_ }
    @(Get-ChildItem -Path $componentPath -Filter "*.bas") | ForEach-Object { $componentMap[$_.BaseName] = $_ }

    if ($componentMap.Count -eq 0) {
        Write-Error "Ei löytynyt yhtään .bas tai .cls -tiedostoa kansiosta: $componentPath"
        throw "No component files found"
    }

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Löytyi $($componentMap.Count) komponenttia:" -ForegroundColor Green
    $componentMap.Keys | Sort-Object | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

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

            # 3 = msoAutomationSecurityForceDisable: estää makrojen (myös AutoExecin ja
            # lomakkeiden Form_Open/Load-koodin) automaattisen suorituksen avattaessa.
            # HUOM: Tämä EI estä pääsyä VBA-projektiin - se hallitaan erikseen Access-asetuksen
            # "Trust access to the VBA project object model" kautta, joka pitää olla päällä
            # koneella jolla skripti ajetaan.
            $access.AutomationSecurity = 3

            # Painetaan Shift pohjaan lisävarmistukseksi AutoExecia vastaan
            $win32::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 200

            # OpenCurrentDatabase(FilePath, [Exclusive], [Password])
            $access.OpenCurrentDatabase($databasePath, $false, "")
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
        finally {
            # Vapautetaan Shift aina, onnistui avaus tai ei
            $win32::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)

    # --- 5. VBA-komponenttien käsittely ---
    if ($isOpened) {

        try {
            # Kytke Accessin sisäiset varoitukset pois päältä
            $access.DoCmd.SetWarnings($false)
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Access-varoitukset poistettu käytöstä."

            # KRIITTINEN: VBA-projektiin pääsy suoraan $access-objektista
            # VBA-projektiin pääsee vain $access.VBE.ActiveVBProject-kautta (ei DAO-tietokantaobjektin kautta)
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
            $successCount = 0
            $failureCount = 0
            foreach ($name in ($componentMap.Keys | Sort-Object)) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')       [KÄSITTELY] $name" -ForegroundColor Cyan

                $fullModulePath = $componentMap[$name].FullName
                $isBoundComponent = $false  # Form_* ja Report_* ovat sidottuja — ei voi luoda Add()-komennolla
                $ext = $componentMap[$name].Extension

                if ($ext -eq ".bas") {
                    $componentType = 1  # vbext_ct_StdModule
                }
                elseif ($name -like "Form_*" -or $name -like "Report_*") {
                    $isBoundComponent = $true
                    $componentType = 100  # Sidottu komponentti — ei luotavissa Add():lla
                }
                else {
                    $componentType = 2  # vbext_ct_ClassModule
                }

                if (-not (Test-Path $fullModulePath)) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Komponenttitiedostoa $(Split-Path $fullModulePath -Leaf) ei löydy levyltä. Ohitetaan päivitys." -ForegroundColor Red
                    $failureCount++
                    continue
                }

                $component = $null
                $codeModule = $null
                try {
                    $moduleContent = Read-VbaSourceFile -Path $fullModulePath

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
                    # KORJAUS: Tarkistetaan että $codeStartIndex on validi (estää PowerShellin käänteisen taulukon)
                    if ($codeStartIndex -gt ($lines.Count - 1)) {
                        $cleanCode = ""
                    }
                    else {
                        $cleanCode = ($lines[$codeStartIndex..($lines.Count - 1)] -join "`r`n")
                    }

                    # KRIITTINEN: AddFromString liittää koodin moduulin loppuun implisiittiseen
                    # "tyhjään tilaan", mikä tuottaa tyhjiä sulkeita () DeleteLines-ajon jälkeen.
                    # Korjaus: käytetään InsertLines(1, ...) joka KIRJOITTAA riville 1 liittämisen sijaan.
                    # Varmistetaan silti, että cleanCode päättyy CRLF-rivinvaihtoon.
                    # Tarkistetaan tyhjyys ENNEN CRLF-liitosta — muuten \r\n menee IsNullOrWhiteSpace-testin läpi
                    if ([string]::IsNullOrWhiteSpace($cleanCode)) {
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ⚠ VAROITUS: Tiedosto $name on tyhjä tai sisältää vain headerit. Ohitetaan." -ForegroundColor Yellow
                        $failureCount++
                        continue
                    }
                    $cleanCode = $cleanCode.TrimEnd([char]13, [char]10) + "`r`n"

                    # Etsi tai luo komponentti
                    try {
                        $component = $vbaProject.VBComponents.Item($name)
                        Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ Komponentti löytyi, päivitetään sisältö..."
                    }
                    catch {
                        # Lomakkeita (Form_*) ja raportteja (Report_*) ei voi luoda VBComponents.Add()-komennolla!
                        # Ne ovat sidottuja komponentteja ja pitää olla jo olemassa kannassa.
                        if ($isBoundComponent) {
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Sidottu komponentti $name ei ole olemassa kannassa." -ForegroundColor Red
                            Write-Host "$(Get-Date -Format 'HH:mm:ss')             Lomakkeita (Form_*) ja raportteja (Report_*) ei voi luoda automaattisesti, vain päivittää." -ForegroundColor Yellow
                            $failureCount++
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

                    # Lisää uusi koodi InsertLines-menetelmällä riville 1.
                    # EI käytetä AddFromString() - se tuottaa tyhjiä sulkeita () moduulin loppuun.
                    $codeModule.InsertLines(1, $cleanCode)

                    $newLineCount = $codeModule.CountOfLines
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✓ VALMIS: $name ($oldLineCount → $newLineCount riviä)" -ForegroundColor Green
                    $successCount++

                }
                catch {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')          ✗ VIRHE: Komponentin $name päivitys epäonnistui: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')             Virhetyyppi: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
                    Write-Host "$(Get-Date -Format 'HH:mm:ss')             Stack: $($_.ScriptStackTrace)" -ForegroundColor Yellow
                    $failureCount++
                }
                finally {
                    # Vapautetaan tämän komponentin COM-viitteet ennen seuraavaan siirtymistä.
                    if ($null -ne $codeModule) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($codeModule) | Out-Null }
                    if ($null -ne $component) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($component) | Out-Null }
                }
            }

            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [KOMPONENTIT] Kaikki komponentit käsitelty."
            Write-Host ""
            Write-Host "$(Get-Date -Format 'HH:mm:ss') === YHTEENVETO ===" -ForegroundColor Cyan
            Write-Host "  Onnistuneet: $successCount / $($componentMap.Count)" -ForegroundColor Green
            if ($failureCount -gt 0) {
                Write-Host "  Epäonnistuneet: $failureCount" -ForegroundColor Red
            }

            # 5.2 Käännetään ja tallennetaan moduulit ennen sulkemista, jotta rikkinäinen liitetty
            # koodi näkyy tässä lokissa eikä jää huomaamatta seuraavaan manuaaliseen avaukseen asti.
            # 126 = acCmdCompileAndSaveAllModules (tarkistettu erikseen — vääriä RunCommand-arvoja
            # on aiemmin päätynyt tähän koodikantaan, joten arvoa ei pidä muuttaa ilman varmistusta).
            try {
                $access.RunCommand(126)
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Moduulit käännetty ja tallennettu." -ForegroundColor DarkGreen
            }
            catch {
                Write-Host "$(Get-Date -Format 'HH:mm:ss')    ⚠ Kääntäminen epäonnistui - tarkista moduulit käsin (Debug → Compile): $($_.Exception.Message)" -ForegroundColor Yellow
            }

            # 5.3 Tallenna ja sulje
            # HUOM: CloseCurrentDatabase() tallentaa automaattisesti VBA-projektin.
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    [TALLENNUS] Suljetaan ja tallennetaan tietokanta..."

            # Laita varoitukset takaisin päälle
            $access.DoCmd.SetWarnings($true)

            # CloseCurrentDatabase tallentaa VBA-projektin automaattisesti
            $access.CloseCurrentDatabase()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Tiedosto $databasePath päivitetty onnistuneesti!" -ForegroundColor Green

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

    # Suljetaan Access ensin — tämä mitätöi lasten COM-viittaukset (VBProject ym.) turvallisesti
    if ($null -ne $access) {
        try {
            $access.Quit()
            Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ Access.Quit() suoritettu."
        }
        catch {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss')    ⚠ Access.Quit() epäonnistui (prosessi oli ehkä jo kaatunut)."
        }
        Start-Sleep -Milliseconds 500
    }

    # Vapautetaan kaikki COM-viittaukset Quit():n jälkeen
    foreach ($obj in @($vbaProject, $access)) {
        if ($null -ne $obj) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        }
    }
    Write-Host "$(Get-Date -Format 'HH:mm:ss')    ✓ COM-objektit vapautettu."

    # Pakotetaan roskienkeruu COM-viitteiden välittömäksi vapauttamiseksi
    # Ilman tätä Access.exe voi jäädä prosessilistalle kunnes GC ajaa automaattisesti
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable access -ErrorAction SilentlyContinue
    Remove-Variable vbaProject -ErrorAction SilentlyContinue

    Write-Host "$(Get-Date -Format 'HH:mm:ss') [OK] Siivous valmis." -ForegroundColor Green
}