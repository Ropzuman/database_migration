# Access_automaatio.ps1

# TARKOITUS: Korvaa VBA-moduulit (.bas) ja luokkamoduulit (.cls) kannassa tehden siitä 64-bittiselle Officelle sopivan.

# KÄYTTÄYTYMINEN:
# - Kysyy polut tietokantojen hakemistoon ja komponenttitiedostojen hakemistoon.
# - Avaa jokaisen .accdb-tietokannan, poistaa määritellyt komponentit, tuo uudet
#   .bas- tai .cls-tiedostot ja tallentaa muutokset paikoilleen ennen sulkemista.
# - Käyttää retry-logiikkaa lukkojen kiertämiseksi.
# - Katso alempaa ohjeet polkuihin ja korvattavien komponenttien nimiin.

# - HUOM: Kun muutoksia ajetaan verkkosijaintiin, pitää käyttää verkkosijainnin nimeä \\proense01\projektit\ 
# - esim. "\\proense01\projektit\24PRO260 Vermo Lämmönsiirrinasema\Z\DB\"

# Muokatun tiedoston voi tallentaa muodossa .ps1 haluamaansa sijaintiin ja suorittaa seuraavasti:
# - Avaa PowerShell Administratorina ja suorita komento Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# - Suorita "polku tiedostoon"\"tiedoston_nimi".ps1


$access = New-Object -ComObject Access.Application
# Huom: Accessin Visible-asetus (GUI) voi vaikuttaa siihen, miten se käsittelee tietokantoja.
# Joskus piilotettu tila ($false) on nopeampi, mutta testausta tarvitaan.
$access.Visible = $false
$access.DisplayAlerts = $false

# --- Polut ja asetukset (Päivitä oletusarvot tähän halutessasi. Varsinkin uusien moduulien sijainti yleensä kiinteä.) ---
$DefaultAccessFilesPath = ''
$DefaultComponentPath = ''

Write-Host "Access files folder: $DefaultAccessFilesPath" -ForegroundColor Cyan
$inputAccess = Read-Host -Prompt 'Lisää polku Access kantaan (paina Enter käyttääksesi oletusta)'
if ([string]::IsNullOrWhiteSpace($inputAccess)) { $accessFilesPath = $DefaultAccessFilesPath } else { $accessFilesPath = $inputAccess }

Write-Host "Component files folder: $DefaultComponentPath" -ForegroundColor Cyan
$inputComponent = Read-Host -Prompt 'Lisää polku komponenttitiedostoille (paina Enter käyttääksesi oletusta)'
if ([string]::IsNullOrWhiteSpace($inputComponent)) { $componentPath = $DefaultComponentPath } else { $componentPath = $inputComponent }

# Validate paths
if (-not (Test-Path $accessFilesPath -PathType Container)) {
    Write-Error "Access files folder does not exist: $accessFilesPath"
    exit 1
}
if (-not (Test-Path $componentPath -PathType Container)) {
    Write-Error "Component files folder does not exist: $componentPath"
    exit 1
}

# --- Komponenttien nimet (HUOM: Määrittele kaikki moduulit ja luokat tähän) ---
# Tässä määritellään kaikki ne komponentit, jotka poistetaan ja tuodaan uudelleen.
# Älä sisällytä tiedostopäätettä (.bas/.cls) nimiin.
$componentNames = @(
    "Module1", 
    "Module2", 
    "Module3", 
    "cls_DatabaseHandler", 
    "cls_AutoCADHelper"
) 

# Retry-asetukset
$maxRetries = 3
$retryDelaySeconds = 1

# Käsittele kaikki .accdb tiedostot kohdekansiossa
# Huom: Voit lisätä -Filter "*.mdb" jos sinulla on vanhempia tietokantoja
Get-ChildItem -Path $accessFilesPath -Filter "*.accdb" | ForEach-Object {
    $databasePath = $_.FullName
    Write-Host "--- KÄSITTELLÄÄN: $databasePath ---"

    $retryCount = 0
    $isOpened = $false
    $database = $null
    
    # 1. Poista tiedoston Vain luku -attribuutti
    Set-ItemProperty -Path $databasePath -Name IsReadOnly -Value $false -Force
    Write-Host "  ✅ Poistettiin Vain luku -attribuutti tiedostojärjestelmästä."

    # 2. Avaa tietokanta (Retry-logiikalla lukkojen kiertämiseksi)
    do {
        try {
            # Accessin OpenCurrentDatabase ei tue ReadOnly-parametria suoraan, 
            # joten luotamme tiedostojärjestelmän asetuksiin ja retry-logiikkaan.
            $access.OpenCurrentDatabase($databasePath) 
            $database = $access.CurrentDb()
            $isOpened = $true
            Write-Host "  - Tietokanta avattu onnistuneesti."
        } catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "  ⚠️ Avaus epäonnistui: $($_.Exception.Message). Yritetään uudelleen (Yritys $retryCount / $maxRetries)."
                Start-Sleep -Seconds $retryDelaySeconds
            } else {
                Write-Error "  ❌ VIRHE: Tiedostoa ei voitu avata $maxRetries yrityksen jälkeen. Jätetään käsittelemättä."
                $isOpened = $false
                # Nosta virhe, jotta päästään pää-catch-lohkoon
                throw $_
            }
        }
    } while (-not $isOpened -and $retryCount -lt $maxRetries)
    
    # Käsittely jatkuu vain, jos tiedosto avattiin onnistuneesti
    if ($isOpened) {
        
        try {
            # Tarkista, onko VBA-projekti käytettävissä
            $vbaProject = $database.Application.VBE.ActiveVBProject
            Write-Host "  - VBA-projekti avattu."

            # 3. Poista vanhat moduulit ja luokat
            foreach ($name in $componentNames) {
                try {
                    $component = $vbaProject.VBComponents.Item($name)
                    $vbaProject.VBComponents.Remove($component)
                    Write-Host "  ✅ Poistettiin vanha komponentti: $name"
                } catch {
                    # Hiljainen virhe, jos komponenttia ei löydy (se on ok)
                }
            }

            # 4. Tuo uudet moduulit ja luokat
            foreach ($name in $componentNames) {
                $basPath = Join-Path $componentPath "$($name).bas"
                $clsPath = Join-Path $componentPath "$($name).cls"
                $fullModulePath = $null

                # Tarkista, onko olemassa .bas vai .cls
                if (Test-Path $basPath) {
                    $fullModulePath = $basPath
                } elseif (Test-Path $clsPath) {
                    $fullModulePath = $clsPath
                }

                if (-not $fullModulePath) {
                    Write-Error "  ❌ VIRHE: Uutta komponenttitiedostoa ($name.bas/.cls) ei löydy polusta $componentPath. Ohitetaan tuonti."
                    continue
                }
                
                try {
                    $vbaProject.VBComponents.Import($fullModulePath)
                    Write-Host "  ✅ Tuotiin uusi komponentti: $name"
                } catch {
                    Write-Error "  ❌ VIRHE: Komponentin $name tuonti epäonnistui: $($_.Exception.Message)"
                }
            }
            
            # 5. Tallenna muutokset ja sulje tietokanta
            
            # Pakotetaan tietokannan tallennus. acCmdSaveDatabase = 19
            $acCmdSaveDatabase = 19
            $database.Application.DoCmd.RunCommand($acCmdSaveDatabase) 
            Write-Host "  ✅ Tietokanta tallennettiin paikoilleen."
            
            # Sulje tietokanta (tämä vapauttaa tiedostolukon)
            $access.CloseCurrentDatabase()
            Write-Host "  ✅ Tiedosto $databasePath päivitetty onnistuneesti."

        } catch {
            # Virheenkäsittely
            Write-Error "❌ Virhe VBA-käsittelyssä tai tallennuksessa: $($_.Exception.Message)"
            
            # Suljetaan tietokanta virhetilanteen sattuessa
            try {
                $access.CloseCurrentDatabase() 
            } catch {
                # Hiljainen virhe sulkemisessa, jos se on jo suljettu
            }
        }
    } # end if ($isOpened)
} # end ForEach-Object

# 6. Sulje Access
Write-Host "--- Suljetaan Access-prosessi ---"
$access.Quit()

# Vapauta COM-objekti muistista
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null 
Remove-Variable access -ErrorAction SilentlyContinue