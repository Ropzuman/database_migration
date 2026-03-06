# Muutosloki – Excel_automaatio.ps1

**Tiedosto:** `Automations/Excel_automaatio.ps1`  
**Päivämäärä:** 2026-03-06  
**Tekijä:** GitHub Copilot (Senior System Architect – 64-bit Migration Agent)  
**Peruste:** Code Review -raportti (`Automations/excel_automaatio code review`)  
**Yhteenveto:** 2 kriittistä, 3 korkean prioriteetin, 2 keskitason ja 2 matalan prioriteetin löydöstä — kaikki korjattu.

---

## 🚨 Kriittiset muutokset

### CRIT 1 – Atominen tiedoston korvaus (tietohäviön esto)

**Ongelma:** Skripti poisti alkuperäisen tiedoston (`Remove-Item`) ennen kuin uudelleennimeäminen (`Rename-Item`) oli suoritettu. Jos `Rename-Item` epäonnistui (verkkoyhteysongelma, lukkovirhe, liian pitkä polku), alkuperäinen tiedosto oli jo tuhlattu ja temp-tiedosto jäi hylynä levylle — palauttamaton datan menetys.

**Korjaus:** Toteutettu `.bak`-varmuuskopioon perustuva atominen korvausstrategia:

1. `Move-Item` siirtää alkuperäisen `→ .bak` (palautettavissa)
2. `Rename-Item` nimeää temp-tiedoston lopulliseksi
3. Onnistuessa `.bak` poistetaan; epäonnistuessa `.bak` palautetaan alkuperäiseksi

```powershell
Move-Item -Path $workbookPath -Destination $backupPath -Force -ErrorAction Stop
try {
    Rename-Item -Path $tempWorkbookPath -NewName (Split-Path $workbookPath -Leaf) -Force -ErrorAction Stop
    Remove-Item -Path $backupPath -Force -ErrorAction SilentlyContinue
}
catch {
    Move-Item -Path $backupPath -Destination $workbookPath -Force
    throw
}
```

---

### CRIT 2 – Retry-silmukan `throw` lopetti koko eräajon

**Ongelma:** Kun tiedoston avaaminen epäonnistui kaikkien yritysten jälkeen, `throw $_` propagoitui suoraan `ForEach-Object`-putkeen ja lopetti kaikkien jäljellä olevien tiedostojen käsittelyn. Yksi lukittu tiedosto esti koko eräajon.

**Korjaus:** `throw $_` korvattu `$isOpened = $false` + `break` -rakenteella. Ulompi `if ($isOpened)`-tarkistus ohittaa tiedoston ja `ForEach-Object` jatkaa normaalisti.

```powershell
# Ennen:  throw $_
# Jälkeen:
$isOpened = $false
break  # Ei throw — ForEach-Object jatkaa seuraavaan tiedostoon
```

---

## ⚠️ Korkean prioriteetin muutokset

### HIGH 1 – UTF-8 BOM injektoitui VBA-koodiin

**Ongelma:** `Get-Content -Encoding UTF8` voi PowerShell 5.1:ssä palauttaa UTF-8 BOM -tavun (U+FEFF) merkkijonon ensimmäisenä merkkinä. Header-parseri ei tunnistanut BOM-prefiksillä varustettua `Option Explicit` -riviä, jolloin BOM injektoitui `AddFromString()`-kutsun kautta VBA-moduuliin aiheuttaen vaikeasti diagnosoitavia käännösvirheitä.

**Korjaus:** `Get-Content` korvattu `System.IO.StreamReader`:lla, joka käsittelee BOM:n automaattisesti. Lisäksi erillinen tarkistus varmuuden vuoksi:

```powershell
$reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
$moduleContent = $reader.ReadToEnd()
$reader.Close()
if ($moduleContent.Length -gt 0 -and [int][char]$moduleContent[0] -eq 0xFEFF) {
    $moduleContent = $moduleContent.Substring(1)
}
```

---

### HIGH 2 – `String.Replace(".xlsm")` korruptoi polkuja, joissa kansionimessä on `.xlsm`

**Ongelma:** `$workbookPath.Replace(".xlsm", "_MIGRATED.xlsm")` korvasi **kaikki** `.xlsm`-esiintymät koko polkumerkkijonossa. UNC-polku kuten `\\server\proj.xlsm files\tiedosto.xlsm` muuttui virheelliseksi `\\server\proj_MIGRATED.xlsm files\tiedosto_MIGRATED.xlsm`.

**Korjaus:** `System.IO.Path`-metodit käsittelevät vain tiedostonimen osan:

```powershell
$wbDir  = [System.IO.Path]::GetDirectoryName($workbookPath)
$wbStem = [System.IO.Path]::GetFileNameWithoutExtension($workbookPath)
$tempWorkbookPath = [System.IO.Path]::Combine($wbDir, $wbStem + "_MIGRATED.xlsm")
```

---

### HIGH 3 – COM-objekteja ei vapautettu eksplisiittisesti → Excel-haamuproessit

**Ongelma:** `$workbook = $null` ei vapauta COM-viittausta välittömästi — .NET:n roskienkeruu voi viivästyä sekunteista minuutteihin, jolloin Excel.exe jää henkiin taustaprosessina. Virhetilanteessa myöskään `finally`-lohko ei koskaan ajanut GC:tä.

**Korjaukset:**

- `$workbook.Close()` → `$workbook.Close($false)` + `ReleaseComObject($workbook)` ennen `$null`-nollausta (sekä onnistumis- että virhepolulla)
- `finally`-lohkoon lisätty `[System.GC]::Collect()` ja `[System.GC]::WaitForPendingFinalizers()` välittömän vapauttamisen varmistamiseksi

---

## ℹ️ Keskitason muutokset

### MED 1 – Stale temp-tiedostoa ei tarkistettu ennen `SaveAs`-kutsua

**Ongelma:** Jos edellinen ajo epäonnistui `SaveAs`→`Rename`-välissä, levylle jäi `_MIGRATED.xlsm`-jäänne. Seuraava ajo kutsui `SaveAs` samalle tiedostopolulle ilman tarkistusta — Excel joko ylikirjoitti sen hiljaa tai heitti lupaevirheen.

**Korjaus:** Lisätty `Test-Path`-tarkistus ennen `SaveAs`-kutsua:

```powershell
if (Test-Path $tempWorkbookPath) {
    Write-Warning "⚠ Väliaikainen tiedosto löytyi jäänteenä edellisestä ajosta: $tempWorkbookPath"
    Remove-Item -Path $tempWorkbookPath -Force -ErrorAction Stop
}
```

---

### MED 2 – `exit 1` lopetti koko PowerShell ISE -istunnon

**Ongelma:** Bittisyystarkistuksessa `Start-Sleep -Seconds 10` + `exit 1` lopetti koko ISE-istunnon, jos skriptiä ajettiin ISE-ympäristössä. `Start-Sleep` oli myös tarpeeton, koska virheviesti jää näkyville joka tapauksessa.

**Korjaus:** `Start-Sleep` poistettu, `exit 1` korvattu `return`-lauseella:

```powershell
# Ennen: Start-Sleep -Seconds 10; exit 1
# Jälkeen:
return  # ISE-yhteensopiva; exit 1 lopettaisi koko ISE-istunnon
```

---

## 💡 Pienet parannukset

### LOW 1 – Tyhjä oletuspolku antoi harhaanjohtavan virheilmoituksen

**Ongelma:** Kun `$DefaultModulePath = ''` ja käyttäjä painoi Enter, `Test-Path ""` epäonnistui viestillä `Module files folder does not exist:` — tyhjä polku oli selittämätön.

**Korjaus:** Lisätty eksplisiittinen tarkistus tyhjälle oletukselle. Promptissa näytetään `(ei oletusta asetettu)` selkeästi:

```powershell
$defaultModuleDisplay = if ([string]::IsNullOrWhiteSpace($DefaultModulePath)) { "(ei oletusta asetettu)" } else { $DefaultModulePath }
Write-Host "Oletuspolku moduuleille: $defaultModuleDisplay" -ForegroundColor Cyan
```

---

### LOW 2 – `Get-ChildItem` ei skannaa alihakemistoja (dokumentoitu rajoitus)

**Muutos:** Lisätty kommentti, joka dokumentoi eksplisiittisesti, että `.bas`-tiedostoja skannataan vain ylätasolta:

```powershell
$basFiles = Get-ChildItem -Path $modulePath -Filter "*.bas"  # vain ylätaso, ei alihakemistoja
```

---

## Ei muutettu (tunnetut rajoitukset)

| Löydös | Syy |
|--------|-----|
| **LOW 2 `-Recurse`** | Skriptin käyttötapaus on tasainen kansiohierarkia. `-Recurse` lisättäisiin vain, jos moduulit organisoidaan alihakemistoihin — vaatii muutoksia myös `$name`-resoluutiologiikkaan. |
| **SQL-injektio VBA-koodeissa** | Käsitelty erillisessä `Module1.bas`-refaktoroinnissa (`OnTurvallinenSQL`). |

---

## Delta Review v2 – lisäkorjaukset (2026-03-06)

### MED (uusi) – `StreamReader` ei ollut `try-finally`-suojassa → tiedostokahvan vuoto

**Ongelma:** Jos `ReadToEnd()` heitti poikkeuksen (levyvirhe, lukuoikeus evätty kesken), suoritus hyppäsi ulompaan `catch`-lohkoon ja `reader.Close()` jäi ajamatta. `StreamReader` piti tiedostokahvan auki GC:n keräilyyn asti — isommilla erillä toistuvat virheet kasasivat avoimia kahvoja.

**Korjaus:** `StreamReader`-elinikä kääritty omaan `try-finally`-lohkoon:

```powershell
$reader = $null
try {
    $reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
    $moduleContent = $reader.ReadToEnd()
}
finally {
    if ($null -ne $reader) { $reader.Close(); $reader = $null }
}
```

---

### LOW (uusi) – Palautusvirhe korvaa alkuperäisen virheviestin

**Ongelma:** Jos `Rename-Item` epäonnistui ja lisäksi `Move-Item`-palautus epäonnistui, `throw` heitti palautusvirheen eikä alkuperäistä `Rename-Item`-virhettä. Operaattori näki väärän viestin eikä saanut tietoa tiedostojen tilasta.

**Korjaus:** Alkuperäinen viesti tallennetaan ennen palautusyritystä. Palautuksen epäonnistuessa operaattorille kerrotaan kummankin tiedoston sijainti:

```powershell
$renameError = $_.Exception.Message
# ... palautusyritys ...
throw [System.Exception]::new("Rename failed: $renameError", $_.Exception)
```

---

### LOW (carried) – Ajokohtainen yhteenveto lisätty

**Ongelma:** Skripti tulosti aina `"Kaikki työkirjat käsitelty!"` riippumatta epäonnistumisista. Yksittäiset virheet hautautuivat lokitulosteeseen.

**Korjaus:** Viisi laskuria (`$wbSuccess`, `$wbSkipped`, `$wbFailed`, `$modSuccess`, `$modFailed`) + värikoodattu `YHTEENVETO`-blokki:

```powershell
Write-Host "$(Get-Date -Format 'HH:mm:ss') === YHTEENVETO ===" -ForegroundColor Cyan
Write-Host "  Työkirjat: $wbSuccess onnistui / $wbSkipped ohitettu / $wbFailed epäonnistui"
Write-Host "  Moduulit:  $modSuccess onnistui / $modFailed epäonnistui"
```

---

## Delta Review v3 – lisäkorjaukset (2026-03-06)

### LOW (uusi, kosmeettinen) – `$modFailed++` puuttui tiedosto-guard-haarasta

**Ongelma:** Kaikki muut virhepolut kasvattivat `$modFailed`-laskuria, mutta `if (-not (Test-Path $fullModulePath))` -haara käytti vain `continue`-lausetta — yhteenveto aliraportoi virheitä.

**Korjaus:** `$modFailed++` lisätty tiedoston puuttumisen haaran alkuun vastaavasti kuin Access-skriptissä.
