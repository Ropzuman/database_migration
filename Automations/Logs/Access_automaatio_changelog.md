# Muutosloki – Access_automaatio.ps1

**Tiedosto:** `Automations/Access_automaatio.ps1`  
**Päivämäärä:** 2026-03-06  
**Tekijä:** GitHub Copilot (Senior System Architect – 64-bit Migration Agent)  
**Peruste:** Code Review -raportit (`Automations/access_automaatio code review`) — 3 iteraatiota (v1 → v2 → v3)  
**Yhteenveto:** 1 kriittinen, 3 korkean prioriteetin, 3 keskitason, 4 matalan prioriteetin löydöstä — kaikki korjattu. Lisäksi 1 kosmeettinen v3-löydös.

---

## 🚨 Kriittiset muutokset

### CRIT 1 – `exit 1` ohitti `finally`-lohkon → Access-haamuproessit

**Ongelma:** Polkujen tarkistus (`Test-Path`) ajettiin **sen jälkeen**, kun `New-Object -ComObject Access.Application` oli jo luonut COM-objektin (rivi 39). Molemmissa tarkistuksissa käytettiin `exit 1`, joka PowerShell 5.1:ssä ohittaa `finally`-lohkon kokonaan — Access.exe jäi käyntiin taustaprosessina ilman mahdollisuutta sammuttaa se ohjelmallisesti.

**Korjaukset:**

- `exit 1` korvattu `throw`-lauseella molemmissa polkujen tarkistuksissa → `finally` ajetaan aina
- Bittisuustarkistuksessa `exit 1` + `Start-Sleep` korvattu `return`-lauseella (ISE-yhteensopivuus)

```powershell
# Ennen: if (-not (Test-Path ...)) { exit 1 }
# Jälkeen:
if (-not (Test-Path $componentPath -PathType Container)) {
    Write-Error "Komponenttikansio ei löydy: '$componentPath'"
    throw "Invalid component path: $componentPath"
}
if (-not (Test-Path $databasePath -PathType Leaf)) {
    Write-Error "Access-tiedostoa ei löydy: '$databasePath'"
    throw "Invalid database path: $databasePath"
}
```

---

## ⚠️ Korkean prioriteetin muutokset

### HIGH 1 – `Report_*`-komponentit loivat väärän ClassModulen hiljaa

**Ongelma:** `Form_*`-komponentteja varten oli olemassa suoja (`$isBoundComponent = $true`), mutta `Report_*`-komponenteilta se puuttui. Kun `Report_Asiakkaat.cls` käsiteltiin eikä sitä löytynyt VBA-projektista, koodi putosi läpi `VBComponents.Add(2)`-kutsuun — luoden orvon `ClassModule`-nimeltä `Report_Asiakkaat` eikä koskaan päivittänyt oikeaa raporttimoduulia.

**Korjaus:** Lisätty `elseif ($name -like "Report_*")` -haara, joka asettaa saman `$isBoundComponent = $true` -lipun:

```powershell
if ($name -like "Form_*") {
    $isBoundComponent = $true
    $componentType = 100  # sidottu — ei luotavissa Add():lla
}
elseif ($name -like "Report_*") {
    $isBoundComponent = $true
    $componentType = 100  # sidottu — ei luotavissa Add():lla
}
```

---

### HIGH 2 – UTF-8 BOM injektoitui VBA-koodiin

**Ongelma:** `Get-Content -Encoding UTF8` voi PowerShell 5.1:ssä palauttaa UTF-8 BOM -tavun (U+FEFF) merkkijonon ensimmäisenä merkkinä, jolloin BOM injektoitui `InsertLines()`-kutsun kautta VBA-moduuliin.

**Korjaus:** `Get-Content` korvattu `System.IO.StreamReader`:lla BOM-automaattiohjauksella + erillinen varmuustarkistus:

```powershell
$reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
$moduleContent = $reader.ReadToEnd()
$reader.Close()
if ($moduleContent.Length -gt 0 -and [int][char]$moduleContent[0] -eq 0xFEFF) {
    $moduleContent = $moduleContent.Substring(1)
}
```

**Myöhemmin (v2):** `StreamReader` kääritty `try-finally`-lohkoon — `Dispose()` kutsutaan myös `ReadToEnd()`-poikkeuksen sattuessa, jottei tiedostokahva jää auki:

```powershell
$reader = $null
try {
    $reader = [System.IO.StreamReader]::new($fullModulePath, [System.Text.Encoding]::UTF8, $true)
    $moduleContent = $reader.ReadToEnd()
}
finally {
    if ($null -ne $reader) { $reader.Dispose(); $reader = $null }
}
```

---

### HIGH 3 – `AutomationSecurity`-kommentti oli täysin väärä

**Ongelma:** Kommentti sanoi *"estää automaattisen makrojen suorituksen"* — arvo 1 (`msoAutomationSecurityLow`) tekee täysin päinvastaisen: se **sallii** kaikki makrot. Asetus on välttämätön VBE-rajapinnan käyttöön, mutta harhaanjohtava kommentti olisi johtanut ylläpitäjän poistamaan sen.

**Korjaus:**

```powershell
# Aseta msoAutomationSecurityLow jotta VBA-projektiin päästään käsiksi.
# HUOM: Arvo 1 SALLII kaikki makrot — tämä on tarkoituksellista, VBE-rajapinta vaatii sen.
# Access on näkymätön ($access.Visible = $false), joten tietoturvariski on rajattu.
$access.AutomationSecurity = 1  # msoAutomationSecurityLow — sallii VBA-projektin muokkauksen
```

---

## ℹ️ Keskitason muutokset

### MED 1 – Kuollut `$database`-muuttuja poistettu

**Ongelma:** `$database = $access.CurrentDb()` tallensi DAO-tietokantaobjektin muuttujaan, jota ei koskaan käytetty mihinkään logiikkaan — se vain vapautettiin `finally`-lohkossa. Stale COM-viite ilman käyttötarkoitusta.

**Korjaus:** `$database`-muuttuja poistettu kokonaan skriptistä. `finally`-lohkon vapautussilmukka päivitetty vastaavasti.

---

### MED 2 – Tyhjyystarkistus tehtiin CRLF-lisäyksen jälkeen

**Ongelma:** `IsNullOrWhiteSpace($cleanCode)` tarkistettiin sen jälkeen kun `$cleanCode` oli jo saanut `\r\n`-liitoksen — tarkistus läpäisi tyhjätkin moduulit, koska `"\r\n"` ei ole `IsNullOrWhiteSpace`-testissä tyhjä.

**Korjaus:** Tarkistus siirretty ennen CRLF-liitosta. Kommentti selittää järjestysvaatimuksen:

```powershell
# Tarkistetaan tyhjyys ENNEN CRLF-liitosta — muuten \r\n menee IsNullOrWhiteSpace-testin läpi
if ([string]::IsNullOrWhiteSpace($cleanCode)) {
    $failureCount++
    continue
}
$cleanCode = $cleanCode.TrimEnd([char]13, [char]10) + "`r`n"
```

---

### MED 3 – Tyhjä oletuspolku antoi harhaanjohtavan virheilmoituksen

**Ongelma:** Kun `$DefaultComponentPath = ''` ja käyttäjä painoi Enter, seuraava `Test-Path`-kutsu epäonnistui viestillä `path does not exist:` — tyhjä polku näkyi selittämättömänä.

**Korjaus:** Eksplisiittinen tarkistus ennen polkukyselyä + `(ei oletusta asetettu)` -näyttö:

```powershell
$defaultCompDisplay = if ([string]::IsNullOrWhiteSpace($DefaultComponentPath)) {
    "(ei oletusta asetettu)"
} else { $DefaultComponentPath }
Write-Host "Oletuspolku komponenteille: $defaultCompDisplay" -ForegroundColor Cyan
```

---

## 💡 Pienet parannukset

### LOW 1 – Tiedostojen skannaus: hajautustaulu deduplikointia varten (v2)

**Ongelma:** `$basFiles + $clsFiles` -rakenne ei deduplikoinut nimiä. Jos kansiossa oli sekä `Module1.bas` että `Module1.cls`, `$componentNames` sisälsi `"Module1"` kahdesti → käsiteltiin kahdesti, `.cls` ohitettiin hiljaa, yhteenveto-laskuri oli väärä.

**Korjaus:** Hajautustaulu `$componentMap` (nimi → tiedosto-objekti). `.bas` ylikirjoittaa `.cls`-merkinnän törmäystilanteessa:

```powershell
$componentMap = @{}
@(Get-ChildItem -Path $componentPath -Filter "*.cls") | ForEach-Object { $componentMap[$_.BaseName] = $_ }
@(Get-ChildItem -Path $componentPath -Filter "*.bas") | ForEach-Object { $componentMap[$_.BaseName] = $_ }
```

---

### LOW 2 – `$failureCount` puuttui tiedosto-guard-haarasta (v2)

**Ongelma:** Kaikki muut virhepolut (`catch`, tyhjä moduuli, sidottu komponentti) kasvattivat `$failureCount`-laskuria, mutta `if (-not $fullModulePath)` -haara käytti vain `continue`-lausetta. Yhteenveto aliraportoi virheitä.

**Korjaus:** `$failureCount++` lisätty tiedoston puuttumiseen:

```powershell
if (-not (Test-Path $fullModulePath)) {
    Write-Host "✗ VIRHE: Komponenttitiedostoa ... ei löydy levyltä." -ForegroundColor Red
    $failureCount++
    continue
}
```

---

### LOW 3 – Käsittelyn yhteenveto (`YHTEENVETO`) lisätty

**Ongelma:** Skripti tulosti aina `"Kaikki komponentit käsitelty!"` riippumatta siitä, kuinka monta komponenttia epäonnistui.

**Korjaus:** `$successCount` ja `$failureCount` -laskurit + `YHTEENVETO`-blokki ajon lopussa:

```powershell
Write-Host "$(Get-Date -Format 'HH:mm:ss') === YHTEENVETO ===" -ForegroundColor Cyan
Write-Host "  Onnistuneet: $successCount / $($componentMap.Count)" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "  Epäonnistuneet: $failureCount" -ForegroundColor Red
}
```

---

### LOW 4 – `finally` järjestys korjattu + `$database`-kommentti päivitetty (v2+v3)

**Ongelma 1:** `$vbaProject` vapautettiin ennen `$access.Quit()` -kutsua → mahdollinen COM-fault lapsen vapautuksessa ennen vanhemman sulkemista.

**Korjaus:** `Quit()` ensin, sitten kaikki COM-objektit `foreach`-silmukalla, lopuksi `GC::Collect()`:

```powershell
if ($null -ne $access) {
    try { $access.Quit() } catch { }
    Start-Sleep -Milliseconds 500
}
foreach ($obj in @($vbaProject, $access)) {
    if ($null -ne $obj) {
        try { [Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
    }
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
```

**Ongelma 2 (v2→v3):** Kommentti viittasi poistettuun `$database`-muuttujaan.

**Korjaus (v3):** Kommentti päivitetty kuvaamaan oikeaa käyttömallia:

```powershell
# VBA-projektiin pääsee vain $access.VBE.ActiveVBProject-kautta (ei DAO-tietokantaobjektin kautta)
```

---

### LOW 5 – Käsittelyjärjestys deterministiseksi (v3, kosmeettinen)

**Ongelma:** `foreach ($name in $componentMap.Keys)` iteroi hajautustaulun avaimia satunnaisessa järjestyksessä. Listausaskel käytti oikein `Sort-Object`, mutta käsittelysilmukka ei, mikä teki lokivertailun eri ajojen välillä hankalaksi.

**Korjaus:**

```powershell
foreach ($name in ($componentMap.Keys | Sort-Object)) {
```

---

## Ei muutettu (tunnetut rajoitukset)

| Löydös | Syy |
|--------|-----|
| **Polkujen tarkistus ennen `New-Object`** | Review suositteli siirtämistä COM-luonnin eteen. Jätetty `try`-lohkon sisälle, koska `throw` takaa jo `finally`-suorituksen. Refaktorointi ei muuta toiminnallista lopputulosta. |
| **`.cls`-tiedostojen tyypin automaattinen tunnistaminen** | Luokkamoduulin tyypin (`ClassModule` vs `Form`/`Report`) voisi tunnistaa tiedoston `VERSION 1.0 CLASS` -headerista. Jätetty toteutettavaksi tarvittaessa. |
