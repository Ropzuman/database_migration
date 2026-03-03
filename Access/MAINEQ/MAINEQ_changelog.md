# Muutosloki — MAINEQ-moduuli

**Tiedosto:** `Access/MAINEQ/` (DataToACAD.bas, GeneralCodes.bas, USysCheck.bas, Form_DBUsers.cls, ja muut)
**Päivämäärä:** 2025-11-11 / 2025-11-12
**Tekijä:** 64-bit migraatioprojekti (Pro gradu)

---

## Kriittiset muutokset (Vaihe 1–2)

### 64-bittinen API-korjaus — USysCheck.bas

- `GetUserNameA` ja `GetComputerNameA`-funktioiden `Declare`-lauseet päivitetty:
  - Lisätty `PtrSafe`-avainsana 64-bittistä Office/AutoCAD-ympäristöä varten
  - `nSize`-parametri muutettu `LongPtr` → `Long` (Win32 API kirjoittaa 4-tavuisen DWORD:in, ei 8-tavuisen)
- `GetLastError`- ja `SHGetPathFromIDList`-funktiot päivitetty `PtrSafe`-määreellä

### Option Explicit — kaikki moduulit

- Lisätty `Option Explicit` kaikkiin moduuleihin, joista se puuttui (6 lomaketta)

### Me.-etuliite — Access-lomakkeet

- Kaikki lomakeohjauspyynnöt päivitetty `Me.`-etuliitteellä (`ControlName` → `Me.ControlName`) implisiittisten viittausvirheiden estämiseksi

### Debug.Print-rivien poisto

- Poistettu noin 50 `Debug.Print`-riviä `DataToACAD.bas`-tiedostosta
- Poistettu 2 `Debug.Print`-riviä `MoottTilaus`-moduulista
- Poistettu 1 `Debug.Print`-rivi `Revisiointi`-moduulista

---

## Siivous ja optimointi (Vaihe 3)

### Kommenttien suomennus — DataToACAD.bas

- Moduuliotsikko käännetty suomeksi (Moduuli / Tarkoitus / Alkuperäinen / Muokattu)
- Kaikkien funktioiden otsikkolohkot käännetty: `CrsRefLink`, `get_filename`, `inch`, `makeFiles`, `MakeListNoLoopID`, `MakeLocFiles`
- Kaikki inline-kommentit käännetty suomeksi:
  - Muuttujan esittelykommentit (lainausmerkki, tiedostokahva, silmukkalaskurit jne.)
  - Ohjauslogiiikan kommentit (tiedostojen alustus, sulkumerkit, siivous, virheenkäsittely)
  - Silmukkakommentit (`' tables` → `' Käydään läpi kaikki taulukot` jne.)
- Englanninkieliset `MsgBox`-viestit käännetty suomeksi

### Kommenttien suomennus — GeneralCodes.bas

- Moduuliotsikko käännetty suomeksi
- Julkisten muuttujien otsikko käännetty: `' Public Variables for Revision Tracking'` → suomeksi
- Kaikkien funktioiden otsikkolohkot käännetty: `IsLoaded`, `HaeViimPaiva`, `Optiot`, `Positiot`, `Vaihekulma`, `MotKaapUh`, `LisaaNo`
- Mukautetun `Replace()`-funktion poisto-huomio tiivistetty suomeksi (4 riviä)
- Muuttujan esittelykommentit käännetty: `Original length`, `Numeric value`, `Loop counter` → suomeksi
- Kirjoitusvirhe korjattu: `vaihekulmna` → `vaihekulman`, `käytjölle` → `käytölle`

### Kommenttien suomennus — USysCheck.bas

- `OPENFILENAME`-tyypin esittelykommentti käännetty suomeksi
- `KORJATTU`-huomio laajennettu selittämään `nSize As Long` vs. `LongPtr`-valinta
- Moduulitason tilamuuttujien kommentit käännetty suomeksi
- `Show_last`- ja `Show_last_criteria`-funktioiden kommentit käännetty

### Kommenttien suomennus — Form_DBUsers.cls

- `WhosOn`-alirutiinin kaikki kommentit käännetty:
  - Tietuetyypin esittely, tietokantapolun haku, tiedostopolun muodostus
  - Tiedoston olemassaolon tarkistus, tietuesilmukka, virheotsikot
- `MsgBox`-viestit käännetty: `"Couldn't populate the list"` → `"Tiedostoa ei löytynyt."`, `"Error: "` → `"Virhe: "`

### Poistetut kehitysaikaiset merkinnät

- Poistettu `' Updated 2025-11-11: Added DAO prefix` -tyyppiset väliaikais-muistiinpanot `Form_DBUsers.cls`- ja `USysCheck.bas`-tiedostoista (3 kpl / tiedosto)

---

## Tiedostoyhteenveto

| Tiedosto | Vaihe 1–2 | Vaihe 3 |
|---|---|---|
| `USysCheck.bas` | API PtrSafe + Long-korjaus | Kommentit suomeksi ✅ |
| `DataToACAD.bas` | Debug.Print-rivit poistettu | Kommentit suomeksi ✅ |
| `GeneralCodes.bas` | Option Explicit | Kommentit suomeksi ✅ |
| `Form_DBUsers.cls` | Me.-etuliitteet | Kommentit suomeksi ✅ |
| `MoottTilaus.bas` | Debug.Print-rivit poistettu | — |
| `Revisiointi.bas` | Debug.Print-rivi poistettu | — |
| Muut lomakkeet (6 kpl) | Option Explicit lisätty | — |
