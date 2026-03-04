# Muutosloki — MAINEQ-moduuli

**Tiedosto:** `Access/MAINEQ/` (DataToACAD.bas, GeneralCodes.bas, USysCheck.bas, Form_DBUsers.cls, ja muut)
**Päivämäärä:** 2025-11-11 / 2025-11-12 / 2026-03-04
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
  - Ohjauslogiikan kommentit (tiedostojen alustus, sulkumerkit, siivous, virheenkäsittely)
  - Silmukkakommentit (`' tables` → `' Käydään läpi kaikki taulukot` jne.)
- Englanninkieliset `MsgBox`-viestit käännetty suomeksi

### Kommenttien suomennus — GeneralCodes.bas

- Moduuliotsikko käännetty suomeksi
- Julkisten muuttujien otsikko käännetty: `' Public Variables for Revision Tracking'` → suomeksi
- Kaikkien funktioiden otsikkolohkot käännetty: `IsLoaded`, `HaeViimPaiva`, `Optiot`, `Positiot`, `Vaihekulma`, `MotKaapUh`, `LisaaNo`
- Mukautetun `Replace()`-funktion poisto-huomio tiivistetty suomeksi (4 riviä)
- Muuttujan esittelykommentit käännetty: `Original length`, `Numeric value`, `Loop counter` → suomeksi
- Kirjoitusvirheet korjattu: `vaihekulmna` → `vaihekulman`, `käytjölle` → `käytölle`

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

- Poistettu `' Updated 2025-11-11: Added DAO prefix` -tyyppiset väliaikaismuistiinpanot `Form_DBUsers.cls`- ja `USysCheck.bas`-tiedostoista (3 kpl / tiedosto)

---

## AutoExec-korjaus (2026-03-04)

### Puuttuva SniffUser()-funktio — USysCheck.bas

- **Ongelma:** AutoExec-makron `RunCode`-toiminto kutsui `=SniffUser()`, jota ei ollut MAINEQ-projektissa lainkaan → Access antoi virheen 2425 käynnistyksen yhteydessä
- **Korjaus:** `SniffUser()`-funktio lisätty `USysCheck.bas`:iin
  - Hakee verkkokäyttäjänimen olemassa olevalla `wu_GetUserName`-API:lla
  - Hakee tietokoneen nimen `Environ("COMPUTERNAME")`:lla (ei vaadi lisä-API:a)
  - Kirjoittaa kirjautumistietueen `UsysUsers`-tauluun (kentät: NetworkUser, DBUser, ComputerName, LoginTime)
  - Virheet käsitellään hiljaisesti — ei keskeytä sovelluksen käynnistystä

---

## Tiedostoyhteenveto

| Tiedosto | Vaihe 1–2 | Vaihe 3 | AutoExec-korjaus |
|---|---|---|---|
| `USysCheck.bas` | API PtrSafe + Long-korjaus | Kommentit suomeksi ✅ | `SniffUser()` lisätty ✅ |
| `DataToACAD.bas` | Debug.Print-rivit poistettu | Kommentit suomeksi ✅ | — |
| `GeneralCodes.bas` | Option Explicit | Kommentit suomeksi ✅ | — |
| `Form_DBUsers.cls` | Me.-etuliitteet | Kommentit suomeksi ✅ | — |
| `MoottTilaus.bas` | Debug.Print-rivit poistettu | — | — |
| `Revisiointi.bas` | Debug.Print-rivi poistettu | — | — |
| Muut lomakkeet (6 kpl) | Option Explicit lisätty | — | — |

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
