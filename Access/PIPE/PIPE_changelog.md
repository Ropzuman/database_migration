# Muutosloki: PIPE-tietokanta

**Tiedosto:** `Access/PIPE/` (kaikki .bas- ja .cls-tiedostot)
**Päivämäärä:** 2025-11-12
**Haara:** `PIPE` (eriytetty `main`-haarasta)

---

## Kriittiset muutokset

### 64-bit ja API-korjaukset

- **`Koodit.bas`:** API-määrittelyt (`GetUserNameA`, `GetComputerNameA`) tarkistettu — `#If VBA7 Then / PtrSafe` jo käytössä ✅; `nSize`-parametri `As Long` (ei `LongPtr`) ✅
- **`Form_DBUsers.cls`:** Poistettu kaatava `DBEngine.Workspaces(0).Databases(0).Close`-kutsu — korvattu suoralla `CurrentDb.Name`-kutsuilla (sama korjaus kuin MAINEQ/instru3-moduuleissa)
- **`Form_DBUsers.cls`:** `NetworkName.Value` → `Me.NetworkName.Value` (×2) — `Option Explicit` vaatii `Me.`-etuliitteen lomakkeen kontrolleille

### Kuollut koodi poistettu

- **`Form_DBUsers.cls`:** `iLOF As Integer` ja `iLOF = LOF(iLDBFile)` poistettu — muuttuja hankittu mutta ei koskaan käytetty
- **`Koodit.bas`:** `Dim Doku As String` poistettu `AvaaBlock`-aliohjelmasta — määritelty mutta ei koskaan käytetty
- **`Form_Linkkien vaihto.cls`:** `Dim Taulut() As String` ja `Dim s As Integer` poistettu — kumpikaan ei käytetty

---

## Siivous ja optimointi

### Kommentit suomeksi

Kaikki englanninkieliset kommentit käännetty suomeksi kaikissa kymmenessä PIPE-tiedostossa:

- Moduulitason otsikkolohkot (`' Lomake:`, `' Tarkoitus:`, `' Kuvaus:` jne.)
- Aliohjelma- ja funktio-otsikot (aiempi `' Purpose:`, `' Process:`, `' Notes:` → `' Tarkoitus:`, `' Toiminta:`, `' Huom:`)
- Kaikki inline-kommentit
- MsgBox-virheilmoitukset, kehotukset ja käyttäjäviestit

### Käännetyt MsgBox-viestit (esimerkkejä)

| Vanha (englanniksi) | Uusi (suomeksi) |
|---|---|
| `"Error opening block:"` | `"Virhe blokin avaamisessa:"` |
| `"Error reading pipeline data:"` | `"Virhe putkilinjojen tietojen lukemisessa:"` |
| `"Error updating links:"` | `"Virhe linkkien päivittämisessä:"` |
| `"Cannot activate AutoCAD:"` | `"AutoCAD-ikkunaa ei voitu aktivoida:"` |
| `"Pick pipeline:"` | `"Poimi putkilinja:"` |
| `"No pipeline blocks found."` | `"Putkilinjojen blokkeja ei löytynyt."` |
| `"Last block! Start from first?"` | `"Viimeinen blokki! Aloitetaanko alusta?"` |
| `"Unknown"` (käyttäjänimi) | `"Tuntematon"` |

### Muut parannukset

- **`Form_DBUsers.cls`:** `WhosOn`-funktion otsikko tiivistetty suomenkieliseksi
- **`Koodit.bas`:** `SetStartup`-funktion muuttujakommentit ja rakenne selkeytetty
- **`Form_Linkkien vaihto.cls`:** Pääsilmukan kommentit käännetty ja selkeytetty
- Kaikissa tiedostoissa `' Päivitetty:`-rivit päivitetty suomenkielisillä kuvauksilla

---

## Tiedostoyhteenveto

| Tiedosto | Kriittiset korjaukset | Kommentit | Kuollut koodi |
|---|---|---|---|
| `Form_DBUsers.cls` | ✅ (crash-fix, Me.-etuliite) | ✅ | ✅ |
| `Koodit.bas` | — | ✅ | ✅ |
| `Form_Linkkien vaihto.cls` | — | ✅ | ✅ |
| `Form_TYÖKALUT.cls` | — | ✅ | — |
| `Form_USysFlowPickNo.cls` | — | ✅ | — |
| `Form_USysPipeFromTo.cls` | — | ✅ | — |
| `Form_USysPipeToOther.cls` | — | ✅ | — |
| `Form_frmOpenPIPELINE.cls` | — | ✅ | — |
| `Form_zFunc.cls` | — | ✅ | — |
| `Form_Venttiiliblokkien vaihto.cls` | — | ✅ | — |

---

## Täydennys: Vaihe 3 uusintaskannaus (2025-11-13)

Toinen kierros Vaihe 3 -skannauksesta löysi ja korjasi seuraavat jäljellä olevat englanninkieliset kommentit:

### Form_TYÖKALUT.cls

- Ryhmäotsikot `'===== ... =====` käännetty suomeksi (9 kpl):
  - `Pipeline to Manual Valve` → `Putkilinja käsiventtiileihin`
  - `Pipeline Workflow (Commands 11-13)` → `Putkilinjatyönkulku (Commands 11-13)`
  - `Editor Forms (Commands 18-19, 40)` → `Editorilomakkeet (Commands 18-19, 40)`
  - `Manual Valve Workflow (Commands 21-23)` → `Käsiventtiilityönkulku (Commands 21-23)`
  - `Instrument Valve Workflow (Commands 31-33)` → `Instrumenttiventtiilityönkulku (Commands 31-33)`
  - `Field Instrument Workflow (Commands 41-43)` → `Kenttälaitetyönkulku (Commands 41-43)`
  - `Instrument Loop Workflow (Commands 51-53)` → `Instrumenttisilmukkatyönkulku (Commands 51-53)`
  - `Helper Functions` → `Apufunktiot`
- `' Tarkoitus:` -rivit (22 kpl) käännetty suomeksi

### Form_frmOpenPIPELINE.cls

- `' Controls:` → `' Kontrollit:`

### Form_USysPipeFromTo.cls

- `' Tarkoitus: Pick TO reference from AutoCAD drawing` → `' Tarkoitus: Poimii TO-viittauksen AutoCAD-piirustuksesta`

### Form_zFunc.cls

- `' Tarkoitus: Delete zDetails records that no longer exist in InstrumentIndex` → `' Tarkoitus: Poistaa zDetails-tietueet, joita ei enää löydy InstrumentIndex-taulukosta`

Lopputulos: 0 englanninkielistä kommenttirakennetta kaikissa kymmenessä PIPE-tiedostossa.
