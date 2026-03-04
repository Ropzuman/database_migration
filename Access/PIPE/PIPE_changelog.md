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
