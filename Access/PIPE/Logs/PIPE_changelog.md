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

---

## Code Review -kierros: Vakaus- ja tietoturvakorjaukset (2026-03-07)

### Kriittiset korjaukset

#### Form_zFunc.cls — SQL-injektiovulneraabiliteetti poistettu

- **Command5_Click:** Korvattu vaarallinen SQL-merkkijonoketjutus (`L & acode & L`) parametrisoidulla `DAO.QueryDef`-kyselyllä (`[pArea]`, `[pLoop]`, `[pSymb]`, `[pSuffix]`)
- Poistettu käyttämättömät muuttujat: `Dim tbl2`, `Dim qry`, `Dim ssSQL`, `Dim L` — kuollut koodi siivottu
- Virheenkäsittelijä päivitetty: `Set tbl2 = Nothing` → `Set qdf = Nothing`

#### Form_TYÖKALUT.cls — Kriittinen logiikkabugi + COM-muistivuodot

- **Command43_Click:** Korjattu kriittinen bugi `= vbYesNo` → `= vbYes` — kenttälaitteiden päivitys ei ennen koskaan ajanut (vertailu palauttaa aina `False`)
- **Command13, 23, 33, 43, 53:** Lisätty `vbExclamation + vbDefaultButton2` kaikkiin tuhoisia tietokantapäivityksiä vahvistaviin dialogeihin — "Ei" on nyt oletuspainike vahingon estämiseksi
- **LueTiedot, LueTiedotByAttribute, LueTiedotByBlockAndAttribute — ErrorHandler:** Lisätty COM-objektien siivous (`Joukko.Delete`, `Set Joukko = Nothing`, `Set oDOC = Nothing`, `Set oACAD = Nothing`) — muistivuoto estettiin suurten DWG-eräajojen aikana

#### Koodit.bas — API-parametrien tarkkuus

- `api_GetUserName` ja `api_GetComputerName`: Lisätty eksplisiittinen `ByRef nSize As Long` molempiin haruihin (`#If VBA7 Then` ja `#Else`) — poistaa implisiittisyyden, joka voi aiheuttaa muistivirheitä 64-bit-ympäristössä

#### Form_USysPipeToOther.cls — Luettavuus ja Me.-etuliite

- **OpenPipeline:** Pitkä `If/Or`-ehtolauseke (`PIPELINE_F Or PIPELINE Or ...`) refaktoroitu `Select Case` -rakenteeksi — DRY-periaate, helpompi laajentaa uusilla blokkityypeillä
- **TPipeline.Value** → **Me.TPipeline.Value** kolmessa kohdassa (OpenPipeline, AvaaBlokki, BFind_Click) — eksplisiittinen lomakeviittaus, parempi Option Explicit -yhteensopivuus

#### Form_DBUsers.cls — Virheenkäsittely ja vanhentuneet komennot

- **WhosOn — Err_WhosOn:** Tallennetaan `Err.Number`/`Err.Description` muuttujiin ennen `Close`-kutsua — alkuperäinen virhetieto ei enää katoa `On Error Resume Next` -kontekstin takia; `If Err = 68` toimi aina väärin (0 = 68 on False)
- **WhosOn:** Lisätty `iLDBFile > 0` -vartiointi (`If iLDBFile > 0 Then Close #iLDBFile`) — sulkee tiedostokahvan turvallisesti vain jos se on avattu
- **sPath-muodostus:** `+`-operaattori → `&`-operaattori merkkijonoketjutuksessa (VBA-standardikäytäntö)
- **Command27_Click:** `net send` → `msg.exe` — vanha Windows-komento poistettu käytöstä Vistasta alkaen; päivitetty nykyiseen vastineeseen

### Päivitetty tiedostoyhteenveto

| Tiedosto | Korjaukset 2026-03-07 |
|---|---|
| `Form_zFunc.cls` | SQL-injektio, kuollut koodi |
| `Form_TYÖKALUT.cls` | Kriittinen bugi, COM-vuodot, MsgBox-oletukset |
| `Koodit.bas` | API ByRef-tarkkuus |
| `Form_USysPipeToOther.cls` | Select Case, Me.-etuliite |
| `Form_DBUsers.cls` | Err-tallennus, tiedostokahva, net send |

---

## Code Review -päivitys: Toinen kierros (2026-03-07)

### Form_DBUsers.cls — VBA And-operaattorin oikosuljemattomuus

- **WhosOn — bMach/bUser-silmukat:** `While .bMach(i) <> 0 And i <= 32` → `Do While i <= 32` + `If .bMach(i) = 0 Then Exit Do`
  - VBA arvioi aina **molemmat** `And`-haarat — jos `i = 33`, ohjelma yrittää lukea `.bMach(33)` ennen rajakieltoa → "Subscript out of range" (Error 9)
  - Sama korjaus tehty sekä tietokoneen nimen (`bMach`) että käyttäjänimen (`bUser`) silmukkaan

### Form_TYÖKALUT.cls — Valintajoukkojen nimitörmäykset

- **PoimiJoukko, PoimiJoukkoByAttribute, PoimiJoukkoByBlockAndAttributeValue:** `For i = 0 To ... SelectionSets.Count - 1` -poisto → `On Error Resume Next` / `oDOC.SelectionSets.Item("APUPICK").Delete` / `On Error GoTo ErrorHandler`
  - Jos edellinen ajo kaatui ennen joukon tuhoamista, `"APUPICK"` jäi AutoCADin muistiin → `.Add("APUPICK")` heitti virheen ja koko skripti pysähtyi
  - `.Item().Delete` on suorempi ja kutsu epäonnistuu hiljaisesti (On Error Resume Next) jos joukkoa ei enää ole
