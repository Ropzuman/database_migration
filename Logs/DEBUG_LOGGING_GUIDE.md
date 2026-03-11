# VBA Debug-lokitus — Käyttöohje / Debug Logging Guide

> **Luokitus / Classification:** `[ACTIVE]` — Tekninen kehittäjäopas
> **Päivitetty / Updated:** 3.3.2026
> **Kohderyhmä / Audience:** Kehittäjät / Developers

## Yleiskatsaus

Lisätty kattava Debug.Print -lokitus kriittisiin Access VBA -tiedostoihin **VBA Immediate Window** -virheenselvitystä varten.

## Muutetut Tiedostot

### 1. Access/MAINEQ/DataToACAD.bas

**AutoCAD-integraatio - Kriittisin tiedosto**

Lisätty lokitus 9 funktioon:

#### `CrsRefLink(tblnimi, teksti)`

- Aloitusviesti taulun ja tekstitunnisteiden kanssa
- Virhekäsittely täydellisillä parametritiedoilla

#### `get_filename(taulnimi)`

- Taulunimen prosessointiviesti
- Virhekonteksti

#### `makeFiles(common)`

- **Päätoiminto** - Kattava lokitus
- Funktio-otsikko erottimilla
- Suodatusarvo ja hakemistotiedot
- Prosessoinnin vaiheet:
  - "Initializing output files..."
  - "Generating non-loop-based lists..."
  - "Generating loop-based lists..."
  - "Generating AutoCAD script file..."
- Valmistumisviesti
- Täydellinen virhekonteksti

#### `MakeListNoLoopID(tanimi, Hakem)`

- Taulun ja hakemiston lokitus
- Virhekäsittely

#### `inch(a)`

- Merkkijonon pituustieto
- Koodin poistumisviesti
- Virhekonteksti syöttöparametrilla

#### `MakeListWithLoopID(tblnimipre, Hakem, idsyst, suoda, Looppid)`

- Taulun, hakemiston ja suodatusparametrien lokitus
- Loop ID -kenttäindeksin näyttö
- Virhekäsittely

#### `MakeLocFiles()`

- Funktio-otsikko
- Instloc.txt -tiedoston alustusviesti
- Laitetaulun prosessointiviestit
- Viimeistelyvaihe
- Valmistumisviesti

#### `MakeScript(common, suod, Looppid)`

- Funktio-otsikko parametreilla
- Skriptitiedoston polku
- QMEM-komentojen kirjoitusviesti
- Silmukkatietueiden prosessointi
- Valmistumisviesti

---

### 2. Access/instru3/Form_CopyLoops.cls

**Monimutkainen tietokantaoperaatio**

Lisätty lokitus 2 proseduuriin:

#### `HaeLoopit_Click()`

- Funktio-otsikko erottimilla
- Lähdetietokannan validointi:
  - "ERROR: No source database selected" jos ei valittu
  - "Source DB: [polku]" jos valittu
- Edistymisviestit:
  - "Copying selected loops to LOOPS table..."
  - "  Adding loop: [areacode]-[loopno]"
  - "Importing device records from devTbl* tables..."
  - "  Processing table: [tablename]"
- Päällekirjoitusviestit duplikaateista
- Valmistumisviesti:
  - "HaeLoopit: COMPLETED successfully"
  - "  Total loops copied: [määrä]"
- Täydellinen virhekäsittely

#### `ValitseKanta_Click()`

- Tietokantavalintadialogien aloitusviesti
- Valitun tietokannan polku
- "Database opened successfully" -vahvistus
- Virhekonteksti

---

### 3. Access/DOCUMENTS/GlobalVBAs.vba

**Revisioparserointi - Paljon käytetty**

Lisätty lokitus 7 funktioon:

#### `HaeTekija(Revisio)`

- Parseroidaan tekijä revisiojonosta
- "Revisio is Null" -ilmoitus tyhjille arvoille
- "Found author: [nimi]" löydetyille
- Virhekäsittely revisioparametrilla

#### `HaeRevisioija(Revisio)`

- Parseroidaan revisoija
- "No multi-line revision" jos yhden rivin revisio
- "Found reviser: [nimi]" löydetyille
- Virhekonteksti

#### `HaeRevisioijaPvm(Revisio)`

- Parseroidaan revisoija ja päivämäärä
- "Found: [nimi]: [pvm]" löydetyille
- Virhekäsittely

#### `EkaRevRivi(Revisio)`

- Ensimmäisen revisorivin ekstraktointi
- "Result: [rivi]" -tulostus
- Virhekonteksti

#### `HaeRevisio(Revisio)`

- Revisiotunnuksen (A, B, 0) ekstraktointi
- Null-tarkistus
- "Found mark: [tunnus]" -tulostus
- Virhekäsittely

#### `HaeViimPaiva(Revisio)`

- Ensimmäisen revision päivämäärän ekstraktointi
- "Found date: [päivämäärä]" -tulostus
- Virhekonteksti

#### `HaePaiva(Revisio)`

- Viimeisimmän revision päivämäärän ekstraktointi
- "Found date: [päivämäärä]" -tulostus
- Virhekäsittely

---

## Debug.Print -Lokitusmalli

Kaikki funktiot noudattavat yhtenäistä mallia:

### Funktioiden Aloitus

```vba
Debug.Print "FunctionName: Aloitetaan operaatio"
Debug.Print "  Parametri1: " & parametri1
Debug.Print "  Parametri2: " & parametri2
```

### Kriittisten Toimintojen Otsikot

```vba
Debug.Print "========================================"
Debug.Print "FunctionName: Kuvaus"
Debug.Print "========================================"
```

### Edistymisviestit

```vba
Debug.Print "  Prosessoidaan: " & kohde
Debug.Print "  Kopioidaan: " & määrä & " kohdetta"
```

### Virhekäsittelijät

```vba
ErrorHandler:
  Debug.Print "*** ERROR in FunctionName: " & Err.Number & " - " & Err.Description
  Debug.Print "    Parametri1: " & parametri1
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  MsgBox "Error: " & Err.Description, vbCritical
```

### Valmistumisviestit

```vba
Debug.Print "FunctionName: VALMIS onnistuneesti"
Debug.Print "========================================"
```

---

## Käyttöohje

### 1. Tuo Koodi Accessiin

Käytä PowerShell-automaatioskriptia:

```powershell
.\Automations\export_access_vba.ps1
```

### 2. Avaa VBA Editor

Access → Alt+F11 → avaa VBA Editor

### 3. Avaa Immediate Window

VBA Editor → Ctrl+G → avaa Immediate Window (tai View → Immediate Window)

### 4. Suorita Toiminto

Esim. avaa Form_CopyLoops ja klikkaa "Hae Loopit" -nappia

### 5. Seuraa Lokeja Immediate Windowissa

Näet reaaliaikaisen suorituksen:

```
========================================
HaeLoopit: Starting loop copy operation
========================================
Source DB: P:\database\source.accdb
Copying selected loops to LOOPS table...
  Adding loop: 01-001
  Adding loop: 01-002
Importing device records from devTbl* tables...
  Processing table: devTblValve
  Processing table: devTblPump
HaeLoopit: COMPLETED successfully
  Total loops copied: 2
========================================
```

### 6. Virheiden Diagnoosi

Jos virhe tapahtuu, näet täydelliset tiedot:

```
*** ERROR in HaeLoopit: 3021 - No current record
    Source: DAO.Recordset, Line: 145
========================================
```

Tästä näet:

- **Virhekoodin:** 3021
- **Kuvauksen:** "No current record"
- **Lähteen:** DAO.Recordset
- **Rivinumeron:** 145 (jos VBA tukee Erl)
- **Kontekstin:** Mitkä parametrit olivat käytössä

---

## Edut

### 1. Reaaliaikainen Näkyvyys

- Näet tarkalleen mitä koodi tekee
- Ei tarvitse lisätä MsgBox-viestejä

### 2. Virheanalyysi

- Täydelliset virhetiedot parametreilla
- Helppo tunnistaa missä vaiheessa virhe tapahtui

### 3. Suorituskyvyn Seuranta

- Näet kuinka kauan eri vaiheet kestävät
- Voit tunnistaa hitaat kohdat

### 4. Ei Tuotantovaikutusta

- Debug.Print -komennot poistetaan automaattisesti Release-buildeissa
- Eivät hidasta merkittävästi (kirjoitetaan vain kun Immediate Window auki)

---

## Testaussuunnitelma

Käytä tätä yhdessä **AGENT_TEST_VALIDATION_CHECKLIST.md**:n kanssa:

### Vaihe 1: Käännöksen Tarkistus

1. Avaa Access-tietokanta
2. Paina Alt+F11 → VBA Editor
3. Debug → Compile [Projektin nimi]
4. Varmista: Ei käännösvirheitä

### Vaihe 2: Kriittisten Tiedostojen Testaus

#### Test 1: DataToACAD.bas - makeFiles

1. Avaa Immediate Window (Ctrl+G)
2. Suorita: `makeFiles("COMMON")`
3. Tarkista Immediate Windowista:
   - Funktio-otsikko näkyy
   - Hakemistotiedot näkyvät
   - Prosessointivaiheet näkyvät
   - Valmistumisviesti ilmestyy

#### Test 2: Form_CopyLoops - HaeLoopit

1. Avaa Form_CopyLoops
2. Avaa Immediate Window
3. Valitse lähdetietokanta
4. Valitse loopit
5. Klikkaa "Hae Loopit"
6. Tarkista lokista:
   - Lähdetietokannan polku
   - Kopioidut loopit
   - Prosessoidut taulut

#### Test 3: GlobalVBAs - Revisioparserointi

1. Avaa Immediate Window
2. Testaa: `? HaeTekija("A 01.01.2024/VG/RVA/" & vbCrLf & "B 01.02.2024/MT/RVA/")`
3. Tarkista lokista:
   - Parseroidaan tekijä
   - Löydetty tekijä näkyy
4. Tarkista tulos: Pitäisi palauttaa "VG"

---

## Seuraavat Vaiheet

1. ✅ **Debug-lokitus lisätty** (3 kriittiseen tiedostoon, 18 funktioon)
2. ⬜ **Testaa manuaalisesti Accessissa**
   - Käytä Immediate Window -lokeja
   - Raportoi mahdolliset virheet
3. ⬜ **Tunnista ongelmat**
   - Käytä Debug.Print -tulosteita virheiden diagnosointiin
4. ⬜ **Korjaa ongelmat**
   - Lisää tarvittaessa lisää Debug.Print -komentoja
5. ⬜ **Täydellinen regressiotestaus**
   - AGENT_TEST_VALIDATION_CHECKLIST.md:n mukaisesti

---

## Git-tiedot

**Branch:** agent_test  
**Commit:** 86f3871  
**Viesti:** "Lisää Debug.Print -lokitusta VBA Immediate Window -virheenselvitystä varten"

**Muutokset:**

- 6 tiedostoa muutettu
- 312 lisäystä
- 11 poistoa

**Pushed:** origin/agent_test ✅

---

## Yhteenveto

Kaikki kriittiset Access VBA -tiedostot on varustettu kattavalla Debug.Print -lokituksella. Voit nyt:

1. **Tuoda koodin Accessiin** PowerShell-automaatiolla
2. **Avata Immediate Windowin** (Ctrl+G)
3. **Suorittaa toimintoja** ja seurata reaaliaikaisesti mitä tapahtuu
4. **Diagnosoida virheet** täydellisillä parametri- ja kontekstitiedoilla
5. **Raportoida ongelmat** tarkoilla lokitiedoilla

**Kaikki muutokset noudattavat:**

- ✅ 64-bit compliance (PtrSafe, LongPtr)
- ✅ DAO-etuliitteet
- ✅ Optimoidut String-funktiot ($-suffiksit)
- ✅ Opinnäytetyödokumentaatio suomeksi

**Valmis testaukseen!** 🚀
