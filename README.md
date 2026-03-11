# Tietokantamigraatio: 32-bit → 64-bit Office VBA
> Suunnittelijoiden tietokantajärjestelmä päivitetty 64-bittiseksi M365-yhteensopivaksi.
> *Design system database upgrade to 64-bit M365 compatibility.*

---

## 📌 Projektin tila / Project Status

> **Versio / Version:** 2.0 — 64-bit M365 Compatible
> **Viimeksi päivitetty / Last updated:** 11.3.2026
> **Yhteensopivuus / Compatibility:** Microsoft 365 (64-bit), Excel, Access, AutoCAD 2019+
> **Tila / Status:** ✅ Migraatio valmis — kaikki moduulit testattu

---

## 🗂️ Sisällysluettelo / Table of Contents

1. [Yleiskatsaus / Overview](#yleiskatsaus--overview)
2. [Mitä muuttui / What Changed](#mitä-muuttui--what-changed)
3. [Järjestelmävaatimukset / System Requirements](#järjestelmävaatimukset--system-requirements)
4. [Tietokantatyökalut / Database Tools](#tietokantatyökalut--database-tools)
5. [AutoCAD-integraatio / AutoCAD Integration](#autocad-integraatio--autocad-integration)
6. [Automaatiot / Automations](#automaatiot--automations)
7. [Vianmääritys / Troubleshooting](#vianmääritys--troubleshooting)
8. [Kehittäjille / For Developers](#kehittäjille--for-developers)
9. [Muutoshistoria / Changelog](#muutoshistoria--changelog)

---

## 🔍 Yleiskatsaus / Overview

Tämä järjestelmä on suunnittelijoille ja projektipäälliköille tarkoitettu tietokantapohjainen työkalu, joka yhdistää **MS Access -tietokannan**, **Excel-työkirjat** ja **AutoCAD 2019** -piirustukset yhdeksi kokonaisuudeksi.

*This is a database-driven toolset for designers and project managers that connects an MS Access database, Excel workbooks, and AutoCAD 2019 drawings into a unified workflow.*

**Järjestelmä koostuu kolmesta osasta / The system has three parts:**

| Osa / Part | Mitä tekee / What it does |
|---|---|
| **Access-tietokanta** | Säilyttää laitteiden, piirien ja dokumenttien tiedot / Stores equipment, circuit, and document data |
| **Excel-kyselyt** | Hakee tietoja kannasta ja tuottaa listat ja tulosteet / Fetches data and produces lists and printouts |
| **AutoCAD-integraatio** | Lukee ja kirjoittaa lohkoattribuutteja AutoCAD-piirustuksiin / Reads and writes block attributes in AutoCAD drawings |

---

## 🔄 Mitä muuttui / What Changed

Järjestelmä päivitettiin toimimaan nykyaikaisessa **64-bittisessä Microsoft 365** -ympäristössä. Aiempi versio toimi vain 32-bittisessä Officessa.

*The system was updated to work in modern 64-bit Microsoft 365. The previous version only worked in 32-bit Office.*

**Käyttäjälle näkyvät muutokset / Changes visible to users:**
- ✅ Kaikki toiminnot toimivat Microsoft 365 (64-bit) -ympäristössä
- ✅ Yhteydet tietokantaan ovat nopeampia ja luotettavampia
- ✅ Automaatioskriptit toimivat 64-bit PowerShellissä

**Tekniset muutokset (kehittäjille) / Technical changes (for developers):**

<details>
<summary>▶ Näytä tekniset muutokset / Show technical details</summary>

- Kaikki `Declare`-lauseet päivitetty `PtrSafe`-avainsanalla ja `LongPtr`-tyypeillä
- Tietokanta-ajuri vaihdettu: `Microsoft.Jet.OLEDB.4.0` → `Microsoft.ACE.OLEDB.12.0`
- `Nz()`-funktio korvattu `IIf(IsNull(), 0, Value)` -rakenteella (Excel VBA -yhteensopivuus)
- Kaikki koodikommentit suomeksi Ä/Ö-kirjaimia käyttäen
- `ScreenUpdating` ja `Calculation`-suojaukset lisätty virheenkäsittelijöineen
- Ajuri-fallback: 16.0 → 15.0 → 12.0 eri Office-versioiden tukemiseksi
- Kaikki laskurimuuttujat muutettu `Long`-tyyppisiksi 64-bit-yhteensopivuuden vuoksi

Katso täydelliset muutokset: `Logs/CHANGELOG_64bit_and_perf.md`

</details>

---

## 💻 Järjestelmävaatimukset / System Requirements

| Komponentti / Component | Vaatimus / Requirement |
|---|---|
| Office | Microsoft 365 (64-bit) |
| Access | Microsoft Access (sisältyy M365:een / included in M365) |
| Excel | Microsoft Excel (sisältyy M365:een / included in M365) |
| AutoCAD | AutoCAD 2019 tai uudempi / or newer |
| Windows | Windows 10/11 (64-bit) |
| PowerShell | 5.1+ (64-bit) — automaatioita varten / for automations |

> ⚠️ **Tärkeää / Important:** Varmista että Office on **64-bittinen** versio. Tarkista: Excel → Tiedosto → Tili → Tietoja Excelistä. Ylhäällä pitää näkyä "64-bit".
> *Make sure Office is the **64-bit** version. Check: Excel → File → Account → About Excel. It must say "64-bit".*

---

## 🛠️ Tietokantatyökalut / Database Tools

### Kytkentälista / Connection List

Kytkentälista-työkalu hakee tietoja Access-kannasta ja tuottaa tulosteen Excel-pohjasta.

*The Kytkentälista tool fetches data from the Access database and generates a printout from an Excel template.*

**Kolme päätoimintoa / Three main functions:**

#### 1. Hae tiedot (HaeData) / Get Data

- Hakee tiedot Access-tietokannasta Excel-arkeille DB1 ja DB2
- Tarkistaa tietokantatiedoston olemassaolon ennen yhteyden avaamista
- Käsittelee yhteysvirheet selkokielisillä virheilmoituksilla

*Fetches data from the Access database to Excel sheets DB1 and DB2. Validates the database file before connecting.*

#### 2. Aja tarkistus (Checkout) / Run Check

- Tarkistaa TEMPLATE-otsikot DB1-dataa vasten
- Hakee dokumentin metatiedot DB2:sta
- Täyttää Info-arkin tiedoilla
- Raportoi puuttuvat otsikot ERRORS-arkille

*Validates TEMPLATE headers against DB1 data, fetches document metadata from DB2, fills the Info sheet, reports missing headers to ERRORS sheet.*

#### 3. Luo tuloste (GenPrintout) / Generate Printout

- Luo uuden työkirjan TEMPLATE-pohjasta
- Täyttää DB1-datan ja dokumentin metatiedot
- Tuottaa tulostusvalmiit dokumentit

*Creates a new workbook from the TEMPLATE, populates it with DB1 data and document metadata, produces print-ready output.*

**Tietokantayhteyden asetukset / Database connection settings:**

- DB1 hakee piirikaavio- ja IO-terminaalidatan (`Circuit_Diagrams_IO_Terminals`)
- DB2 hakee dokumenttimetatiedot (`DOCUMENTS`-taulu tai `_qryForExcel`-kysely)
- Hakupolku: DB2-sarake `WorkPath` määrittää oletustallennushakemiston
- Tiedostonimi: DB2-sarake `File` — puuttuessa käytetään pohjan nimeä + `.xlsx`

---

## 🗺️ AutoCAD-integraatio / AutoCAD Integration

AcadDATA-työkalu lukee lohkoattribuutteja AutoCAD-piirustuksista ja kirjoittaa muutokset takaisin.

*The AcadDATA tool reads block attributes from AutoCAD drawings and writes changes back.*

### Tuo tiedot (TuoDATA) / Import Data

Ajettavissa makroina: `TuoDATA_All` (kaikki lohkot) tai `TuoDATA_Selected` (edellinen AutoCAD-valinta).

**Start-arkin asetukset / Start sheet settings:**

| Kenttä / Field | Kuvaus / Description |
|---|---|
| **D7** | Lohkojen nimet pilkulla erotettuna. `*` = kaikki lohkot. / Block names, comma-separated. `*` = all blocks. |
| **D5** | Entiteettityyppi: `"Blokit"`, `"Tekstit"` tai `"Blokit ja tekstit"` / Entity scope |
| **Nykyinen** (checkbox) | Tuo nykyisestä avoimesta AutoCAD-piirustuksesta / Import from currently active drawing |
| **Lista** (checkbox) | Tuo TIEDLISTA-arkin tiedostoluettelosta / Import from file list on TIEDLISTA sheet |

Ulostulosarakkeet: PATH, DWG, BLOCK, HANDLE, XCord, YCord, Layer + yksi sarake per attribuuttitagi.

### Vie tiedot (VieDATA) / Export Data

Kirjoittaa muokatut attribuuttiarvot takaisin lohkoihin HANDLE-tunnisteen avulla.

*Writes edited attribute values back to blocks using the HANDLE identifier.*

- Kutsuu `oBlock.Update` jokaisen lohkon jälkeen — muutokset näkyvät heti ruudulla
- Kutsuu `oDOC.Regen 1` lopussa kaikkien näkymien päivittämiseksi
- Käyttää `HeaderMap`-sanakirjaa nopeaan TAG→sarake-hakuun

### Kaksoisnapsautusnavigaatio / Double-click Navigation

Dataarkilta kaksoisklik riville zoomaa AutoCAD näyttämään kyseisen entiteetin.

*Double-clicking a row on the data sheet zooms AutoCAD to that entity.*

<details>
<summary>▶ Kehittäjätiedot / Developer details</summary>

- Dynaamisten lohkojen tunnistus `EffectiveName`-ominaisuuden kautta (anonyymit sisäiset nimet käsitellään oikein)
- DXF-tyyppisuodatin: `FilterType(0)=0` / `FilterData(0)="INSERT"` (luotettava AutoCAD 2019 late binding -ympäristössä)
- Debug-jäljitys: `Public Const DEBUG_TRACE As Boolean` tiedostossa `Excel/Moduulit/AcadDATA/Koodit.bas`
- Lisätiedot: `Logs/ACADDATA_DEVELOPER_NOTES.md` ✅

</details>

---

## ⚙️ Automaatiot / Automations

Automaatioskriptit päivittävät VBA-moduulit tiedostoihin ilman manuaalista kopiointia. Kaikki yksityiskohtaiset ohjeet: [`Automations/README.md`](Automations/README.md).

*Automation scripts update VBA modules in files without manual copy-pasting. Full instructions: [`Automations/README.md`](Automations/README.md).*

| Skripti / Script | Käyttötarkoitus / Purpose |
|---|---|
| `Automations/Access_automaatio.ps1` | Päivittää VBA-moduulit .accdb-tietokantaan / Updates VBA modules in .accdb database |
| `Automations/Access_automaatio_batch.ps1` | Eräajo useille tietokannoille / Batch update for multiple databases |
| `Automations/Excel_automaatio.ps1` | Päivittää VBA-moduulit .xlsm-työkirjoihin / Updates VBA modules in .xlsm workbooks |

**Esivaatimukset / Prerequisites:**
1. Aja 64-bittisessä PowerShellissä (x64) — skripti tarkistaa tämän automaattisesti
2. Microsoft Access/Excel asennettuna
3. Trust Center: ota käyttöön "Luota VBA-projektin objektimalliin" / "Trust access to the VBA project object model"

> ⚠️ **Tärkeää / Important:** Skriptit käyttävät **suoraa koodinkorvausta** eikä `VBComponents.Import()` -metodia. Tämä estää näkymättömän metatietovioittumisen. Katso `Automations/Logs/Access_automaatio_changelog.md`.

---

## 🔧 Vianmääritys / Troubleshooting

### Yleisimmät ongelmat / Most common issues

| Ongelma / Problem | Syy / Cause | Ratkaisu / Fix |
|---|---|---|
| "Tietokantaa ei löydy" | Tiedostopolku väärä tai tiedosto siirretty | Tarkista polku faceplate-kentässä |
| Excel jäätyy Checkout-ajon aikana | Vanha versio (korjattu v2.0) | Varmista käytössä on versio 2.0 |
| "Provider not found" -virhe | Väärä Office-bittisyys tai puuttuva Access-ajuri | Tarkista Office on 64-bit |
| AutoCAD ei vastaa tuonnin aikana | Suuri piirustus tai verkkoasema | Odota — älä sulje AutoCADia |
| PowerShell-skripti ei toimi | 32-bit PowerShell käytössä | Avaa "Windows PowerShell" (ei x86-versiota) |

### Testiohjeet kehittäjille / Test instructions for developers

1. **Access:** Debug → Compile VBA-editorissa, testaa `SniffUser`, `CustomMessage`
2. **Excel:** Klikkaa "Hae tiedot" → tarkista DB1/DB2 täyttyvät ilman virheitä
3. **Excel:** Klikkaa "Aja tarkistus" → tarkista Info-arkki täyttyy, ei ERRORS-merkintöjä
4. **Excel:** Klikkaa "Luo tuloste" → tarkista uusi työkirja luodaan oikein

---

## 👨‍💻 Kehittäjille / For Developers

### Tiedostorakenne / File Structure

```
Access/              — Access VBA -moduulit / Access VBA modules
Excel/Moduulit/
  AcadDATA/          — AutoCAD-integraatiomoduuli / AutoCAD integration module
  Listojen kyselyt/  — Kytkentälista-kyselymoduulit / Connection list query modules
AutoCAD/             — AutoCAD DVB-projektit / AutoCAD DVB projects
Automations/         — PowerShell-automaatioskriptit / PowerShell automation scripts
  Apuskriptit/       — Skannarit ja apuskriptit / Scanners and helper scripts
Logs/                — Muutoslokit ja analyysit / Changelogs and analysis docs
_archive/            — Arkistoidut vanhat dokumentit / Archived old documents
```

### Tekniset viitedokumentit / Technical Reference Documents

- `Logs/ACADDATA_DEVELOPER_NOTES.md` — AcadDATA-moduulin kehittäjämuistiinpanot
- `Logs/FACEPLATE_FEATURES.md` — Faceplate-lomakkeen ominaisuudet
- `Logs/LISTOJEN_KYSELYT_REFACTORING.md` — Kysely-moduulin refaktorointilogiikka
- `Logs/REFACTORING_DOCUMENTATION.md` — Projektitason refaktorointidokumentti
- `Logs/REFACTORING_DOCUMENTATION.md` — Projektitason refaktorointikuvaus (sarakekartoitus dokumentoitu sisällä)

### Tietokanta-ajuri / Database Driver

Kaikki yhteydet käyttävät ajuria `Microsoft.ACE.OLEDB.12.0` automaattisella fallback-logiikalla:
versio 16.0 → 15.0 → 12.0.

---

## 📋 Muutoshistoria / Changelog

Täydelliset muutoslokit / Full changelogs:

- **Projektin päämuutosloki:** `Logs/CHANGELOG_64bit_and_perf.md`
- **Viimeisin vaihe:** `Logs/PHASE3_PHASE4_WORKSPACE_CHANGELOG_2026-03-09.md`
- **Access MAINEQ:** `Access/MAINEQ/Logs/MAINEQ_changelog.md`
- **Access PIPE:** `Access/PIPE/Logs/PIPE_changelog.md`
- **Access LoopCircuit:** `Access/LoopCircuit/Logs/LoopCircuit_changelog.md`
- **Access FunctionDiagrams:** `Access/FunctionDiagrams/Logs/FunctionDiagrams_changelog.md`
- **Access DOCUMENTS:** `Access/DOCUMENTS/Logs/DOCUMENTS_changelog.md`
- **Excel AcadDATA:** `Excel/Moduulit/AcadDATA/Logs/AcadDATA_changelog.md`
- **Excel Listojen kyselyt:** `Excel/Moduulit/Listojen kyselyt/Logs/ListojenKyselyt_changelog.md`
- **AutoCAD:** `AutoCAD/exported/PHASE3_changelog.md`
- **Automaatiot:** `Automations/Logs/Access_automaatio_changelog.md`, `Automations/Logs/Excel_automaatio_changelog.md`