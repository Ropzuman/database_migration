# Muutosloki — Access/DOCUMENTS

**Tiedosto:** `Access/DOCUMENTS/` (koko kansio)
**Päivämäärä:** 2026-03-19
**Haara:** `main`

---

## Muutokset — regressiokorjaus ja refaktorointi (2026-03-19)

### Kriittinen regressiokorjaus: käyttäjänimi palautui aina Unknown

- `Form_USysNewDistribution.cls` / `GetNetworkUserName`: palautuslogiikka korjattu käyttämään jaettua `NetworkUserName()`-funktiota (`ForDocuments.vba`) ja fallbackina `"Unknown"` vain virhetilanteessa.
- `Form_USysReserve.cls` / `GetNetworkUserName`: sama korjaus kuin yllä.
- Vaikutus: `SelectedBy`-kenttään ja varausleimoihin tallentuu taas oikea käyttäjänimi.

### Arkkitehtuurimuutos: globaalit välimuuttujat -> TempVars/OpenArgs

- Lomakkeiden välinen tiedonsiirto on siirretty pois globaaleista `Common`/`CurRecord`-tyyppisistä arvoista kohti nimettyjä `TempVars`-avaimia (esim. `DOC_*`).
- `OpenArgs`-välitys lisätty tunnisteiden ja yksittäisten tekstiarvojen siirtoon (`USysEditDistribution`, `USysOpenFile`, `USysShowCommon`).
- Vaikutus: pienempi riski ristikkäisille sivuvaikutuksille samanaikaisissa formiavausketjuissa.

### Yhteensopivuusmuutokset ForDocuments-moduuliin

- `ForDocuments.vba`: API-julistukset suojattu `#If VBA7 Then`-ehdoilla 32/64-bit -haarautukseen.
- Kansiovalinta (`ValitseHakem`) käyttää `Application.FileDialog`-pohjaista toteutusta molemmissa haaroissa.
- Lisätty yhteisiä asetus-/hakufunktioita revisio- ja common-tekstin välitykseen.

### Tarkennettu käynnistysvakaus

- `GlobalVBAs.vba`: `SetStartup()` palauttaa Boolean-arvon ja tarkistaa `UsysUsers`-taulun olemassaolon ennen lokikirjoitusta.
- Tarkoitus: AutoExec-käynnistys ei kaadu puuttuvaan tai lukittuun lokitauluun.

### Nopea symbolitarkistus (2026-03-19)

- Ajettiin kansiotason staattinen tarkistus `*.cls/*.vba/*.bas`-tiedostoihin.
- Tulos: `Option Explicit` puuttuvia tiedostoja **0 kpl**.
- Duplikaattiproseduurit: löydökset kohdistuivat vain `ForDocuments.vba`-tiedoston `#If VBA7 Then / #Else` -haarojen rinnakkaisiin määrittelyihin (`ValitseHakem`, `DummyFunc`, `BrowseCallbackProc`), jotka ovat tässä toteutuksessa odotettuja eikä niistä synny aktiivisessa haarassa käännösvirhettä.
- Johtopäätös: tarkistuksessa ei löytynyt uusia toimintakelpoisia `Ambiguous name` / `Variable not defined` -blokkeritasoisia löydöksiä.

---

## Kriittiset muutokset — Code Review -kierros 3 (2026-03-09)

### Monikäyttäjäongelma korjattu — Form_USysNewDistribution.cls

Talisoja havaittiin, että `Form_Load`-tapahtuman `UPDATE USysRecipients SET [To Distribution] = No` nollaa **kaikkien** käyttäjien valinnat. Kahdessa samanaikaisessa istunnossa toisen käyttäjän valinnat tuhoutuivat, kun toinen avasi lomakkeen.

**Ratkaisu:** Lisätty `SelectedBy`-sarake `USysRecipients`-tauluun käyttäjäkohtaiseen tunnistukseen. Kaikki kolme aliohjelmaa päivitetty:

- **`Form_Load`**: `DB.Execute UPDATE` rajataan `WHERE SelectedBy = Kayttaja` — nollataan vain **oman** käyttäjätunnuksen unohtuneet valinnat
- **`Lista_DblClick`**: Valinnan yhteydessä kirjataan `taulu.Fields("SelectedBy") = NetworkUserName`, poiston yhteydessä tyhjennetään `= Null`
- **`Valmis_Click`**:
  - Vastaanottajien tarkistuskysely suodatettu `WHERE [To Distribution] = Yes AND SelectedBy = Kayttaja`
  - `INSERT INTO ... SELECT ... FROM USysRecipients` suodatettu `AND SelectedBy = Kayttaja`
  - Tallennuksen jälkeen `Siivotaan` käyttäjän valinnat `UPDATE ... WHERE SelectedBy = Kayttaja`
- Lisäkorjaus: poistettu aiemmin jäänyt `DB.Close` `Lista_DblClick`-Cleanup-lohkosta; `Dim Kayttaja As String` siirretty oikealle paikalle `Valmis_Click`-aliohjelman alkuun (VBA-vaatimus)

**Huom:** Tämä edellyttää `SelectedBy Text(50)` -kentän lisäämistä `USysRecipients`-tauluun Access-tietokannassa.

### Tietoturva: komentoinjektio poistettu

- `Form_DBUsers.cls` / `Command27_Click`: `net send` -kutsu korvattu `msg.exe`-kutsulla. Syötemerkkijono puhdistetaan (poistetaan `"`, `&`, `|`) komentoinjektioriskin eliminoimiseksi.
- `Form_USysOpenFile.cls` / `Command52_Click`: `Shell("C:\Windows\explorer.exe /e, " & Hakem)` korvattu turvallisella `Application.FollowHyperlink Hakem` -kutsulla, joka antaa käyttöjärjestelmän hoitaa polun avaamisen ilman Shell-komentoriviä.

### Tietoturva: SQL-injektioriskit

- `Form_USysReserve.cls` / `TalletaNappi_Click`: `Me.TNumber.Value` puhdistetaan `Replace(..., "'", "")` -kutsulla ennen SQL-lauseeseen ketjuttamista.
- `Form_USysNewDistribution.cls` / `Valmis_Click`: `Me.DistrNo` puhdistetaan `Replace(..., "'", "")` -kutsulla.
- `Form_USysAddToDistr.cls` / `Add_Click`: SQL-konstruoitu recordset korvattu `DCount`-tarkistuksella (`CLng()`-tyyppimuunnoksin) — poistaa injektioriskin kokonaan olemassaolotarkistuksessa.

### Vakaus: CurrentDb-viitteen sulkeminen poistettu

- `Form_SETTINGS.cls`: Poistettu `DB.Close` kahdesta kohdasta (normaalikulku + virheenkäsittely). `CurrentDb`-viitettä ei tule koskaan sulkea `.Close`-metodilla — riittää `Set DB = Nothing`.
- `Form_USysAddToDistr.cls`: Poistettu `DB.Close` Cleanup-lohkosta.
- `Form_DBUsers.cls` / `WhosOn`: Poistettu `dbCurrent.Close` `Exit_WhosOn`-lohkosta.
- `Form_USysNewDistribution.cls`: Poistettu `DB.Close` sekä `Form_Load`- että `Valmis_Click`-koodista.
- `Form_USysReserve.cls`: Korjattu `CurrentDb.OpenRecordset` → eksplisiittinen `DB`-muuttuja; poistettu implisiittinen `CurrentDb`-sulkeminen.

### Vakaus: Puuttuvat transaktiot ja Rollback lisätty

- `Form_USysAddToDistr.cls`: Lisätty `inTransaction`-lippu oikeaan Rollback-käsittelyyn; vanha `On Error Resume Next DBEngine.Rollback` korvattu ehdollisella tarkistuksella.
- `Form_USysNewDistribution.cls` / `Valmis_Click`: Lisätty `DBEngine.BeginTrans / CommitTrans / Rollback` — aiemmin kahden taulun (`UsysDISTRIBUTION` + `USysRecipByDistr`) tallennusoperaatio ei ollut transaktionaalinen. Virhetilanteessa tietokantaan saattoi jäädä orpo distribuutio ilman vastaanottajia.

### Vakaus: Excel "zombie"-prosessit estetty

- `Form_USysExcelReport.cls` / `OK_Click`: Lisätty yhtenäinen `Cleanup`-lohko, joka kutsuu `xlApp.Quit` ja `Set xlApp = Nothing` **aina** — myös virhetilanteessa. Aiemmin `ErrorHandler` ohitti sulkemisen, jolloin jokainen epäonnistunut ajo jätti piilotetun `EXCEL.EXE`-prosessin muistiin. Lisätty myös tiedoston olemassaolotarkistus (`Dir(ExcelPolku)`) ennen Excelin käynnistystä. Vaihdettu `Excel.Application` → `Object` (late binding) vakauden parantamiseksi.

### Vakaus: Puuttuvat EOF-tarkistukset raporttimoduuleissa

- `Report_TRANSMITTAL.cls`, `Report_TRANSMITTAL Copy.cls`, `Report_Copy of TRANSMITTAL.cls`: Lisätty `If Not taulu.EOF Then` -tarkistus ennen kenttien lukemista. Ilman tarkistusta tyhjä `Projinfo`-taulu aiheutti ajoaikaisen "No current record" -virheen. Kaikissa lisätty myös `On Error GoTo ErrorHandler` sekä `taulu.Close` Cleanup-lohkossa.

### Vakaus: Puuttuvat taulu.Close-kutsut

- `Form_USysAddToDistr.cls`: Lisätty `taulu.Close` ennen `Set taulu = Nothing` Cleanup-lohkossa.
- `Form_USysReserve.cls`: Korjattu — aiempi labyrinttimainen `Exit Sub` useassa kohdassa korvattu yhtenäisellä `Cleanup`-lohkolla.
- `Form_USysExcelReport.cls`: Lisätty `DTaulu.Close` / `ExcelTaulu.Close` silmukan jälkeen ja Cleanup-lohkossa.

### Suorituskyky

- `Form_USysAddToDistr.cls`: Olemassaolotarkistus muutettu `dbOpenDynaset` + `RecordCount` → `DCount` — ei avata recordsettiä pelkän lukumäärän tarkistamiseen.
- `Form_USysNewDistribution.cls` / `Valmis_Click`: Vastaanottajien kopiointisilmukka (Do While + AddNew per rivi) korvattu yhdellä `INSERT INTO ... SELECT`-kyselyllä (`DB.Execute SQL, dbFailOnError`).
- `Report_TRANSMITTAL`-tiedostot: `CurrentDb.OpenRecordset("Projinfo")` muutettu `dbOpenSnapshot`-tilaan — raportti ei muokkaa dataa.

### Form_SETTINGS.cls: Tyhjyystarkistus

- Lisätty `If Len(Nz(Me.Nimi.Value, "")) > 0 Then` -ehto ennen `AppTitle`-tallennusta. Estää tyhjän sovelluksen otsikon asettamisen.

---

## Kriittiset muutokset — alkuperäinen kierros (2026-03-03)

### 64-bit API -korjaukset

- `GlobalVBAs.vba` ja `ForDocuments.vba`: `GetUserNameA`- ja `GetComputerNameA`-kutsuissa `nSize` muutettu `LongPtr` → `ByRef Long` — API kirjoittaa LPDWORD-osoittimeen (32-bit DWORD, 4 tavua), ei 64-bittiseen LongPtr:iin. Aiheutti Type Mismatch -ajovirheen 64-bittisessä Office-ympäristössä.
- `ForDocuments.vba`: `lngStringLength` muutettu `LongPtr` → `Long` samasta syystä.

### DAO-transaktio-ongelmat

- Kaikki `DB.BeginTrans` / `DB.CommitTrans` / `DB.Rollback` -kutsut muutettu `DBEngine.*` -muotoon — DAO:ssa transaktiot kuuluvat `DBEngine`-tasolle, ei `Database`-objektille. Aiheutti "function or interface marked as restricted" -käännösvirheen.
- Korjattu tiedostoissa: `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`.

### DAO:lle sopimaton `.State`-tarkistus

- Poistettu kaikki `If taulu.State = 1 Then taulu.Close` -rakenteet — `.State`-ominaisuus kuuluu ADO:lle, ei DAO:lle. Aiheutti "method or data member not found" -käännösvirheen.
- Korjattu tiedostoissa: `Form_USysAddedDistr.cls`, `Form_USysOpenFile.cls`, `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`, `Form_USysNewDistribution.cls`.

### Me.-etuliite (Option Explicit -yhteensopivuus)

- Lisätty `Me.`-etuliite kaikkiin lomakekontrolliviittauksiin kaikissa 20 lomake- ja raporttitiedostossa. Ilman etuliitettä `Option Explicit` tulkitsee kontrollin nimen määrittelemättömäksi muuttujaksi → "Variable not defined" / "method or data member not found".
- Erityistapaukset: `Form_SETTINGS.cls`-funktiossa `PoimiPolku(Kentta As Control)` parametri `Kentta` on tarkoituksellisesti ilman `Me.` (se on `Control`-tyyppinen funktioparametri, ei lomakekontrolli).

### Muut ajonaikaiset virheet

- `Form_USysRevText.cls`: `MsgBox`-merkkijonossa `\"` → `""` (JSON-eskejppaus jäi koodiin, aiheutti syntaksivirheen VBA:ssa).
- `Form_USysOpenFile.cls`, rivi `Tied.RowSource` → `Me.Tied.RowSource` (jäi edelliseltä kierrokselta).
- `Form_USysEditDistribution.cls`: `Lista.Requery` → `Me.Lista.Requery`.
- `Form_USysReserve.cls`: `TPaperSize`, `TDiscipline`, `TNumber`, `TClientNo`, `ClientNo` → kaikki `Me.`-etuliitteellä.

---

## Siivous ja optimointi

### Kuollut koodi poistettu

- `GlobalVBAs.vba`: Kaikki `Debug.Print`-rivit poistettu parsintafunktioista (`HaeTekija`, `HaeRevisioija`, `HaeRevisioijaPvm`, `EkaRevRivi`, `HaeRevisio`, `HaeViimPaiva`, `HaePaiva`) — nämä olivat kehityksenaikaisia lokitulostuksia, jotka jäivät koodiin.
- `Form_USysRevText.cls`: Käyttämätön `Dim Rivit As String` -muuttuja poistettu `Form_Resize`-aliohjelmasta.
- Tuplarivit poistettu parsintafunktioiden alusta.

### Resurssien sulkeminen

- `GlobalVBAs.vba` / `SetStartup`: Lisätty puuttuvat `taulu.Close` ja `DB.Close` ennen `Set x = Nothing` -kutsuja. Ilman sulkemista DAO-tietueet voivat jäädä auki tietokantaan.

### Kommentit suomeksi

- `GlobalVBAs.vba`: Kaikki englanninkieliset `'''`-JSDoc-kommentit ja `'`-inline-kommentit muutettu suomeksi revisioparsintafunktioissa.
- `GlobalVBAs.vba`: Englanninkielinen funktiolistablokkikommentti (`' Functions for parsing revisions`) käännetty suomeksi.
- Kaikki 24 tiedostoa: Lisätty suomalainen moduuliotsikko (LOMAKE/MODUULI/RAPORTTI, SOVELLUS, KUVAUS, PÄIVITETTY).

### Virheenkäsittely

- `Form_USysOpenFile.cls`: `Command52_Click` — kommentti "Hakemistoa ei ole" suomennettu.
- `Form_USysShowCommon.cls`: Otsikon kirjoitusvirheet korjattu (`Yleiskkäyttöinen` → `Yleiskäyttöinen`, `muokkauksest` → `muokkaus`; kuvaus täsmennetty).
- `Form_USysOpenFile.cls`: Vanha Windows NT -polku `C:\WINNT\EXPLORER.EXE` → `C:\Windows\explorer.exe` (nyt korvattu `FollowHyperlink`-kutsulla kokonaan).

---

## Tiedostot joihin ei tehty muutoksia

| Tiedosto | Syy |
|---|---|
| `Form_USysDISTRIB.cls` | Kaikki koodi kommentoitu pois alkuperäisessä |
| `Report_USYSTRANSMITTALFP.cls` | Tyhjä tiedosto |

### 64-bit API -korjaukset

- `GlobalVBAs.vba` ja `ForDocuments.vba`: `GetUserNameA`- ja `GetComputerNameA`-kutsuissa `nSize` muutettu `LongPtr` → `ByRef Long` — API kirjoittaa LPDWORD-osoittimeen (32-bit DWORD, 4 tavua), ei 64-bittiseen LongPtr:iin. Aiheutti Type Mismatch -ajovirheen 64-bittisessä Office-ympäristössä.
- `ForDocuments.vba`: `lngStringLength` muutettu `LongPtr` → `Long` samasta syystä.

### DAO-transaktio-ongelmat

- Kaikki `DB.BeginTrans` / `DB.CommitTrans` / `DB.Rollback` -kutsut muutettu `DBEngine.*` -muotoon — DAO:ssa transaktiot kuuluvat `DBEngine`-tasolle, ei `Database`-objektille. Aiheutti "function or interface marked as restricted" -käännösvirheen.
- Korjattu tiedostoissa: `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`.

### DAO:lle sopimaton `.State`-tarkistus

- Poistettu kaikki `If taulu.State = 1 Then taulu.Close` -rakenteet — `.State`-ominaisuus kuuluu ADO:lle, ei DAO:lle. Aiheutti "method or data member not found" -käännösvirheen.
- Korjattu tiedostoissa: `Form_USysAddedDistr.cls`, `Form_USysOpenFile.cls`, `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`, `Form_USysNewDistribution.cls`.

### Me.-etuliite (Option Explicit -yhteensopivuus)

- Lisätty `Me.`-etuliite kaikkiin lomakekontrolliviittauksiin kaikissa 20 lomake- ja raporttitiedostossa. Ilman etuliitettä `Option Explicit` tulkitsee kontrollin nimen määrittelemättömäksi muuttujaksi → "Variable not defined" / "method or data member not found".
- Erityistapaukset: `Form_SETTINGS.cls`-funktiossa `PoimiPolku(Kentta As Control)` parametri `Kentta` on tarkoituksellisesti ilman `Me.` (se on `Control`-tyyppinen funktioparametri, ei lomakekontrolli).

### Muut ajonaikaiset virheet

- `Form_USysRevText.cls`: `MsgBox`-merkkijonossa `\"` → `""` (JSON-eskejppaus jäi koodiin, aiheutti syntaksivirheen VBA:ssa).
- `Form_USysOpenFile.cls`, rivi `Tied.RowSource` → `Me.Tied.RowSource` (jäi edelliseltä kierrokselta).
- `Form_USysEditDistribution.cls`: `Lista.Requery` → `Me.Lista.Requery`.
- `Form_USysReserve.cls`: `TPaperSize`, `TDiscipline`, `TNumber`, `TClientNo`, `ClientNo` → kaikki `Me.`-etuliitteellä.

---

## Siivous ja optimointi

### Kuollut koodi poistettu

- `GlobalVBAs.vba`: Kaikki `Debug.Print`-rivit poistettu parsintafunktioista (`HaeTekija`, `HaeRevisioija`, `HaeRevisioijaPvm`, `EkaRevRivi`, `HaeRevisio`, `HaeViimPaiva`, `HaePaiva`) — nämä olivat kehityksenaikaisia lokitulostuksia, jotka jäivät koodiin.
- `Form_USysRevText.cls`: Käyttämätön `Dim Rivit As String` -muuttuja poistettu `Form_Resize`-aliohjelmasta.
- Tuplarivit poistettu parsintafunktioiden alusta.

### Resurssien sulkeminen

- `GlobalVBAs.vba` / `SetStartup`: Lisätty puuttuvat `taulu.Close` ja `DB.Close` ennen `Set x = Nothing` -kutsuja. Ilman sulkemista DAO-tietueet voivat jäädä auki tietokantaan.

### Kommentit suomeksi

- `GlobalVBAs.vba`: Kaikki englanninkieliset `'''`-JSDoc-kommentit ja `'`-inline-kommentit muutettu suomeksi revisioparsintafunktioissa.
- `GlobalVBAs.vba`: Englanninkielinen funktiolistablokkikommentti (`' Functions for parsing revisions`) käännetty suomeksi.
- Kaikki 24 tiedostoa: Lisätty suomalainen moduuliotsikko (LOMAKE/MODUULI/RAPORTTI, SOVELLUS, KUVAUS, PÄIVITETTY).

### Virheenkäsittely

- `Form_USysOpenFile.cls`: `Command52_Click` — kommentti "Hakemistoa ei ole" suomennettu.
- `Form_USysShowCommon.cls`: Otsikon kirjoitusvirheet korjattu (`Yleiskkäyttöinen` → `Yleiskäyttöinen`, `muokkauksest` → `muokkaus`; kuvaus täsmennetty).
- `Form_USysOpenFile.cls`: Vanha Windows NT -polku `C:\WINNT\EXPLORER.EXE` → `C:\Windows\explorer.exe`.

---

## Tiedostot joihin ei tehty muutoksia

| Tiedosto | Syy |
|---|---|
| `Form_USysDISTRIB.cls` | Kaikki koodi kommentoitu pois alkuperäisessä |
| `Report_USYSTRANSMITTALFP.cls` | Tyhjä tiedosto |
