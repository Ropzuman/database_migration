# Phase 3 – Viimeistely: Kattava Koodianalyysi

**Päiväys:** 2025-07-14  
**Analysoituja tiedostoja:** 71  
**Löydettyjä ongelmia:** ~75 yksittäistä havaintoa

---

## Prioriteettijärjestys

| Taso | Kuvaus |
|------|--------|
| 🔴 KRIITTINEN | Toiminnallinen virhe tai prosessivuoto |
| 🟠 TÄRKEÄ | Puuttuva virheenkäsittely tai resurssivuoto |
| 🟡 HUOMIO | Kuollut koodi, käyttämättömät muuttujat, siivous |
| 🔵 PIENOISVIRHE | Kommenttivirhe, enkoodaus |

---

## Access / DOCUMENTS

### ForDocuments.vba

- 🔵 Kommentissa encoding-vika: `"m??ritetty"` → pitää olla `"määritetty"` (`IsLoaded`-funktio)
- 🟡 `lpSelPath As Long` — julistettu `ValitseHakem`-funktiossa, ei koskaan käytetty (kuollut muuttuja)
- 🟡 `Const LPTR` — julistettu mutta ei koskaan käytetty (kuollut vakio)

### GlobalVBAs.vba

✅ OK

### Form_DISTRIBUTION.cls

- 🔴 **KRIITTINEN:** `ExcelRep()`-funktio luo `xlApp = CreateObject("Excel.Application")` mutta `xlApp.Quit` **ei koskaan kutsuta** ennen `Set xlApp = Nothing` → Excel-prosessi jää orpolaiseksi joka kutsukerralla

### Form_DOCUMENTS.cls

✅ OK

### Form_SETTINGS.cls

✅ OK

### Form_USysAddDocument.cls

- 🟠 `Form_Load`: `Set taulu = CurrentDb.OpenRecordset(...)` — `taulu.Close` puuttuu ennen `Set taulu = Nothing`

### Form_DBUsers.cls

✅ OK

### Form_USysAddedDistr.cls

✅ OK

### Form_USysAddToDistr.cls

- 🟠 `Form_Load`: Ei `On Error` -käsittelijää

### Form_USysDISTRIB.cls

- 🟡 **KUOLLUT KOODI:** Koko `Form_Load`-runko (~20 riviä) on kommentoitu pois

### Form_USysDocs.cls

✅ OK

### Form_USysEditDistribution.cls

- 🟠 `Form_Load` siivous: `Set taulu = Nothing` ilman edeltävää `taulu.Close`
- 🟠 `Lista_DblClick` siivous: sekä `Set taulu = Nothing` että `Set Taulu2 = Nothing` ilman `.Close`-kutsuita

### Form_USysExcelReport.cls

✅ OK

### Form_USysNewDistribution.cls

- 🟡 `Valmis_Click`: `Taulu2 As DAO.Recordset` julistettu mutta **ei koskaan avattu eikä käytetty** (kuollut muuttuja); siivousalue viittaa siihen tarpeettomasti

### Form_USysNewRecipient.cls

- 🟠 `Lista_DblClick` siivous: `Set taulu = Nothing` ilman edeltävää `taulu.Close`

### Form_USysOpenFile.cls

✅ OK

### Form_USysRecipientsFrm.cls

✅ OK

### Form_USysReserve.cls

✅ OK

### Form_USysRevText.cls

✅ OK

### Form_USysShowCommon.cls

✅ OK

### Form_USysStart.cls

- 🟡 `Public ajastin As Long` — asetetaan `Timer`-arvoksi `Form_Load`:ssa mutta **arvoa ei koskaan lueta missään** (käyttämätön muuttuja)

### Report_TRANSMITTAL.cls

✅ OK

### Report_Copy of TRANSMITTAL.cls

✅ OK

### Report_TRANSMITTAL Copy.cls

✅ OK

### Report_USYSTRANSMITTALFP.cls

- 🟡 Sisältää vain `Option Compare Database` + `Option Explicit` — kokonaan tyhjä raporttilomake; tarkista onko tämä tarkoituksellinen

---

## Access / instru3

### general.bas

✅ OK

### Sivunumerointi.bas

✅ OK

### USysCheck.bas

✅ OK

### Form_CopyLoops.cls

- 🟠 `HaeLoopit_Click`: `KohdeTaulu = CurrentDb.OpenRecordset("devTable")` avataan **joka silmukkakierroksella** myös niillä kierroksilla, joilla laitetauluun ei kosketakaan → resurssivuoto
- 🟡 ~15 kappaletta aktiivisia `Debug.Print`-lauseita tuotantokoodissa (`HaeLoopit_Click`, `ValitseKanta_Click`)

### Form_SizingOut.cls

✅ OK

### Form_DBUsers.cls

- 🟡 `X = Dir(SPath)` — X saa arvon mutta arvoa **ei koskaan tarkisteta** eikä käytetä (käyttämätön sijoitus)

### Form_Linkkien vaihto.cls

- 🟡 `Taulut() As String` — julistettu, ei koskaan käytetty (kuollut muuttuja)
- 🟡 `s As Integer` — julistettu, ei koskaan käytetty (kuollut muuttuja)

---

## Access / LoopCircuit

### General.bas

✅ OK

### USysCheck.bas

✅ OK

### Module1.bas (CustomMessage)

✅ OK

### For ACAD Utility.bas (APIKoodit)

- 🟡 `lpSelPath As LongPtr` — julistettu `ValitseHakem`-funktiossa, ei koskaan käytetty (kuollut muuttuja)
- 🟡 `Const LPTR` — julistettu mutta ei koskaan käytetty (kuollut vakio)

### Form_Tee Kuvat.cls

- 🟡 **KUOLLUT KOODI:** ~20 kommentoitua `LisaaLokiin "..."` -debuggauslokin kutsua ympäri `TeeKuvat_Click`, `HaeIPoints`, `VaihdaOtsikkotiedot`

### Form_DBUsers.cls

✅ OK

### Form_Linkkien vaihto.cls

✅ OK

---

## Access / Lukituskaavio

### APIKoodit.bas

✅ OK

### Koodit.bas

- 🟡 `KillLinks`: `LinkCount` kasvatetaan silmukassa mutta **arvoa ei koskaan käytetä** (käyttämätön muuttuja)
- 🔴 `AvaaBlock`: Puuttuu `Exit Function` ennen `ErrorHandler:`-tunnistetta → virheettömässä suorituksessa koodi **putoaa virhekäsittelijään** normaalin suorituksen jälkeen (ohjausvirtavirhe)

### Form_Aloitus.cls

- 🟡 `Linkkaa`-aliohjelma: `rstLinkki As DAO.Recordset` julistettu mutta **ei koskaan avattu eikä käytetty** (kuollut muuttuja)
- 🟡 `tdfLinkki` — ei aseteta `Nothing`-arvoksi `TableDefs.Append`-kutsun jälkeen
- 🟠 `Command0_Click`: Ei virheenkäsittelijää

### Form_Interlocking.cls

- 🟡 `Blokki_Change()`: Runko kokonaan kommentoitu pois — käyttämätön tapahtumakäsittelijä kuollella koodilla (~20 riviä)
- 🟡 `Blokki_Click()`: Runko kokonaan kommentoitu pois — tyhjä aliohjelma kuollella koodilla
- 🟡 `Command157_Click()`: **Täysin tyhjä aliohjelma** (ei sisältöä)
- 🟡 `Ctl__Click()`: **Täysin tyhjä aliohjelma** (ei sisältöä)
- 🟠 `LueAttribuutit`: `Asetukset = CurrentDb.OpenRecordset("SETTINGS")` — tietuejoukko **ei koskaan suljettu** eikä vapautettu
- 🟡 `Command151_Click`: `BlockCount`-muuttuja kasvatetaan mutta **arvoa ei koskaan näytetä tai käytetä** (käyttämätön); `Taul`-tietuejoukko jää siivoamatta `Exit Sub`-haaran kautta

### Form_IntLoopDescr20Update.cls

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`
- 🟡 `DB As DAO.Database` — julistettu mutta **ei koskaan avattu eikä käytetty** (kuollut muuttuja)
- 🟠 Ei lainkaan virheenkäsittelyä

### Form_Funktiokaavio.cls

- 🟡 `Command0_Click`: `MyQuery As DAO.QueryDef`, `sSQL`, `ssSQL`, `MyString` — julistettu mutta **käyttämätön** (kuolleet muuttujat); `L = Chr(34)` asetetaan mutta **ei koskaan käytetä**
- 🟠 `Command0_Click`: Ei virheenkäsittelijää; `qry`, `qry2`, `tbl` -tietuejoukot **eivät sulkeudu** proseduurin lopussa
- 🟠 `Command12_Click`: Ei virheenkäsittelijää; `Taul`-tietuejoukko jää siivoamatta

### Form_LineForm.cls

✅ OK

### Form_Linkkien vaihto.cls

- 🟡 `Taulut() As String` — kuollut muuttuja (sama kuin instru3-versio)
- 🟡 `s As Integer` — kuollut muuttuja

### Form_TOACAD_Loops subform.cls

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`

### Form_TOACAD_Motors subform.cls

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`

### Form_TOACAD_Sekvens subform.cls

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`
- 🟡 Kommentoitu `Me.Parent.Text_2.VALUE = DESC2` -sijoitus — kuollut koodi

### Form_TOACAD_Sekvens2 subform.cls

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`
- 🟡 `DESC1_Click()`: **Täysin tyhjä aliohjelma**
- 🟡 Kommentoidut `Text_1`, `Text_2`, `Text_3` -sijoitukset — kuollut koodi

---

## Access / MAINEQ

### GeneralCodes.bas

- 🔵 `HaeViimPaiva`-funktion kommentissa kirjoitusvirhe: `"Sisyötteessä"` → pitää olla `"Syötteessä"`

### DataToACAD.bas

- 🟡 `MakeLocFiles`: **Kovakoodattu projektikohtainen polku** `"p:\acaddata\projekti\agropm10\tyo\instloc.txt"` — ei siirrettävä muihin projekteihin

### For ACAD Utility.bas (MAINEQ)

✅ OK

### USysCheck.bas (MAINEQ)

✅ OK

### Form_EQUIPMENT.cls

- 🟡 `ExtData_Click()`: Runko kokonaan kommentoitu — **tyhjä aliohjelma kuollella koodilla**
- 🟡 `Form_Current()`: Runko kokonaan kommentoitu — **tyhjä aliohjelma kuollella koodilla**
- 🟡 `EqGroup_AfterUpdate()`: Sisältää vain kommentoitu kutsu — **käytännössä tyhjä tapahtumakäsittelijä**
- 🟡 `suffix_afterupdate()`: Ainoa aktiivinen rivi kommentoitu — kuollut koodi
- 🟠 `EqType_AfterUpdate`: `Set Taulukko = Nothing` ilman edeltävää `Taulukko.Close`

### Form_MAINEQ_form.cls

✅ OK

### Form_GeneroiMoottorikuvat.cls

- 🟡 `Tiedot As DAO.Recordset` — julistettu, **ei koskaan avattu** (kuollut muuttuja); siivousalueella tyhjä `Set Tiedot = Nothing` -kutsu
- 🟠 `Kuvat.Close` puuttuu ennen `Set Kuvat = Nothing` -kutsuita
- 🟡 `Set Kuvat = Nothing` kutsutaan **kahdesti** proseduurin lopussa (redundantti)
- 🟠 `GenKuvat_Click`: Ei yleistä virheenkäsittelijää (vain paikallinen AutoCAD-yhteyden tarkistus)

### Form_KuvienGenerointi.cls

- 🟠 `GenRev_Click`: `Revisiot`-tietuejoukko avataan silmukan sisällä mutta **ei koskaan suljeta** eikä vapauteta (`Revisiot.Close` ja `Set Revisiot = Nothing` puuttuvat)
- 🟡 `GenRev_Click`: `i` ja `j` -laskurit **eivät nollaudu** ulomman silmukan iteraatioiden välillä → virheelliset sivunumerot toisesta asiakirjasta eteenpäin
- 🟠 `GenRev_Click`: `Kuvat.Close` puuttuu
- 🟠 `Teetaulukko_Click`: `Taulu`-tietuejoukko avataan useita kertoja `Select Case` -haaroissa mutta **ei koskaan suljeta**; kaikki neljä tietuejoukkomuuttujaa (`KuviinTbl`, `zTitleTbl`, `TEKSTITTbl`, `Taulu`) puuttuvat `.Close` ennen `Set = Nothing`
- 🟠 `Teetaulukko_Click`: Ei virheenkäsittelijää
- 🟠 `GenKuvat_Click`: `Tiedot.Close` puuttuu; `Set Kuvat = Nothing` kahdesti

### Form_Revisiointi.cls

- 🟠 `PaivitaIDRev`: `Taul`-tietuejoukko avataan **kahdesti** peräkkäin ilman välissä tapahtuvaa `Taul.Close` → ensimmäinen yhteys jää avoimeksi
- 🟠 `Tee_Click`: `Vert`-tietuejoukko jää **sulkematta ja vapauttamatta** proseduurin lopussa
- 🟡 `PaivitaIDRev`: `Left$(Taul.Fields(0), InStr(Taul.Fields(0), " ") - 1)` — jos kentässä ei ole välilyöntiä, `InStr` palauttaa 0 → `0 - 1 = -1` → `Left$` kaatuu ajonaikavirheeseen (potentiaalinen kaatuminen)

---

## Access / PIPE

### Koodit.bas

- 🟠 `AvaaBlock`: `Taulu = CurrentDb.OpenRecordset("SELECT * FROM MANVALVEDATA ...")` — tietuejoukko **ei koskaan suljettu eikä vapautettu** (ei `Taulu.Close`, ei `Set Taulu = Nothing`)

### Form_zFunc.cls

- 🟡 `Command3_Click`: `q1 As String` — julistettu mutta **ei koskaan käytetty** (kuollut muuttuja)
- 🟡 `Command3_Click`: `qq1 As String` — julistettu mutta **ei koskaan käytetty** (kuollut muuttuja)

### Form_TYÖKALUT.cls

✅ OK — erittäin puhdas rakenne, hyvä virheenkäsittely läpi koko tiedoston

### Form_USysFlowPickNo.cls

- 🔵 **Enkoodauskorruptio:** Useissa kommenteissa rikkinäiset skandinaaviset merkit (esim. `LÃ¶ytÃ¤Ã¤` → pitää olla `Löytää`; `tyhjÃ¤` → `tyhjä`) — koko tiedostossa systemaattinen ongelma

### Form_USysPipeFromTo.cls

- 🔵 **Enkoodauskorruptio:** Laajamittainen skandinaavisten merkkien vioittuminen kommenteissa läpi koko tiedoston (esim. `TyhjentÃ¤Ã¤` → `Tyhjentää`; `sÃ¤iliÃ¶n` → `säiliön`; `TekstikenttÃ¤` → `Tekstikenttä`)

### Form_USysPipeToOther.cls

- 🔵 **Enkoodauskorruptio:** Sama laajamittainen enkoodausongelma kuin yllä (esim. `kiertÃ¤Ã¤` → `kiertää`; `YhdistetÃ¤Ã¤n` → `Yhdistetään`; `alustetÃ¤Ã¤n` → `alustetaan`)

---

## Access / FunctionDiagrams

### General.bas

- 🟡 `Show_last(criterias As Variant)`: parametri `criterias` **julistettu mutta ei koskaan käytetty** proseduurin rungossa (käyttämätön parametri)
- 🟡 `Show_last_criteria(criterias As Variant)`: sama ongelma — parametri **ei koskaan käytetty**
- 🟡 `Declare PtrSafe Function wu_GetUserName` ja `GetOpenFileName` — puuttuu `#If VBA7 Then` ehdollinen käännösblokki

### USysCheck.bas

- 🟠 `SniffUser`: `Taulu.Close` **puuttuu** ennen `Set Taulu = Nothing` ja `Set DB = Nothing`

---

## Access / Function_descriptions_html

### GeneralCodes.bas

✅ OK

### For ACAD Utility.bas

⚠️ Tiedostoa ei löydy (`For ACAD Utility.bas`) — tarkista sijainti tai onko poistettu

---

## AutoCAD / exported / Arkistotulostus

### General.bas

- 🟡 `ValitseHakem`: `lpSelPath As Long` — julistettu mutta **ei koskaan käytetty** (kuollut muuttuja)
- 🟡 `ValitseHakem`: `Const LPTR` — julistettu mutta **ei koskaan käytetty** (kuollut vakio)
- 🟡 `OPENFILENAME`-tyypissä `lCustData As Long` ja `lpfnHook As Long` — pitäisi olla **`LongPtr`** 64-bittistä yhteensopivuutta varten
- 🔵 `MPlot`-aliohjelma: kommentti `'N?ytet??n tulostusformi` — tallennettu väärällä merkistöllä, skandit korruptoituneet
- 🟡 `Declare PtrSafe`-lauseet ilman `#If VBA7 Then`-suojaa

---

## AutoCAD / exported / MultiPlot

### General.bas

- 🟡 `lpSelPath As Long` — kuollut muuttuja (sama kuin Arkistotulostus)
- 🟡 `Const LPTR` — kuollut vakio
- 🟡 `OPENFILENAME`: `lCustData As Long`, `lpfnHook As Long` → pitäisi olla **`LongPtr`**
- 🟡 `DummyFunc(ByVal param As Long) As Long` — `param` pitäisi olla **`LongPtr`** koska se vastaanottaa `AddressOf`-funktion osoitteen
- 🔵 `MPlot` ja `LueTiedosto`: suomenkieliset kommentit vioittuneita (enkoodauskorruptio)
- 🟡 Puuttuu `#If VBA7 Then` -suoja kaikille `Declare`-lauseille

---

## AutoCAD / exported / MultiPlot_TW

### General.bas

- 🟡 Samat ongelmat kuin MultiPlot/General.bas:
  - `lpSelPath As Long` kuollut muuttuja; `Const LPTR` kuollut vakio
  - `lCustData As Long`, `lpfnHook As Long` → `LongPtr`
  - `DummyFunc` parametri `Long` → `LongPtr`
  - `MPlot`-kommentin enkoodauskorruptio
  - Puuttuu `#If VBA7 Then`

---

## AutoCAD / exported / InterlockIng

### Start.bas

- 🟡 `OPENFILENAME`-tyypissä `lCustData As Long`, `lpfnHook As Long` → pitäisi olla **`LongPtr`**
- 🟡 Puuttuu `#If VBA7 Then` -suoja `Declare`-lauseille

---

## AutoCAD / exported / Jonotulostus_64bit

### General.bas

- 🔴 **KRIITTINEN:** Puuttuu `Option Explicit`
- 🔵 `MPlot`/`JTulostus`-kommentti `'N?ytet??n tulostusformi` — enkoodauskorruptio

### FileDialogs.cls

- 🔴 `ShowSave`: `udtStruct.lStructSize = Len(udtStruct)` — **pitää olla `LenB(udtStruct)`** (64-bitin rakennekoko lasketaan tavuina, ei merkkeinä; `Len` antaa virheellisen arvon ja `GetSaveFileName` epäonnistuu hiljaa)
- 🟡 `OPENFILENAME`-tyypissä `lCustData As Long` → pitäisi olla **`LongPtr`**
- 🟡 Puuttuu `#If VBA7 Then` -suoja `Declare`-lauseille

---

## Excel / Moduulit / AcadDATA

### AcadHelpers.bas

✅ OK

### DATA.bas ja Koodit.bas

- 🔴 **KRIITTINEN:** Molemmat tiedostot sisältävät **täsmälleen saman koodin** — julkiset prosedurit `TuoDATA`, `TuoDATA_All`, `TuoDATA_Selected` ja julkiset muuttujat `oACAD`, `oDOC` on määritelty kahdesti. Jos molemmat moduulit ladataan samaan VBA-projektiin, aiheutuu **käännösvirhe tuplamarittelystä**. Kuolleesta versiosta on päätettävä kumpi on oikea ja toinen siirrettävä `OLD/`-kansioon.
- *Kumpikin tiedosto erikseen* — virheenkäsittely, `wasCalcAuto`/`prevScreen`/`prevEvents`-palauttaminen: ✅ Hyvä rakenne

---

## Excel / Moduulit / Listojen kyselyt

### Module1.bas

- 🟡 Useat `Debug.Print`-lauseet aktiivisena tuotantokoodissa — `BeginFastMode`, `EndFastMode`, `HaeData`, koko tiedosto täynnä diagnostiikkalokia. Harkitse ohjaaminen `DEBUG_TRACE`-vakion taakse (kuten AcadDATA/Koodit.bas tekee).

### Module2.bas

✅ OK

### Module3.bas

✅ OK

### ThisWorkbook.cls

✅ OK

---

## Koontitaulukko — Kriittiset korjaustarpeet

| # | Tiedosto | Ongelma | Taso |
|---|---------|---------|------|
| 1 | `Form_DISTRIBUTION.cls` (DOCUMENTS) | `xlApp.Quit` puuttuu → Excel-prosessivuoto | 🔴 |
| 2 | `Form_IntLoopDescr20Update.cls` (Lukituskaavio) | `Option Explicit` puuttuu | 🔴 |
| 3 | `Form_TOACAD_Loops/Motors/Sekvens/Sekvens2` (Lukituskaavio) | `Option Explicit` puuttuu (4 tiedostoa) | 🔴 |
| 4 | `Koodit.bas` (Lukituskaavio) | `AvaaBlock`: putoaa virheenkäsittelijään normaalilla suorituksella | 🔴 |
| 5 | `Jonotulostus_64bit/General.bas` | `Option Explicit` puuttuu | 🔴 |
| 6 | `FileDialogs.cls` (Jonotulostus_64bit) | `ShowSave`: `Len()` → `LenB()` — 64-bitin rakennekoko virheellinen | 🔴 |
| 7 | `DATA.bas` + `Koodit.bas` (Excel/AcadDATA) | Duplikaattimoduulit — sama koodi kahdessa tiedostossa | 🔴 |
| 8 | `Form_KuvienGenerointi.cls` (MAINEQ) | `Revisiot`-tietuejoukko ei suljetaan silmukassa; laskurit ei nollata | 🟠 |
| 9 | `Koodit.bas` (PIPE) | `AvaaBlock`: `Taulu`-tietuejoukko jää avoimeksi | 🟠 |
| 10 | `Form_Revisiointi.cls` (MAINEQ) | `Taul` avataan kahdesti, ei suljeta; potentiaalinen kaatuminen `Left$`:ssa | 🟠 |
| 11 | `Form_USysFlowPickNo/PipeFromTo/PipeToOther.cls` (PIPE) | Laajamittainen enkoodauskorruptio kommenteissa (3 tiedostoa) | 🔵 |
| 12 | AutoCAD `General.bas` -tiedostot (Arkistotulostus, MultiPlot, MultiPlot_TW) | `lCustData`/`lpfnHook` → `LongPtr`; `DummyFunc` → `LongPtr`; ei `#If VBA7` | 🟡 |

---

## Loppuhuomiot

- **Toistuva kaava:** `Set X = Nothing` ilman `X.Close` esiintyy ~10 eri tiedostossa — suositellaan koodekatselmus koko koodikannalle tämän korjaamiseksi.
- **Toistuva kaava:** `Taulut() As String` + `s As Integer` -pari instru3, LoopCircuit ja Lukituskaavio -versioissa `Form_Linkkien vaihto.cls` — identtiset kuolleet muuttujat kolmessa kopiossa.
- **AutoCAD General.bas -tiedostot:** Kolme tiedostoa (Arkistotulostus, MultiPlot, MultiPlot_TW) ovat käytännössä identtisiä — harkitse yhteistä moduulia koodin toistamisen välttämiseksi.
- **PIPE-lomakkeiden enkoodaus:** Kolmen lomakkeen (`FlowPickNo`, `PipeFromTo`, `PipeToOther`) kommentit on tallennettu väärällä merkistöllä. Tiedostot on avattava ja tallennettava uudelleen oikealla UTF-8/CP1252 -asetuksella.
