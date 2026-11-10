# Muutosloki ó MAINEQ-moduuli

**Tiedosto:** `Access/MAINEQ/` (DataToACAD.bas, GeneralCodes.bas, USysCheck.bas, Form_DBUsers.cls, ja muut)
**P‰iv‰m‰‰r‰:** 2025-11-11 / 2025-11-12 / 2026-03-04
**Tekij‰:** 64-bit migraatioprojekti (Pro gradu)

---

## Kriittiset muutokset (Vaihe 1ñ2)

### 64-bittinen API-korjaus ó USysCheck.bas

- `GetUserNameA` ja `GetComputerNameA`-funktioiden `Declare`-lauseet p√§ivitetty:
  - Lis√§tty `PtrSafe`-avainsana 64-bittist√§ Office/AutoCAD-ymp√§rist√∂√§ varten
  - `nSize`-parametri muutettu `LongPtr` ‚Üí `Long` (Win32 API kirjoittaa 4-tavuisen DWORD:in, ei 8-tavuisen)
- `GetLastError`- ja `SHGetPathFromIDList`-funktiot p√§ivitetty `PtrSafe`-m√§√§reell√§

### Option Explicit ‚Äî kaikki moduulit

- Lis√§tty `Option Explicit` kaikkiin moduuleihin, joista se puuttui (6 lomaketta)

### Me.-etuliite ‚Äî Access-lomakkeet

- Kaikki lomakeohjauspyynn√∂t p√§ivitetty `Me.`-etuliitteell√§ (`ControlName` ‚Üí `Me.ControlName`) implisiittisten viittausvirheiden est√§miseksi

### Debug.Print-rivien poisto

- Poistettu noin 50 `Debug.Print`-rivi√§ `DataToACAD.bas`-tiedostosta
- Poistettu 2 `Debug.Print`-rivi√§ `MoottTilaus`-moduulista
- Poistettu 1 `Debug.Print`-rivi `Revisiointi`-moduulista

---

## Siivous ja optimointi (Vaihe 3)

### Kommenttien suomennus ó DataToACAD.bas

- Moduuliotsikko k√§√§nnetty suomeksi (Moduuli / Tarkoitus / Alkuper√§inen / Muokattu)
- Kaikkien funktioiden otsikkolohkot k√§√§nnetty: `CrsRefLink`, `get_filename`, `inch`, `makeFiles`, `MakeListNoLoopID`, `MakeLocFiles`
- Kaikki inline-kommentit k√§√§nnetty suomeksi:
  - Muuttujan esittelykommentit (lainausmerkki, tiedostokahva, silmukkalaskurit jne.)
  - Ohjauslogiikan kommentit (tiedostojen alustus, sulkumerkit, siivous, virheenk√§sittely)
  - Silmukkakommentit (`' tables` ‚Üí `' K√§yd√§√§n l√§pi kaikki taulukot` jne.)
- Englanninkieliset `MsgBox`-viestit k√§√§nnetty suomeksi

### Kommenttien suomennus ó GeneralCodes.bas

- Moduuliotsikko k√§√§nnetty suomeksi
- Julkisten muuttujien otsikko k√§√§nnetty: `' Public Variables for Revision Tracking'` ‚Üí suomeksi
- Kaikkien funktioiden otsikkolohkot k√§√§nnetty: `IsLoaded`, `HaeViimPaiva`, `Optiot`, `Positiot`, `Vaihekulma`, `MotKaapUh`, `LisaaNo`
- Mukautetun `Replace()`-funktion poisto-huomio tiivistetty suomeksi (4 rivi√§)
- Muuttujan esittelykommentit k√§√§nnetty: `Original length`, `Numeric value`, `Loop counter` ‚Üí suomeksi
- Kirjoitusvirheet korjattu: `vaihekulmna` ‚Üí `vaihekulman`, `k√§ytj√∂lle` ‚Üí `k√§yt√∂lle`

### Kommenttien suomennus ó USysCheck.bas

- `OPENFILENAME`-tyypin esittelykommentti k√§√§nnetty suomeksi
- `KORJATTU`-huomio laajennettu selitt√§m√§√§n `nSize As Long` vs. `LongPtr`-valinta
- Moduulitason tilamuuttujien kommentit k√§√§nnetty suomeksi
- `Show_last`- ja `Show_last_criteria`-funktioiden kommentit k√§√§nnetty

### Kommenttien suomennus ó Form_DBUsers.cls

- `WhosOn`-alirutiinin kaikki kommentit k√§√§nnetty:
  - Tietuetyypin esittely, tietokantapolun haku, tiedostopolun muodostus
  - Tiedoston olemassaolon tarkistus, tietuesilmukka, virheotsikot
- `MsgBox`-viestit k√§√§nnetty: `"Couldn't populate the list"` ‚Üí `"Tiedostoa ei l√∂ytynyt."`, `"Error: "` ‚Üí `"Virhe: "`

### Poistetut kehitysaikaiset merkinn√§t

- Poistettu `' Updated 2025-11-11: Added DAO prefix` -tyyppiset v√§liaikaismuistiinpanot `Form_DBUsers.cls`- ja `USysCheck.bas`-tiedostoista (3 kpl / tiedosto)

---

## AutoExec-korjaus (2026-03-04)

### Puuttuva SniffUser()-funktio ó USysCheck.bas

- **Ongelma:** AutoExec-makron `RunCode`-toiminto kutsui `=SniffUser()`, jota ei ollut MAINEQ-projektissa lainkaan ? Access antoi virheen 2425 k‰ynnistyksen yhteydess‰
- **Korjaus:** `SniffUser()`-funktio lis√§tty `USysCheck.bas`:iin

---

## Koodikatselmointikorjaukset (2026-03-09) ó Er‰t 1ñ4

### Kriittiset tietoturvakorjaukset

#### SQL-injektioiden poisto ó QueryDef-parametrisointi

- **Form_BeltConvFi_Subform.cls** (`Form_BeforeInsert`): Merkkijonoina rakennettu SQL-kysely korvattu `DAO.QueryDef`-pohjaisella parametrisoinnilla. Samalla korjattu `RecordCount`-haku `MAX(Suffix)`-funktioksi, joka on oikea myˆs poistojen j‰lkeen.
- **Form_CONVEYOR_FI_Subform.cls** (`Form_BeforeInsert`): Sama SQL-injektiokorjaus + `MAX(Suffix)`.
- **Form_DRIVES_FI_SubForm.cls** (`Form_BeforeInsert`): SQL-injektio + samanaikaisuusongelma (Race Condition) korjattu. `RecordCount + 1` ? `MAX(Suffix) + 1`.
- **Form_DRIVES_SubForm.cls** (`Form_BeforeInsert`): Sama korjaus kuin edell√§.

#### Haamuprosessien esto ó Excel COM -virheenk‰sittely

- **Form_MoottTilaus.cls** (`Command0_Click`): Lis‰tty `On Error GoTo ErrorHandler` + `Cleanup`-lohko. `Dir()`-tarkistus mallitiedostolle ennen Excel-k‰ynnistyst‰. Early Binding (`Excel.Application`) vaihdettu Late Bindingiiin (`Object`) yhteensopivuuden parantamiseksi. Excel-konstantit korvattu numeerisilla arvoilla (`xlTop` ? `-4160` jne.).
- **Report_MOOTTORIT.cls** (`Report_Open`): Lis‰tty t‰ydellinen `On Error GoTo ErrorHandler` + `Cleanup`-lohko. Poistettu k‰ytt‰m‰tˆn `FileSystemObject`. Late Binding + numeriset Excel-konstantit.
- **Report_PƒƒLAITTEET.cls** (`Report_Open`): Uudelleenkirjoitettu. Sis‰kk‰iset silmukat (`For Sar = 1 To 13`) korvattu `CopyFromRecordset`-metodilla (kertaluokan suorituskyky parannus). Lis‰tty `On Error GoTo`-virheenk‰sittely + `Dir()`-tarkistus. Late Binding + numeriset Excel-konstantit.

#### Merkkijonon korvausbugi ó Form_SiemensConstrCodeLastPosition.cls

- `LoppuKoodi1_DblClick` ja `LoppuKoodi2_DblClick`: `Replace(origType, ".", endCode)` ? `Replace(origType, ".", endCode, 1, 1)`. Est‰‰ tilanteen, jossa useampi piste moottorin tyyppikoodissa (esim. `1LA7.5.3`) korvautuu kaikki kerralla.
- `Null`-arvojen k‰sittely lis‰tty: `endCode` ja `origType` suojattu `IIf(IsNull(...))`.
- `MsgBox`-viestit muutettu nykyaikaisiksi (`@`-erottelusta `vbCrLf`-rivinvaihtoon).
- `DoCmd.Close` tarkennettu muotoon `DoCmd.Close acForm, "SiemensConstrCodeLastPosition"`.

### Suorituskykykorjaukset

#### CrsRefLink ó N+1-pullonkaula poistettu (DataToACAD.bas)

- Funktio avasi aiemmin uuden `DAO.Recordset`-olion **joka kutsulla** ja iteroi koko taulukon l‰pi lineaarisesti.
- Korjaus: Lis‰tty moduulitason `dictCrsRef As Object` (Scripting.Dictionary) ja `blnCrsRefLoaded As Boolean`. Taulukko ladataan `LoadCrsRefCache()`-alirutiinilla kerran ensimm‰isell‰ kutsulla; kaikki myˆhemm‰t haut ovat O(1).

### Siivous

#### Kuollut koodi poistettu ó Form_VDFManuf_Subform.cls

- `MuutaRevisio()`-alirutiinin sis‰ltˆ oli kokonaan kommentoitu ulos (7 rivi‰). Kommentoitu koodi poistettu; tyhj‰ runko s‰ilytetty viittauksia varten.

#### Datan menett‰misen riski poistettu ó Form_Linkkien vaihto.cls

- `DROP TABLE` suoritettiin aiemmin ennen kuin uuden tietokantatiedoston olemassaolo tarkistettiin.
- Korjaus: Lis‰tty `Dir(UusiPolku) = ""`-tarkistus ennen `DROP TABLE`. Jos tiedostoa ei lˆydy, k‰ytt‰j‰lle n‰ytet‰‰n selke‰ virheilmoitus ja kyseinen taulu ohitetaan ó alkuper‰inen linkki s‰ilyy ehj‰n‰.
- Lis‰tty `Taul.Close` / `Set Taul = Nothing` siivous funkition loppuun.
  - Hakee verkkok‰ytt‰j‰nimen olemassa olevalla `wu_GetUserName`-API:lla
  - Hakee tietokoneen nimen `Environ("COMPUTERNAME")`:lla (ei vaadi lis‰-API:a)
  - Kirjoittaa kirjautumistietueen `UsysUsers`-tauluun (kent‰t: NetworkUser, DBUser, ComputerName, LoginTime)
  - Virheet k‰sitell‰‰n hiljaisesti ó ei keskeyt‰ sovelluksen k‰ynnistyst‰

---

## Tiedostoyhteenveto

| Tiedosto | Vaihe 1ñ2 | Vaihe 3 | AutoExec-korjaus |
| --- | --- | --- | --- |
| `USysCheck.bas` | API PtrSafe + Long-korjaus | Kommentit suomeksi ? | `SniffUser()` lis‰tty ? |
| `DataToACAD.bas` | Debug.Print-rivit poistettu | Kommentit suomeksi ? | ó |
| `GeneralCodes.bas` | Option Explicit | Kommentit suomeksi ? | ó |
| `Form_DBUsers.cls` | Me.-etuliitteet | Kommentit suomeksi ? | ó |
| `MoottTilaus.bas` | Debug.Print-rivit poistettu | ó | ó |
| `Revisiointi.bas` | Debug.Print-rivi poistettu | ó | ó |
| Muut lomakkeet (6 kpl) | Option Explicit lis‰tty | ó | ó |

**Tiedosto:** `Access/MAINEQ/` (DataToACAD.bas, GeneralCodes.bas, USysCheck.bas, Form_DBUsers.cls, ja muut)
**P‰iv‰m‰‰r‰:** 2025-11-11 / 2025-11-12
**Tekij‰:** 64-bit migraatioprojekti (Pro gradu)

---

## Kriittiset muutokset (Vaihe 1ñ2)

### 64-bittinen API-korjaus ó USysCheck.bas

- `GetUserNameA` ja `GetComputerNameA`-funktioiden `Declare`-lauseet p‰ivitetty:
  - Lis‰tty `PtrSafe`-avainsana 64-bittist‰ Office/AutoCAD-ymp‰ristˆ‰ varten
  - `nSize`-parametri muutettu `LongPtr` ? `Long` (Win32 API kirjoittaa 4-tavuisen DWORD:in, ei 8-tavuisen)
- `GetLastError`- ja `SHGetPathFromIDList`-funktiot p‰ivitetty `PtrSafe`-m‰‰r‰‰ll‰

### Option Explicit ó kaikki moduulit

- Lis‰tty `Option Explicit` kaikkiin moduuleihin, joista se puuttui (6 lomaketta)

### Me.-etuliite ó Access-lomakkeet

- Kaikki lomakeohjauspyynnˆt p‰ivitetty `Me.`-etuliitteell‰ (`ControlName` ? `Me.ControlName`) implisiittisten viittausvirheiden est‰miseksi

### Debug.Print-rivien poisto

- Poistettu noin 50 `Debug.Print`-rivi‰ `DataToACAD.bas`-tiedostosta
- Poistettu 2 `Debug.Print`-rivi‰ `MoottTilaus`-moduulista
- Poistettu 1 `Debug.Print`-rivi `Revisiointi`-moduulista

---

## Siivous ja optimointi (Vaihe 3)

### Kommenttien suomennus ó DataToACAD.bas

- Moduuliotsikko k‰‰nnetty suomeksi (Moduuli / Tarkoitus / Alkuper‰inen / Muokattu)
- Kaikkien funktioiden otsikkolohkot k√§√§nnetty: `CrsRefLink`, `get_filename`, `inch`, `makeFiles`, `MakeListNoLoopID`, `MakeLocFiles`
- Kaikki inline-kommentit k‰‰nnetty suomeksi:
  - Muuttujan esittelykommentit (lainausmerkki, tiedostokahva, silmukkalaskurit jne.)
  - Ohjauslogiiikan kommentit (tiedostojen alustus, sulkumerkit, siivous, virheenk‰sittely)
  - Silmukkakommentit (`' tables` ? `' K‰yd‰‰n l‰pi kaikki taulukot` jne.)
- Englanninkieliset `MsgBox`-viestit k‰‰nnetty suomeksi

### Kommenttien suomennus ó GeneralCodes.bas

- Moduuliotsikko k‰‰nnetty suomeksi
- Julkisten muuttujien otsikko k‰‰nnetty: `' Public Variables for Revision Tracking'` ? suomeksi
- Kaikkien funktioiden otsikkolohkot k‰‰nnetty: `IsLoaded`, `HaeViimPaiva`, `Optiot`, `Positiot`, `Vaihekulma`, `MotKaapUh`, `LisaaNo`
- Mukautetun `Replace()`-funktion poisto-huomio tiivistetty suomeksi (4 rivi‰)
- Muuttujan esittelykommentit k‰‰nnetty: `Original length`, `Numeric value`, `Loop counter` ? suomeksi
- Kirjoitusvirhe korjattu: `vaihekulmna` ? `vaihekulman`, `k‰ytjˆlle` ? `k‰ytˆlle`

### Kommenttien suomennus ó USysCheck.bas

- `OPENFILENAME`-tyypin esittelykommentti k‰‰nnetty suomeksi
- `KORJATTU`-huomio laajennettu selitt‰m‰‰n `nSize As Long` vs. `LongPtr`-valinta
- Moduulitason tilamuuttujien kommentit k‰‰nnetty suomeksi
- `Show_last`- ja `Show_last_criteria`-funktioiden kommentit k‰‰nnetty

### Kommenttien suomennus ó Form_DBUsers.cls

- `WhosOn`-alirutiinin kaikki kommentit k‰‰nnetty:
  - Tietuetyypin esittely, tietokantapolun haku, tiedostopolun muodostus
  - Tiedoston olemassaolon tarkistus, tietuesilmukka, virheotsikot
- `MsgBox`-viestit k‰‰nnetty: `"Couldn't populate the list"` ? `"Tiedostoa ei lˆytynyt."`, `"Error: "` ? `"Virhe: "`

### Poistetut kehitysaikaiset merkinn‰t

- Poistettu `' Updated 2025-11-11: Added DAO prefix` -tyyppiset v‰liaikais-muistiinpanot `Form_DBUsers.cls`- ja `USysCheck.bas`-tiedostoista (3 kpl / tiedosto)

---

## Tiedostoyhteenveto

| Tiedosto | Vaihe 1ñ2 | Vaihe 3 |
| --- | --- | --- |
| `USysCheck.bas` | API PtrSafe + Long-korjaus | Kommentit suomeksi ? |
| `DataToACAD.bas` | Debug.Print-rivit poistettu | Kommentit suomeksi ? |
| `GeneralCodes.bas` | Option Explicit | Kommentit suomeksi ? |
| `Form_DBUsers.cls` | Me.-etuliitteet | Kommentit suomeksi ? |
| `MoottTilaus.bas` | Debug.Print-rivit poistettu | ó |
| `Revisiointi.bas` | Debug.Print-rivi poistettu | ó |
| Muut lomakkeet (6 kpl) | Option Explicit lis‰tty | ó |

---

## Code review -korjaukset (2026-11-10)

### Kriittiset muutokset

- `USysCheck.bas`, `Form_Revisiointi.cls`: `GetUserNameW` k‰ytt‰‰ `ByVal lpBuffer As LongPtr` + `StrPtr()`; tulos `Left$(‚Ä¶, BuffSize - 1)`. `#If VBA7` -haarat poistettu.
- `USysCheck.bas`: k‰ytt‰m‰tˆn `OPENFILENAME`-tyyppi ja `GetOpenFileName`-m‰‰rittely poistettu.
- `Form_DBUsers.cls`: `net send` korvattu `msg`-komennolla; kohde ja viesti puhdistetaan (komentoinjektion esto); nimenluku rajattu 32 tavuun.
- `DataToACAD.bas`: `MakeLocFiles` ja kovakoodatut agropm10-polut poistettu; `#n`-tiedostokahvat ‚Üí `FreeFile`; `makeFiles` palauttaa arvon (0 / -1); `GoTo`-ohitus korvattu `If`-rakenteella; sulkusulut kirjoitetaan aina (my√∂s `NoLoopIDTables`-tilassa); `Not Fields(ii).Name = Looppid` ‚Üí `ii <> Looppid`; `idsyst As Long`.
- `Form_KuvienGenerointi.cls`, `Form_GeneroiMoottorikuvat.cls`: virheenk‰sittely ja siivous (SingleDocumentMode palautetaan, recordsetit suljetaan); `Documents.Close` poistettu (k‰ytt‰j‰n muita kuvia ei suljeta); SQL-heittomerkit tuplataan; `Joukko.Count`-guard; poiminta transaktiossa; Santa Fe -polut poistettu, moottorikuvissa kansio valitaan FolderPickerill‰.
- `Form_Revisiointi.cls`: guardit (tyhj‰ lista, `InStr` = 0, tyhj‰ revisio), `LIKE`-suodattimen heittomerkit, merkinn‰t transaktiossa virheenk‰sittelyll‰.
- `GeneralCodes.bas`: `Vaihekulma`/`MotKaapUh` k‰sittelev‰t cos? = 1 (ei nollallajakoa); `Integer` ? `Long`/`Double`; `IsLoaded` palauttaa `Boolean`; `HaeViimPaiva` `InStrRev` + guardit; `Positiot` SQL-heittomerkit.
- `Form_MoottTilaus.cls`: Santa Fe -mallipolku korvattu tiedostonvalinnalla; `DROP TABLE` vain jos taulu on olemassa; Excel suljetaan virheess‰ ja `Calculation`/`ScreenUpdating` palautetaan; kuollut `Command3_Click` poistettu.
- `Form_EQUIPMENT.cls`, `Form_EQUIPMENT_FI.cls`: `EqSeq` haetaan numeerisella `Max(Val())`-haulla; SQL/suodatinten lainausmerkit suojattu; puuttuvat `On Error GoTo` -rivit lis‰tty; `DoMenuItem` ‚Üí `RunCommand acCmdFind` / `FilterOn`; FI-version fokusk‰sittelij‰t vertaavat nyt oikeaa painiketeksti‰ (`MUOKKAA_TEKSTI`).
- `Form_Motors_Subform.cls`: `Ilmoitus` `VbMsgBoxResult`, `vbYes`/`vbNo`, Peruuta-valinta k‰sitelty; toistuva koodi `VaihdaRunko`-apurutiiniin.
- `Form_SiemensConstrCodeLastPosition.cls`: kaksoiskoodi yhdistetty `KorvaaPiste`-rutiiniin, MotorTypes-lomakkeen olemassaolotarkistus.
- `Form_USysRevText.cls`: alustus vain kerran (`Form_Resize` ei en‰‰ nollaa syˆtteit‰); `ListIndex`/`Chr(Asc())`-guardit; `NaytaRev`-silmukan yl‰raja `ListCount - 1`; `;` estetty teksteist‰.

### Siivous ja optimointi

- Englanninkieliset virheilmoitukset ja kommentit suomennettu; muutoshistoria- ja kommentoitu koodi poistettu.
- Object-siivous (`Set Ö = Nothing`) ja virheenk‰sittelij‰t yhten‰istetty.
- `For ACAD Utility.bas`: `#If VBA7` poistettu, vain PtrSafe-m‰‰rittely.
- Muokkaamatta: `Form_UsysRevTextDrive.cls`, `Report_*`-luokat ja alilomakkeet (ei havaittu 64-bit- tai virheenk√§sittelyongelmia).
