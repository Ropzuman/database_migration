# Koodikatselmointiraportti: Erä 1 / 4 (Päälaitekanta)

## 📊 Yhteenveto

Ensimmäinen erä sisältää käyttöliittymän taustakoodeja (esim. `Form_BeltConvFi_Subform.cls`, `Form_DBUsers.cls`) ja järjestelmätason moduuleja (`DataToACAD.bas`, `USysCheck.bas`, `GeneralCodes.bas`). Koodi on selvästi toiminnallista ja hoitaa tärkeää liiketoimintalogiikkaa (kuten laitteiden revisiointia ja LISP-tiedostojen generointia AutoCADille). `Option Explicit` on esimerkillisesti käytössä kaikkialla. Koodista löytyy kuitenkin kriittisiä tietoturva-aukkoja (SQL- ja komentorivi-injektiot) sekä merkittäviä suorituskyvyn pullonkauloja.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

1. **OS Command Injection (Komentorivi-injektio)** `Form_DBUsers.cls`
   - Koodi käyttää suoraan käyttöjärjestelmän `Shell`-komentoa: `Call Shell("net send " & Me.NetworkName.Value & " """ & Viesti & """")`
   - **Riski:** Jos lomakkeen kentässä `NetworkName` tai `Viesti` on puolipisteitä (`&` tai `|` tai `;`), paha-aikeinen käyttäjä voi ajaa mitä tahansa Windows-komentoja laitteella. *Net send* -komento itsessään on myös poistunut uudemmista Windows-versioista (tilalla msg.exe), joten tämä voi kaatua virheeseen nykykäyttöjärjestelmissä.
2. **SQL-injektio (SQL Injection)** `Form_BeltConvFi_Subform.cls` ja `Form_CONVEYOR_FI_Subform.cls`
   - Esimerkki: `Set Taulukko = CurrentDb.OpenRecordset("SELECT Suffix From BeltConvProc WHERE Department = '" & Me.Parent.Department.Value & "' AND Eqtype='" & Me.Parent.EqType.Value & "' ...")`
   - **Riski:** Käyttöliittymän arvoja (Department, EqType) yhdistetään suoraan SQL-kyselymerkkijonoon. Jos käyttäjä syöttää kenttään heittomerkin (`'`), ohjelma joko kaatuu SQL-virheeseen tai mahdollistaa tahallisen koodin ajon kantaan.
3. **Resurssivuodot LISP-generoinnissa** `DataToACAD.bas`
   - `CrsRefLink`-funktio avaa ja sulkee `DAO.Recordset` -olion joka ikinen kerta kun sitä kutsutaan.
   - **Riski:** Erittäin vakava suorituskyvyn pullonkaula (N+1 -ongelma) ja mahdollinen muistivuoto/lukkiutuminen, kun taulun tietoja käsitellään toistuvissa isoissa silmukoissa.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

- **API-kutsujen puskurit (`USysCheck.bas`)**: Olet käyttänyt turvallista `PtrSafe`-muotoa, mikä on erinomaista. Huomioi kuitenkin, että `wu_GetUserName`-API palauttaa null-terminoidun merkkijonon (`Chr(0)`). Nykyinen logiikka (InStr) toimii, mutta on syytä varmistaa, ettei puskuria ylitetä.
- **Epäoptimaalinen merkkijononkäsittely (`GeneralCodes.bas - HaeViimPaiva`)**: Pitkien revisiotekstien (rivien) iterointi silmukalla ja `Right$`-komennolla yksittäinen merkki kerrallaan (`i = i + 1`) on tehotonta. Nyky-VBA:ssa voi ja kannattaa käyttää `Split(teksti, vbCrLf)`-funktiota.
- **Null-arvojen käsittely (`GeneralCodes.bas - LisaaNo`)**: Koodi käsittelee Variant-tyyppiä nollasyötteen varalta, mutta koodissa `Val(Tieto)` voi aiheuttaa ongelmia, jos syöte ei olekaan numeerinen.

## 💡 Parannusehdotukset (Ylläpidettävyys)

- **DRY-periaate**: Revisionseurantalogiiikka (`MuutaRevisio()`) on kopioitu identtisenä useisiin luokkiin (`Form_AUXILIARY_FI_SubForm.cls`, `Form_AUXILIARY_SubForm.cls`, jne.). Tämä kannattaa refaktoroida yhteen julkiseen moduuliin (esim. `GeneralCodes`), jonne välitetään Form-objekti parametrina.
- **DAO.QueryDef ja Parametrisointi**: Vältä SQL-merkkijonojen rakentamista VBA:ssa. Rakenna Accessiin tallennettu kysely (QueryDef) ja syötä sille parametrit `QueryDef.Parameters`-ominaisuuden kautta.
- **Välimuisti (Caching) Recordseteille**: Siirrä ristiviitetaulukon lataus välimuistiin. Lataa `CrsRefLisps`-taulu kerran esimerkiksi `Scripting.Dictionary` -objektiin moduulin yläosassa, jolloin haut ovat O(1) aikaluokkaa.

## 🛠️ Korjattu koodi

### 1. SQL-injektion estäminen QueryDef-olioilla (`Form_BeltConvFi_Subform.cls`)

Tämä korjaa tietoturvan, estää kaatumiset ja on vakaampi suorittaa:

```vba
Private Sub Form_BeforeInsert(Cancel As Integer)
    Dim qdf As DAO.QueryDef
    Dim Taulukko As DAO.Recordset
    Dim Uusi As String
    Dim SqlStr As String
    
    ' Käytetään parametrisoitua kyselyä turvallisuuden takaamiseksi (SQL Injection esto)
    SqlStr = "PARAMETERS [pDept] Text, [pEqType] Text, [pEqSeq] Text; " & _
             "SELECT Suffix FROM BeltConvProc " & _
             "WHERE Department = [pDept] AND Eqtype = [pEqType] AND EqSeq = [pEqSeq] " & _
             "ORDER BY Suffix DESC;"
             
    Set qdf = CurrentDb.CreateQueryDef("", SqlStr)
    
    ' Asetetaan parametrit turvallisesti lomakkeen arvoista
    qdf.Parameters("[pDept]") = Nz(Me.Parent.Department.Value, "")
    qdf.Parameters("[pEqType]") = Nz(Me.Parent.EqType.Value, "")
    qdf.Parameters("[pEqSeq]") = Nz(Me.Parent.EqSeq.Value, "")
    
    Set Taulukko = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Taulukko.EOF Then
        Uusi = "01"
    Else
        Uusi = Taulukko.Fields("Suffix").Value
        Uusi = Val(Uusi) + 1
        ' Format on puhtaampi tapa lisätä etunollia:
        Uusi = Format$(Uusi, "00")
    End If
    
    Me.Suffix.Value = Uusi
    
    Taulukko.Close
    Set Taulukko = Nothing
    Set qdf = Nothing
End Sub

2. Suorituskyky: Recordsetin N+1 pullonkaulan poistaminen välimuistilla (DataToACAD.bas)
VBA

' Lisää moduulin alkuun:
Private dictCrsRef As Object
Private CrsRefLoaded As Boolean

' Kutsutaan kerran ennen ajoa tai laiskasti funktion sisällä:
Private Sub LoadCrsRefCache()
    Dim DB As DAO.Database
    Dim tble As DAO.Recordset
    
    Set dictCrsRef = CreateObject("Scripting.Dictionary")
    Set DB = CurrentDb
    Set tble = DB.OpenRecordset("CrsRefLisps", dbOpenForwardOnly) ' Snapshot on nopeampi lukuun
    
    Do Until tble.EOF
        If Not dictCrsRef.Exists(tble!CrsRefID.Value) Then
            dictCrsRef.Add tble!CrsRefID.Value, tble!Lisp.Value
        End If
        tble.MoveNext
    Loop
    
    tble.Close
    Set tble = Nothing
    CrsRefLoaded = True
End Sub

Function CrsRefLink(tblnimi As String, teksti As String) As String
    On Error GoTo ErrorHandler
    
    If tblnimi = "CRSREF" Then
        ' Ladataan välimuistiin vain kerran!
        If Not CrsRefLoaded Then LoadCrsRefCache
        
        ' Nopea haku (O(1)) levyltä lukemisen (Recordset iterointi) sijaan
        If dictCrsRef.Exists(teksti) Then
            CrsRefLink = dictCrsRef(teksti)
        Else
            CrsRefLink = teksti
        End If
    Else
        CrsRefLink = teksti
    End If
    Exit Function

ErrorHandler:
    MsgBox "Error in CrsRefLink: " & Err.Description, vbCritical, "Lookup Error"
    CrsRefLink = teksti
End Function

# Koodikatselmointiraportti: Erä 2 / 4 (Päälaitekanta)

## 📊 Yhteenveto
Toinen erä sisältää merkittävän määrän toisteista käyttöliittymälogiikkaa laitteiden (Drives, Gears, Equipment) hallintaan. Koodi on perusrakenteeltaan tyypillistä Access-VBA:ta ja suorittaa tehtävänsä, mutta nojaa monessa paikassa hauraisiin ratkaisuihin. Mukana on kriittinen AutoCAD-integraatio, joka on valitettavasti sidottu yksittäiseen projektiin. SQL-injektioriski toistuu tässäkin erässä, minkä lisäksi tietokannan samanaikaiskäyttö (Concurrency) on suuressa riskissä huonon sekvenssinumerointilogiikan vuoksi.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

1. **Samanaikaisuusongelma ja Tietojen Korruptio (Race Condition)** `Form_DRIVES_FI_SubForm.cls` / `Form_DRIVES_SubForm.cls`
   - Koodi generoi uuden päälaitteen juoksevan numeron (Suffix) laskemalla olemassa olevat tietueet: `Uusi = "0" & CStr(Taulukko.RecordCount + 1)`.
   - **Riski:** Jos taulusta on poistettu yksikin tietue (esim. recordeja on 5, mutta suurin Suffix on "06"), koodi yrittää luoda uuden tunnuksen "06", mikä johtaa joko duplikaatteihin tai tietokannan virheeseen (Primary Key Violation). Monikäyttäjäympäristössä `RecordCount` ei ole koskaan turvallinen tapa määrittää seuraavaa ID:tä.
2. **SQL-injektio (SQL Injection)** `Form_DRIVES_FI_SubForm.cls`
   - Kuten edellisessä erässä, kysely rakennetaan suoraan merkkijonoja yhdistämällä: `CurrentDb.OpenRecordset("SELECT Suffix From DRIVES WHERE Department = '" & Me.Parent.Department.Value & "' ...")`. Tämä on altis tahallisille injektioille ja tahattomille syntaksivirheille (esim. heittomerkki osaston nimessä).
3. **Kovakoodatut tiedostopolut** `Form_GeneroiMoottorikuvat.cls`
   - Koodiin on kirjoitettu suoraan kovakoodattuja verkkolevyjen polkuja yksittäiselle projektille: `N:\whldata\Projekti\Santa Fe 220018\...`.
   - **Riski:** Tämä estää järjestelmän joustavan käytön tulevissa projekteissa ja tuotannon skaalaamisen. Koodia joutuu käsin puukottamaan jokaisen uuden projektin kohdalla, mikä on valtava ylläpitoriski.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

- **Taikanumerot ohjaimien nimissä** `Form_EQUIPMENT.cls`:
   - Koodi käyttää ohjaimille nimiä kuten `Button77_Click`, `Komento172_Click`, `Komento173_Click`. Nämä ovat täysin epäinformatiivisia ja tekevät koodin lukemisesta ja ylläpidosta lähes mahdotonta toiselle ohjelmoijalle.
- **Virheenkäsittely puuttuu resurssien vapautuksessa** `Form_GeneroiMoottorikuvat.cls`:
   - Vaikka koodi käyttää Late Bindingia (mikä on hyvä päivitys), virhetilanteessa (esim. AutoCAD kaatuu) olioita (`oAcad`, `oDoc`) ei välttämättä nollata kunnolla (Set Object = Nothing). Tämä jättää AutoCADin haamuprosesseja taustalle, jotka syövät koneen muistia ja vaativat lopulta Windowsin uudelleenkäynnistyksen.
- **VBA.Error$ käyttö** `Form_EQUIPMENT.cls`: `MsgBox VBA.Error$` antama virheilmoitus ei kerro loppukäyttäjälle yhtään mitään kontekstista (mikä epäonnistui ja miksi).

## 💡 Parannusehdotukset (Ylläpidettävyys)

- **DRY (Don't Repeat Yourself)**: `MuutaRevisio`-alirutiini on täysin identtisenä lähes jokaisessa erän .cls -tiedostossa (`Form_DRIVES_FI_SubForm`, `Form_Gears_Subform`, jne.). Tämä on klassinen "Copy-Paste" -ohjelmoinnin oire. Siirrä funktio yhteen globaaliin moduuliin, joka ottaa Form-objektin parametrina.
- **Max-funktio juoksevaan numerointiin**: Muuta SQL-kysely hakemaan Suffix-sarakkeen maksimiarvo (`MAX(Suffix)`) sen sijaan, että palauttaisit koko recordsetin. Se on satoja kertoja nopeampi ja huomattavasti luotettavampi tapa hakea uusi järjestysnumero.
- **Asetustiedosto tai taulu**: Siirrä `Form_GeneroiMoottorikuvat.cls` projektipolut Accessin asetus-tauluun (esim. `SysSettings`) tai lue ne asetuslomakkeelta. Älä koskaan säilytä lokaaleja hakemistopolkuja lähdekoodissa.

## 🛠️ Korjattu koodi

### 1. Turvallinen ja luotettava numerointi (Korjaus `Form_DRIVES_FI_SubForm.cls`)
Tämä ratkaisu korjaa Suffix-bugin käyttämällä kyselyssä MAX-funktiota ja Parameter-objektia turvallisuuden takaamiseksi.

```vba
Private Sub Form_BeforeInsert(Cancel As Integer)
    Dim qdf As DAO.QueryDef
    Dim Taulukko As DAO.Recordset
    Dim Uusi As String
    Dim SqlStr As String
    
    ' Hakee taulun SUURIMMAN suffixin tälle kombinaatiolle
    SqlStr = "PARAMETERS [pDept] Text, [pEqType] Text, [pEqSeq] Text; " & _
             "SELECT MAX(Suffix) AS MaxSuffix FROM DRIVES " & _
             "WHERE Department = [pDept] AND Eqtype = [pEqType] AND EqSeq = [pEqSeq];"
             
    Set qdf = CurrentDb.CreateQueryDef("", SqlStr)
    
    ' Suojaus SQL-injektiota vastaan
    qdf.Parameters("[pDept]") = Nz(Me.Parent.Department.Value, "")
    qdf.Parameters("[pEqType]") = Nz(Me.Parent.EqType.Value, "")
    qdf.Parameters("[pEqSeq]") = Nz(Me.Parent.EqSeq.Value, "")
    
    Set Taulukko = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Taulukko.EOF Or IsNull(Taulukko!MaxSuffix) Then
        Uusi = "01" ' Ei aiempia tietueita
    Else
        ' Lisätään 1 suurimpaan arvoon, ei tietueiden määrään!
        Uusi = CStr(Val(Taulukko!MaxSuffix) + 1)
        Uusi = Format$(Uusi, "00") ' Varmistaa 01, 02... 09 muodon
    End If
    
    Me.Suffix.Value = Uusi
    
    ' Siivous
    Taulukko.Close
    Set Taulukko = Nothing
    Set qdf = Nothing
End Sub

# Koodikatselmointiraportti: Erä 3 / 4 (Päälaitekanta)

## 📊 Yhteenveto
Kolmas erä sisältää mielenkiintoisia integraatioita muihin ohjelmistoihin, kuten Exceliin ja AutoCADiin, sekä dynaamista tietokannan linkitysten hallintaa (`Form_Linkkien vaihto.cls`). Koodissa on viitteitä siirtymisestä Late Binding -tekniikkaan, mikä on hyvä parannus yhteensopivuuden kannalta. Kuitenkin virheenkäsittely on näissä integraatioissa puutteellista, mikä aiheuttaa merkittäviä vakausriskejä (muistivuotoja ja haamuprosesseja). Lisäksi kovakoodatut tiedostopolut ja laitelogiikat vaikeuttavat järjestelmän ylläpitoa tulevaisuudessa.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

1. **Datan menettämisen riski (Taulujen linkitys)** `Form_Linkkien vaihto.cls`
   - Koodi pudottaa olemassa olevan taulun: `CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"`. Välittömästi tämän jälkeen suoritetaan uuden linkin luonti komennolla `DoCmd.TransferDatabase`, mutta välissä on `On Error Resume Next`.
   - **Riski:** Jos uusi tietokantatiedosto ei ole saatavilla (esim. verkkolevy on alhaalla tai polku väärä), ohjelma jatkaa suoritusta ja alkuperäinen taulu on pysyvästi poistettu järjestelmästä. Tämä kaataa koko sovelluksen muun toiminnan.
2. **Excel-sovelluksen haamuprosessit (Resource Leak)** `Form_MoottTilaus.cls`
   - Koodi avaa Excelin: `Set XL = CreateObject("Excel.Application")`, mutta siinä ei ole lainkaan virheenkäsittelyä (Error Handler).
   - **Riski:** Jos koodi kaatuu kesken silmukan (esim. datassa on virhe tai tiedostoa ei löydy), `XL.Quit` ja objektien tuhoaminen jäävät tekemättä. Tämä jättää `EXCEL.EXE` -prosessin pyörimään taustalle piilotettuna. Lopulta koneen muisti loppuu, kun raportteja ajetaan useita kertoja.
3. **Dynaamisten SQL-kyselyiden rakentaminen** `Form_Revisiointi.cls`
   - Koodi rakentaa monimutkaisia JOIN-kyselyitä yhdistämällä muuttujia merkkijonoksi: `CurrentDb.OpenRecordset("SELECT [" & ATaulu & "].*, [" & UTaulu & "].* FROM [" & ATaulu & "] RIGHT JOIN [" & UTaulu & "] ...")`.
   - **Riski:** Riippuen siitä, mistä `ATaulu` ja `UTaulu` saavat arvonsa, tämä on potentiaalinen SQL-injektiopiste.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

- **Kovakoodattu projektipolku toistuu:** Aivan kuten kakkoserässä, myös tässä erässä Excel-malli on kovakoodattu yhteen projektiin: `PWB = XL.Workbooks.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\LiiteMoottorienOstoehdotukseenTemplate.xls")`.
- **Vaarallinen korvausoperaatio:** `Form_SiemensConstrCodeLastPosition.cls` käyttää komentoa `Replace(Forms!motortypes!MotorType, ".", endCode)`. Jos moottorin tyyppikoodissa (esim. `1LA7.5.3`) sattuu olemaan vahingossa useampia pisteitä, `Replace` korvaa *ne kaikki* globaalisti, mikä tuhoaa alkuperäisen datan eheyden.
- **Liiketoimintalogiikka käyttöliittymässä:** `Form_Motors_Subform.cls` ja `Form_MotorTypes.cls` sisältävät kovakoodattua logiikkaa sille, miten "SIEMENS" -merkkiset laitteet ja asennustavat ("B3") käyttäytyvät. Tämä on erittäin vaikeasti skaalautuva ratkaisu, jos järjestelmään tulee uusia poikkeusvalmistajia.

## 💡 Parannusehdotukset (Ylläpidettävyys)

- **Transaktiot taulujen päivityksessä:** Jos joudut muokkaamaan linkkejä tai tietokannan rakennetta `Form_Linkkien vaihto.cls` -luokassa, tee tarkistukset tiedoston olemassaolosta `Dir()` -funktiolla *ennen* `DROP TABLE` -komentoa.
- **Virheenkäsittely COM-objekteille (Clean Code):** Aina kun avaat ulkoisen ohjelman (Excel, AutoCAD, Word), käytä `On Error GoTo ErrorHandler` -rakennetta. ErrorHandler-lohkossa tulee varmistaa, että sovellus suljetaan `App.Quit` ja muuttujat nollataan `Set App = Nothing`.
- **InstrRev tai spesifi Replace:** Muuta `Form_SiemensConstrCodeLastPosition.cls` toimimaan niin, että se etsii pisteen paikan `InStr`:llä ja muuttaa vain yhden merkin, tai aseta `Replace`-funktiolle argumentit niin, että se korvaa vain ensimmäisen/viimeisen osuman.

## 🛠️ Korjattu koodi

### 1. Virheenkestävä Excel-integraatio (`Form_MoottTilaus.cls`)
Tämä korjaus lisää olennaisen virheenkäsittelyn, joka estää haamuprosessit ja tarkistaa template-tiedoston olemassaolon.

```vba
Private Sub Command0_Click()
    On Error GoTo ErrorHandler
    
    Dim PWB As Excel.Workbook
    Dim UWB As Excel.Workbook
    Dim XL As Excel.Application
    Dim i As Integer
    Dim Rivi As Integer
    Dim TemplatePath As String
    
    ' Polku mieluiten asetus-taulusta, mutta tässä esimerkki tarkistuksesta
    TemplatePath = "N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\LiiteMoottorienOstoehdotukseenTemplate.xls"
    
    If Dir(TemplatePath) = "" Then
        MsgBox "Excel-mallitiedostoa ei löydy polusta:" & vbCrLf & TemplatePath, vbCritical, "Virhe"
        Exit Sub
    End If
    
    Set XL = CreateObject("Excel.Application")
    ' Piilotetaan Excel ajon ajaksi (nopeuttaa huomattavasti)
    XL.Visible = False
    XL.ScreenUpdating = False
    
    Set PWB = XL.Workbooks.Open(TemplatePath)
    PWB.Sheets.Copy
    Set UWB = XL.ActiveWorkbook
    PWB.Close False
    
    i = 1
    Rivi = 6
    
    If Not Me.RecordsetClone.EOF Then
        Me.RecordsetClone.MoveFirst
        Do While Not Me.RecordsetClone.EOF
            With UWB.ActiveSheet
                .Cells(Rivi, 1).Value = i
                .Cells(Rivi, 2).Value = Me.RecordsetClone.Fields(0).Value
                .Cells(Rivi, 3).Value = Me.RecordsetClone.Fields(1).Value & Me.RecordsetClone.Fields(2).Value
                ' ... jne ...
            End With
            Rivi = Rivi + 1
            i = i + 1
            Me.RecordsetClone.MoveNext
        Loop
        
        ' Formatoinnit
        With XL.Range(UWB.ActiveSheet.Cells(6, 1), UWB.ActiveSheet.Cells(Rivi - 1, 10)).Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    ' Näytetään tulos käyttäjälle
    XL.ScreenUpdating = True
    XL.Visible = True
    AppActivate XL.Caption

Cleanup:
    On Error Resume Next
    ' Nämä on tärkeää vapauttaa, jotta ei synny muistivuotoja
    Set PWB = Nothing
    Set UWB = Nothing
    Set XL = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe Excel-raportin luonnissa: " & Err.Description, vbCritical, "Virhe " & Err.Number
    ' Jos Excel kaatuu, yritetään tappaa prosessi siististi
    If Not XL Is Nothing Then
        XL.Quit
    End If
    Resume Cleanup
End Sub

2. Turvallinen merkin vaihtaminen moottorin tyyppiin (Form_SiemensConstrCodeLastPosition.cls)

Estetään tilanne, jossa vahingossa vaihdetaan useita pisteitä koodin sisältä.
VBA

Private Sub LoppuKoodi1_DblClick(Cancel As Integer)
    Dim endCode As String
    Dim UusiTyyppi As String
    Dim dotPos As Integer
    Dim origType As String
    
    endCode = Nz(Me.LoppuKoodi1, "")
    origType = Nz(Forms!motortypes!MotorType, "")
    
    ' Etsitään pisteen paikka
    dotPos = InStr(1, origType, ".")
    
    If dotPos > 0 Then
        ' Korvataan VAIN ensimmäinen löydetty piste käyttämällä Left ja Mid funktioita
        ' TAI käytetään Replace funktiota määrittämällä osumien määrä yhteen (Count:=1)
        UusiTyyppi = Replace(origType, ".", endCode, 1, 1)
        
        Forms!motortypes!MotorType = UusiTyyppi
        
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
    Else
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
        MsgBox "Koodia ei muutettu!" & vbCrLf & "Tyyppimerkissä ei ollut pistettä." & vbCrLf & "Tarkista luettelosta!", vbExclamation, "Huomio"
    End If
End Sub

# Koodikatselmointiraportti: Erä 4 / 4 (Päälaitekanta)

## 📊 Yhteenveto
Neljäs erä sisältää käyttöliittymän alilomakkeita, revisiohistoriaa käsitteleviä lomakkeita sekä järjestelmän Excel-raportointimoduuleja. Raporttien luonti Exceliin on toteutettu Late Bindingia myötäillen, ja koodissa näkyy hyviä yrityksiä tilapalkin (SysCmd) hyödyntämiseksi käyttäjäkokemuksen parantamiseksi. Kuitenkin arkkitehtuurissa on vakavia suorituskyvyn pullonkauloja massadatan viennissä, ja sovelluksen vakaus on vaarassa puuttuvan virheenkäsittelyn takia.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

1. **Excel-sovelluksen haamuprosessit ja muistivuodot** (`Report_MOOTTORIT.cls`, `Report_PÄÄLAITTEET.cls`)
   - Olet ottanut käyttöön `CreateObject("Excel.Application")`, mutta raporteista puuttuu täysin virheenkäsittely (`On Error GoTo ...`).
   - **Riski:** Jos raportin luonnin aikana tapahtuu virhe (esim. Excel-mallitiedosto on lukittu, levy on täynnä tai verkkoyhteys pätkäisee), VBA-koodi kaatuu kesken ajon. Koska `xlApp.ScreenUpdating = False` on asetettu, piilotettu `EXCEL.EXE` jää ikuisesti pyörimään tietokoneen taustaprosesseihin varaamalla muistia, kunnes koko tietokone käynnistetään uudelleen.
2. **Datan korruptoituminen merkkijonon jäsennyksessä (Delimiter Collision)** (`Form_USysRevText.cls`)
   - Revisiohistoria tallennetaan yhteen pitkään merkkijonoon, joka erotellaan välilyönneillä ja kauttaviivoilla: `Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value...`. Purkaminen (`NaytaRev`) tapahtuu etsimällä `/`-merkkiä (`InStr(Tieto, "/")`).
   - **Riski:** Jos käyttäjä sattuu kirjoittamaan *Kuvaus* (Description) tai *Tekijä* (Drawn) -kenttään kauttaviivan (esim. "Muutos A/B"), koko tallennettu revisiohistoria pirstaloituu ja sen lukeminen kaatuu virheeseen, koska koodi jakaa kentät väärin.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

- **Massadatan suorituskyky (O(N*M) pullonkaula)** `Report_PÄÄLAITTEET.cls`:
   - Exceliin vienti on toteutettu kahdella sisäkkäisellä silmukalla (Rivi kerrallaan ja Sarake kerrallaan: `WB.Sheets(1).Cells(Rivi, Sar).Value = Tiedot.Fields(Sar - 1).Value`).
   - Jos laitteita on satoja tai tuhansia, tämä kestää minuutteja. Excel Interop -kutsut (COM-rajapinnan yli) ovat äärimmäisen hitaita. Accessissa tulisi käyttää Recordsetin suorakirjoitusta (`CopyFromRecordset`), joka on satoja kertoja nopeampi.
- **Kovakoodatut tiedostopolut** (`Report_MOOTTORIT.cls`): Projektin juurihakemisto ("Santa Fe 220018") on jälleen koodattu suoraan raporttiluokan sisään, mikä estää järjestelmän uudelleenkäytön ilman ohjelmistokehittäjän koodimuutoksia.

## 💡 Parannusehdotukset (Ylläpidettävyys)

- **"Dead Code" (Kuollut koodi) poistaminen**: Lomakkeessa `Form_VDFManuf_Subform.cls` koko `MuutaRevisio()`-aliohjelman sisältö on kommentoitu ulos. Kommentoitu koodi tulisi aina poistaa versionhallinnan roskaantumisen estämiseksi.
- **Datan rakenne (Normalisointi)**: Sen sijaan, että revisioita varastoidaan pilkottuina merkkijonoina tekstikenttään (`Form_USysRevText.cls`), ne tulisi tallentaa omaan relaatiotauluunsa (esim. `RevisionHistory` -taulu, jolla on FK päälaitteeseen). Tämä ratkaisisi kauttaviiva-ongelmat lopullisesti.
- **Tyylien formatointi Excelissä**: Solujen formatointi rivi kerrallaan ei ole tarpeen. Määritä Excel-templatessa (MotorTEMPLATE.xls) alue valmiiksi taulukoksi (Excel Table / ListObject), jolloin viivat ja tyylit tulevat automaattisesti oikein kun uutta dataa pudotetaan.

## 🛠️ Korjattu koodi

### 1. Suorituskykyinen ja vakaa Excel-vienti (`Report_PÄÄLAITTEET.cls`)
Tämä refaktoroitu versio käyttää nopeaa `CopyFromRecordset`-metodia ja sisältää ehdottoman tärkeän virheenkäsittelyn muistivuotojen estämiseksi.

```vba
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    Dim xlApp As Excel.Application
    Dim WB As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim Tiedot As DAO.Recordset
    Dim Rivi As Integer
    Dim TemplatePath As String
    
    ' Polku asetuksista
    TemplatePath = "N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MainEqTEMPLATE.xls"
    Set Tiedot = CurrentDb.OpenRecordset("_qryMAINEQCust", dbOpenSnapshot)
    
    If Tiedot.EOF Then
        MsgBox "Ei vietävää dataa.", vbInformation
        Cancel = True
        Exit Sub
    End If
    
    If MsgBox("Haluatko generoida Excel-listan?", vbYesNo + vbQuestion, "Päälaiteluettelo") = vbYes Then
        SysCmd acSysCmdSetStatus, "Käynnistää Exceliä..."
        
        Set xlApp = CreateObject("Excel.Application")
        xlApp.ScreenUpdating = False
        Set WB = xlApp.Workbooks.Open(TemplatePath, , True)
        Set WS = WB.Sheets(1)
        
        SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin (CopyFromRecordset)..."
        Rivi = 9
        WS.Cells.NumberFormat = "@"
        
        ' YLIVOIMAISESTI NOPEIN TAPA: Pudottaa koko Recordsetin kerralla ilman silmukoita!
        WS.Cells(Rivi, 1).CopyFromRecordset Tiedot
        
        ' Haetaan viimeinen rivi muotoiluja varten
        If Not Tiedot.EOF Then Tiedot.MoveLast
        Dim ViimeinenRivi As Integer
        ViimeinenRivi = Rivi + Tiedot.RecordCount - 1
        
        ' Formatoidaan kerralla koko alue
        With WS.Range(WS.Cells(Rivi, 1), WS.Cells(ViimeinenRivi, 13))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = xlAutomatic
            .RowHeight = 18.75
            .VerticalAlignment = xlVAlignTop
            .HorizontalAlignment = xlCenter
        End With
        
        ' Näytetään tulos
        xlApp.ScreenUpdating = True
        xlApp.Visible = True
        AppActivate xlApp.Caption
    End If

Cleanup:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    If Not Tiedot Is Nothing Then Tiedot.Close
    Set Tiedot = Nothing
    Set WS = Nothing
    Set WB = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe havaittu Excel-viennissä: " & Err.Description, vbCritical, "Virhe " & Err.Number
    ' Siivotaan haamuprosessi, jos kaatuu ennen näkymistä
    If Not xlApp Is Nothing Then
        If Not xlApp.Visible Then xlApp.Quit
    End If
    Resume Cleanup
End Sub
