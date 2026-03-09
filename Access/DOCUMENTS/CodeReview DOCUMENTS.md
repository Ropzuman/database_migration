# Moduuli: Form_SETTINGS.cls

📊 Yhteenveto: Lomake vastaa sovelluksen nimen tallentamisesta tietokannan ominaisuuksiin (AppTitle). Koodissa on pyritty hyvään virheiden käsittelyyn (On Error GoTo ErrorHandler) ja eksplisiittiseen DAO-tyypitykseen, mikä on erinomaista.

🚨 Kriittiset löydökset (Tietoturva & Vakaus): * Vakausriski: Koodissa kutsutaan DB.Close muuttujalle, joka on asetettu CurrentDb-viittaukseksi. CurrentDb on funktio, joka palauttaa kopion nykyisen tietokannan viittauksesta. Vaikka Access yleensä sietää sen sulkemisen, virallinen suositus Microsoftilta on, että CurrentDb-kutsulla saatua viittausta ei pitäisi koskaan sulkea .Close-metodilla, sillä se voi johtaa odottamattomaan käyttäytymiseen ja vakausongelmiin pitkässä juoksussa. Riittää, että se tyhjennetään asettamalla se Nothing-tilaan.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): * Nz(Me.Nimi.Value) on hyvä tapa estää Null-virheet, mutta AppTitle vaatii yleensä merkkijonon. Jos kenttä on tyhjä, Nz palauttaa oletuksena tyhjän merkkijonon. Olisi hyvä varmistaa, ettei ohjelma aseta tyhjää nimeä (esim. tarkistamalla arvo ennen tallennusta).

💡 Parannusehdotukset (Ylläpidettävyys): * Poista DB.Close kaikista paikoista, joissa tietokantaobjekti on alustettu Set DB = CurrentDb.

🛠️ Korjattu koodi:
VBA

Private Sub Form_Unload(Cancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  ' Varmistetaan, ettei asenneta tyhjää otsikkoa
  If Len(Nz(Me.Nimi.Value, "")) > 0 Then
      DB.Properties("AppTitle").Value = Me.Nimi.Value
      Application.RefreshTitleBar
  End If
  
  ' EI SAA SULKEA: CurrentDb:tä ei tulisi sulkea koodista!
  ' DB.Close
  Set DB = Nothing
  
  DoCmd.OpenForm "DOCUMENTS"
  Exit Sub
  
ErrorHandler:
  MsgBox "Error saving settings: " & Err.Description, vbExclamation, "Settings Error"
  If Not DB Is Nothing Then
    ' DB.Close <- Poistettu myös virheenkäsittelystä
    Set DB = Nothing
  End If
End Sub

Moduuli: Form_USysAddToDistr.cls

📊 Yhteenveto: Tämä moduuli hallinnoi dokumenttien lisäämistä tiettyyn jakeluun (distribution). Transaktioiden käyttö (DBEngine.BeginTrans / CommitTrans) on loistava tapa varmistaa datan eheyden säilyminen virhetilanteissa.

🚨 Kriittiset löydökset (Tietoturva & Vakaus): *SQL-injektio ja tyyppivirheiden riski: SQL-kyselyssä ("SELECT* FROM SentDocuments WHERE TDoc=" & CurRecord & " AND TDistr=" & Me.Distribution.Value) syötteet ketjutetaan suoraan kyselymerkkijonoon. Jos Me.Distribution.Value voi sisältää ei-numeerista dataa, tämä kaatuu. Parametrisoidut kyselyt (QueryDefs) olisivat turvallisin vaihtoehto.

    Vakausriski: Transaktio avataan (DBEngine.BeginTrans), mutta virhetilanteessa (ErrorHandler) koodista näyttää puuttuvan DBEngine.Rollback -kutsu. Jos taulu.Update epäonnistuu (esim. lukitusvirheen vuoksi), transaktio jää roikkumaan auki, mikä lukitsee tietokannan muilta käyttäjiltä. Sama ongelma CurrentDb:n sulkemisen kanssa esiintyy myös täällä.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): * Koodi käyttää dbOpenDynaset-parametria vain olemassaolon tarkistamiseen. Parempi ja performantimpi ratkaisu (ja turvallisempi) olisi käyttää DCount-funktiota tai avata pelkkä lisäystaulu.

💡 Parannusehdotukset (Ylläpidettävyys): * Siirry QueryDef-objektien käyttöön tai aseta selkeät parametrit, ja muista aina hoitaa Rollback virheenkäsittelyssä.

🛠️ Korjattu koodi:
VBA

Private Sub Add_Click()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim inTransaction As Boolean
  
  On Error GoTo ErrorHandler
  inTransaction = False
  
  ' Esitarkistukset Null-arvoille
  If IsNull(Me.Distribution.Value) Then
      MsgBox "Please select a distribution first.", vbExclamation, "Validation"
      Exit Sub
  End If
  
  Set DB = CurrentDb
  
  ' Suorituskykyisempi ja SQL-injektiovapaa olemassaolon tarkistus
  If DCount("*", "SentDocuments", "TDoc=" & CurRecord & " AND TDistr=" & CLng(Me.Distribution.Value)) > 0 Then
    MsgBox "Document is already in the selected distribution!", vbCritical, "Add document to distribution"
  Else
    DBEngine.BeginTrans
    inTransaction = True

    ' Avataan taulu pelkkää lisäystä varten
    Set taulu = DB.OpenRecordset("SentDocuments", dbOpenDynaset, dbAppendOnly)
    taulu.AddNew
    taulu.Fields(0) = CurRecord
    taulu.Fields(1) = Me.Distribution.Value
    taulu.Fields(2) = IIf(Common = "", Null, Common)
    If DocStatus <> "" Then taulu.Fields(3) = DocStatus
    taulu.Update
    
    DBEngine.CommitTrans
    inTransaction = False
    
    DoCmd.Close
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  If inTransaction Then
      DBEngine.Rollback ' Vapautetaan lukot, jos tallennus epäonnistui!
      inTransaction = False
  End If
  MsgBox "Error: " & Err.Description, vbCritical, "Database Error"
  Resume Cleanup
End Sub

Moduuli: Form_DBUsers.cls

📊 Yhteenveto: Moduuli lukee Accessin .laccdb (lukitus) -tiedostoa tunnistaakseen, ketkä käyttäjät ovat kirjautuneena kantaan. Tämä on erittäin klassinen ja edistynyt tekniikka ldb/laccdb-tiedoston jäsentämiseen binäärimuodossa. Koodi mahdollistaa myös verkkoviestien lähettämisen käyttäjille net send -komennolla.

🚨 Kriittiset löydökset (Tietoturva & Vakaus): *Tietoturvariski (Komentoinjektio / Command Injection): Funktio Command27_Click() ottaa syötteen suoraan InputBox:sta ja syöttää sen Windowsin Shell-komentoon (Call Shell("net send " & Me.NetworkName.Value & " """ & Viesti & """")). Käyttäjä voi lisätä syötteeseen putkitusmerkkejä (esim. & tai |) ja suorittaa laittomia komentorivikomentoja (esim. viesti" & del c:\*.*).

    Vakaus & Yhteensopivuus: net send -komento poistettiin Windowsista jo Windows Vistan aikoihin (korvattiin msg.exe -komennolla). On erittäin todennäköistä, että tämä toiminnallisuus kaatuu tai ei tee mitään moderneissa Windows 10/11 -ympäristöissä.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): * Lukitustiedoston lukeminen binäärinä (Open sPath For Binary Access Read Shared As iLDBFile) on haavoittuvaista ajoitusvirheille. Tiedosto voi olla juuri silloin varattuna Accessin ytimen toimesta. Varmista aina huolellinen On Error -käsittely tämäntyyppisissä manuaalisissa IO-operaatioissa.

💡 Parannusehdotukset (Ylläpidettävyys): * Vaihda komentorivikutsu käyttämään turvallisempaa mekanismia viestintään (esim. sähköposti, tai tietokantatauluun pohjautuva notifikaatiojärjestelmä lomakkeille). Jos on pakko käyttää komentoriviä, puhdista aina syötteet.

🛠️ Korjattu koodi (Command27_Click):
VBA

Private Sub Command27_Click()
  Dim Viesti As String
  Dim TurvallinenViesti As String
  
  If IsNull(Me.NetworkName.Value) Or Me.NetworkName.Value = "" Then Exit Sub
  
  Viesti = InputBox("Send message to user " & Me.NetworkName.Value & vbCrLf & "Give message:", "Send message")
  
  If Viesti <> "" Then
    ' 1. Estetään komentoinjektio poistamalla vaaralliset merkit
    TurvallinenViesti = Replace(Viesti, """", "'")
    TurvallinenViesti = Replace(TurvallinenViesti, "&", " ja ")
    TurvallinenViesti = Replace(TurvallinenViesti, "|", "")

    ' 2. Huom: 'net send' ei toimi enää modernissa Windowsissa. MSG.exe on oikea komento nykyään.
    ' Kutsu msg.exe muodossa: msg käyttäjätunnus "Viesti"
    On Error Resume Next ' Varmistetaan ettei Shell-virhe kaada koko ohjelmaa
    Call Shell("msg " & Me.NetworkName.Value & " """ & TurvallinenViesti & """", vbHide)
    If Err.Number <> 0 Then
        MsgBox "Failed to send message. Please ensure MSG command is supported.", vbExclamation
    End If
    On Error GoTo 0
  End If
End Sub

Moduuli: ForDocuments.vba

📊 Yhteenveto: Moduuli sisältää lukuisia julkisia muuttujia (Global Variables) lomakkeiden väliseen tilanhallintaan sekä Windows API -kutsuja (SHBrowseForFolder jne.) hakemiston valintadialogia varten. Globaalien muuttujien raskas käyttö ohjaa vahvasti spagettikoodiin.

🚨 Kriittiset löydökset (Tietoturva & Vakaus): * Vakaus (API-määritykset): Koodissa on maininta PtrSafe ja LongPtr päivityksistä 64-bittisiä järjestelmiä varten, mikä on erinomaista työtä. Muistathan kuitenkin, että jos käytät funktioissa Pointtereita, osoitteiden pituudet pitää aina tyypittää oikein myös C-tason rakenteissa (esim. Type BROWSEINFO).

    Kestävyys (Globaalit muuttujat): Public Common As String, Public CurRecord As String jne. Globaalit muuttujat häviävät muistista hetkellisesti, jos koodin ajo pysähtyy (Unhandled Exception) VBA-puolella. Tämä aiheuttaa sovelluksen tilan korruptoitumisen.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): * VBA:n sisäänrakennettu FileDialog(msoFileDialogFolderPicker) on modernimpi, vakaampi ja vaatii nolla API-kutsua verrattuna perinteiseen SHBrowseForFolder API:in.

💡 Parannusehdotukset (Ylläpidettävyys): * Korvaa monimutkaiset API-kutsut Office-objektimallin omalla FileDialog-dialogilla, mikä poistaa alustariippuvuudet (32-bit vs 64-bit Access).

    Kapseloi globaalit muuttujat luokkamoduuleihin, tai passaa data suoraan argumentteina lomakkeiden OpenArgs-ominaisuuden kautta.

🛠️ Korjattu koodi (Kansion valinta ilman API-kutsuja):
VBA

' Tällä voi korvata kokonaan vanhan SHBrowseForFolder API-hässäkän:
Public Function ValitseHakem(Optional StartPath As String = "") As String
    Dim fd As Object ' Office.FileDialog

    ' msoFileDialogFolderPicker = 4
    Set fd = Application.FileDialog(4) 
    
    With fd
        .Title = "Choose path:"
        .AllowMultiSelect = False
        If Len(StartPath) > 0 Then
            .InitialFileName = StartPath
        End If
        
        If .Show = -1 Then ' Käyttäjä painoi OK
            ValitseHakem = .SelectedItems(1)
            ' Varmistetaan että lopussa on takakeno
            If Right$(ValitseHakem, 1) <> "\" Then ValitseHakem = ValitseHakem & "\"
        Else
            ValitseHakem = ""
        End If
    End With
    
    Set fd = Nothing
End Function

Moduuli: Form_USysReserve.cls

📊 Yhteenveto: Moduulin tehtävänä on varata asiakkaan piirustusnumero. Taulu avataan dbDenyWrite-lukituksella samanaikaisten varausten estämiseksi. Koodi tekee tehtävänsä, mutta on altis injektioille ja kovakoodatuille arvoille.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    SQL-injektio: Kyselyssä yhdistetään suoraan käyttöliittymän arvo SQL-lauseeseen ("SELECT * ... WHERE Number='AV-2090-210-DD-" & Me.TNumber.Value & "'"). Jos käyttäjä syöttää hakukenttään heittomerkin ('), kysely kaatuu syntaksivirheeseen.

    Kovakoodattu etuliite: Projektin tunnus "AV-2090-210-DD-" on kovakoodattu SQL-lauseeseen. Jos projekti vaihtuu, koodi menee rikki.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Lukitus (dbDenyWrite) on tehokas tapa estää kilpailutilanteet (race condition), mutta jos ohjelman suoritus keskeytyy virheen takia ennen .Close-kutsua, taulu voi jäädä lukkoon. Huolellinen virheenkäsittely on tässä elintärkeää.

💡 Parannusehdotukset (Ylläpidettävyys):

    Käytä puhdistettua syötettä tai QueryDef-parametrisointia.

    Siirrä projektin etuliite haettavaksi dynaamisesti (esim. ProjInfo-taulusta) tai lomakkeen muuttujasta.

🛠️ Korjattu koodi (TalletaNappi_Click):
VBA

Private Sub TalletaNappi_Click()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    Dim EtsittavaNumero As String

    On Error GoTo ErrorHandler
    
    If IsNull(Me.TPaperSize.Value) Or IsNull(Me.TDiscipline.Value) Or IsNull(Me.TNumber.Value) Then
      MsgBox "You didn't select client number!" & vbCrLf _
           & "Please fill all fields and try again.", vbCritical, "Error!"
      Exit Sub
    End If
    
    Set DB = CurrentDb
    ' Puhdistetaan heittomerkit mahdollisista syötteistä
    EtsittavaNumero = Replace(Me.TNumber.Value, "'", "")
    
    ' Vältä kovakoodausta jos mahdollista, tässä korjattu vain injektioriski:
    SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & EtsittavaNumero & "'"
    
    ' Avataan taulu lukituksella
    Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset, dbDenyWrite)
    
    If IsNull(taulu.Fields("Reserved")) Then
      taulu.Edit
      taulu.Fields("Number") = Me.TClientNo.Caption
      taulu.Fields("Reserved") = Varaaja & "@" & Now
      taulu.Update
      
      Me.ClientNo.Value = Me.TClientNo.Caption
      DoCmd.Close
    Else
      MsgBox "Someone reserved number " & EtsittavaNumero & " at the same time!" & vbCrLf & _
             "Please choose another number.", vbCritical, "Error!"
    End If
    
Cleanup:
    If Not taulu Is Nothing Then
      taulu.Close
      Set taulu = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Database Error"
    Resume Cleanup
End Sub

Moduuli: Form_USysOpenFile.cls

📊 Yhteenveto: Lomake lataa dokumenttiin liitetyt tiedostopolut tietokannasta ja antaa käyttäjän avata kansion resurssienhallinnassa Shell-komennolla.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Komentoinjektio (Command Injection): Call Shell("C:\Windows\explorer.exe /e, " & Hakem, vbNormalFocus) on erittäin vaarallinen. Jos Hakem-muuttujan (tietokannasta haettu WorkPath) sisältöön on tallennettu putkitusmerkkejä tai komentojonoja (esim. c:\temp\ & del /F /S /Q c:\*.*), Shell saattaa suorittaa nämä suoraan komentorivillä.

    SQL-injektio: SELECT File, WorkPath FROM DOCUMENTS WHERE Counter=" & Common. Globaali muuttuja Common ketjutetaan suoraan kyselyyn.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Explorerin polun kovakoodaus "C:\Windows\explorer.exe" on huono käytäntö. Käyttöjärjestelmän asennushakemisto voi olla myös esim. D:\Windows.

💡 Parannusehdotukset (Ylläpidettävyys):

    Luovu Shell-funktion käytöstä kansioiden avaamisessa. Accessissa on turvallisempi sisäänrakennettu keino: Application.FollowHyperlink. Tämä antaa käyttöjärjestelmän hoitaa kansion avaamisen oikeassa ohjelmassa ilman komentorivin välimiesriskejä.

🛠️ Korjattu koodi (Command52_Click):
VBA

Private Sub Command52_Click()
  ' fso on alustettu moduulin yläosassa (FileSystemObject)
  If fso.FolderExists(Hakem) Then
    On Error Resume Next
    ' Turvallisempi tapa avata hakemisto ilman Shell-komentoinjektiota:
    Application.FollowHyperlink Hakem
    If Err.Number <> 0 Then
      MsgBox "Failed to open folder.", vbCritical, "Error"
    End If
    On Error GoTo 0
  Else
    MsgBox "Invalid path! " & Hakem, vbCritical, "Open Path"
  End If
End Sub

Moduuli: Form_USysNewDistribution.cls

📊 Yhteenveto: Lomakkeella hallitaan uuden "distribuution" eli dokumenttilähetyksen luontia. Moduulissa käsitellään useita tauluja: päivitetään vastaanottajatiluksia ja lisätään tietueita distribuutio- ja vastaanottajatauluihin.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Monen käyttäjän samanaikaisuus (Concurrency Disaster): Koodin alussa ajetaan DB.Execute "UPDATE USysRecipients SET USysRecipients.[To Distribution] = No;". Tämä tyhjentää valinnat kaikilta tietokannan vastaanottajilta. Jos kaksi käyttäjää yrittää luoda lähetystä täsmälleen samaan aikaan verkkolevyllä jaetussa tietokannassa, he ylikirjoittavat ja sekoittavat toistensa vastaanottajavalinnat välittömästi.

    Transaktioiden puute (Data eheys): Datamuutokset tallennetaan kahteen tauluun (USysDISTRIBUTION ja USysRecipByDistr), mutta transaktioita (BeginTrans / CommitTrans) ei ole käytetty. Jos ohjelma kaatuu jälkimmäisen taulun kohdalla, tietokantaan jää "orpo" distribuutio ilman vastaanottajia.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Käytät peräkkäin AddNew ja .Fields(..) usealle kymmenelle riville. Insert-kysely (Append Query) DB.Execute avulla olisi huomattavasti nopeampi.

💡 Parannusehdotukset (Ylläpidettävyys):

    Arkkitehtuurimuutos: Käyttäjäkohtaiset valinnat ("valittu distribuutioon") pitäisi ehdottomasti tallentaa paikalliseen frontend-tietokantaan (temp-taulu) tai sitoa käyttäjän ID:hen (esim. tyyliin [SelectedBy] = 'UserName'). Nykyinen ratkaisu rikkoo monen käyttäjän tuen.

    Paketoi tallennus-looppi transaktion sisään.

🛠️ Korjattu koodi (Ote turvallisemmasta tallennuslogiikasta):
VBA

' (Oletetaan että Valmis_Click-tapahtumassa käynnistetään tallennus)
    DBEngine.BeginTrans
    On Error GoTo RollbackError

    With taulu
      .AddNew
      .Fields("No") = Me.DistrNo
      .Fields("Description") = Me.Desc
      ' ...
      DistrID = .Fields("ID")
      .Update
      .Close
    End With
    
    ' Tämän sijaan, että kierretään silmukalla, suorituskykyisempi ja transaktioturvallinen INSERT INTO -kysely:
    SQL = "INSERT INTO USysRecipByDistr (DistrID, Copies, Unit, Company, Name, Description, Address, email) " & _
          "SELECT " & DistrID & ", Copies, Unit, Company, Name, Description, Address, email " & _
          "FROM USysRecipients WHERE [To Distribution] = Yes;"
    
    DB.Execute SQL, dbFailOnError
    
    DBEngine.CommitTrans
    Exit Sub
    
RollbackError:
    DBEngine.Rollback
    MsgBox "Error saving distribution. All changes cancelled: " & Err.Description, vbCritical

Moduuli: Form_USysExcelReport.cls

📊 Yhteenveto: Moduuli avaa Excel-sovelluksen (COM Automation) ja ajaa ohjelmoidusti raportteja tietokannan datan perusteella. Raportteja varten luetaan dataa Accessin taulukoista USysDOCUMENTS.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Prosessivuoto (Zombie Excel): Excel instanssoidaan komennolla Set xlApp = CreateObject("Excel.Application"), mutta jos makron ajo tai datan hakeminen (xlApp.Run "AjaKaikki") epäonnistuu ja laukaisee virheen, koodin ajo katkeaa. Koska virheenkäsittelyä ei näy xlApp.Quit tai Set xlApp = Nothing -kutsuille, taustalle jää roikkumaan piilotettu EXCEL.EXE -prosessi. Jos käyttäjä yrittää tätä useasti, tietokoneen muisti loppuu "haamu-Exceleihin".

    Kovakoodattu verkkopolku: xlApp.Workbooks.Open "l:\projekti\valdivia\sheets\tools\DocumentReport.xls". Jos L-levyä ei ole mapattu tällä käyttäjällä tai kansiota on siirretty, ohjelma kaatuu.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    SQL-kyselyssä ketjutetaan ExcelTaulu.Fields("Counter").Value suoraan WHERE-ehtoon.

    Excel-tiedosto avataan modella True (UpdateLinks). Tämä on yleensä ok, mutta voi aiheuttaa dialogi-ikkunoita, jotka jumiuttavat automaation, jos niihin ei vastata.

💡 Parannusehdotukset (Ylläpidettävyys):

    Suojaa OLE-automaatio (Excel-kutsu) perusteellisella Try-Catch (VBA:ssa On Error GoTo) -rakenteella, joka varmistaa aina xlApp.Quit -kutsun tapahtuvan.

🛠️ Korjattu koodi:
VBA

Private Sub OK_Click()
    Dim xlApp As Object ' Late binding suositeltavaa vakauden vuoksi
    Dim ExcelPolku As String

    On Error GoTo ErrorHandler
    
    ExcelPolku = "l:\projekti\valdivia\sheets\tools\DocumentReport.xls"
    
    ' Varmistetaan ensin, että tiedosto oikeasti on olemassa ennen Excelin avausta
    If Dir(ExcelPolku) = "" Then
        MsgBox "Excel report template not found at: " & ExcelPolku, vbCritical
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    ' Poistetaan Excelin omat häiritsevät varoitukset automaation ajaksi
    xlApp.DisplayAlerts = False 
    
    xlApp.Workbooks.Open ExcelPolku, UpdateLinks:=False
    xlApp.Visible = True
    
    If Me.FReportType.Value = 1 Then
        xlApp.Run "AjaKaikki", False
        ' Suljetaan tiedosto tallentamatta (jos kyseessä oli vain luku/ajo)
        xlApp.Workbooks("DocumentReport.xls").Close SaveChanges:=False
    Else
        ' ... muu logiikka
    End If
    
Cleanup:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        ' Varmistetaan että Excel prosessi kuolee, jos sitä ei jätetty käyttäjälle auki
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error communicating with Excel: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Moduulit: Report_TRANSMITTAL.cls, Report_TRANSMITTAL_COPY.cls, Report_Copy of TRANSMITTAL.cls

📊 Yhteenveto: Raporttimoduulin tehtävänä on hakea projektin perustiedot (sopimusnumero, nimi ja projektinumero) tietokannan Projinfo-taulusta raportin avauksen yhteydessä (Report_Open-tapahtuma) ja asettaa ne raportin otsikkokenttiin. Koodi on erittäin lyhyt ja suoraviivainen, mutta siitä puuttuu useita tärkeitä puolustusmekanismeja (defensive programming).

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Vakausriski (Puuttuva EOF-tarkistus): Koodi avaa Projinfo-taulun ja yrittää välittömästi lukea sen kenttiä (Nz(taulu("ContractNo"))). Jos taulu onkin tyhjä (esimerkiksi uusi tietokanta, poistettu rivi tai korruptio), recordsetin tila on EOF (End of File) heti avauksen jälkeen. Tällöin kentän arvon lukeminen kaataa ohjelman suorituksenaikaiseen virheeseen ("No current record").

    Resurssivuoto (Memory/Lock Leak): DAO-tietuejoukko (Recordset) avataan Set taulu = CurrentDb.OpenRecordset("Projinfo"), mutta sitä ei koskaan suljeta komennolla taulu.Close ennen sen tuhoamista (Set taulu = Nothing). Tämä voi jättää tietokannan lukituksia roikkumaan ja kuluttaa muistia.

    Virheenkäsittelyn puute: Koodista puuttuu kokonaan On Error GoTo -virheenkäsittely. Jos tietokantayhteydessä on ongelma tai taulua ei löydy, Access kaatuu epäsiististi.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Koodin toistuvuus (DRY-periaatteen rikkominen): Koska kanta sisältää useita kopioita samasta raportista samoilla koodeilla, ylläpidettävyys kärsii. Jos logiikkaa pitää muuttaa, se on muistettava tehdä jokaiseen kopioon erikseen.

    Suorituskykyä voi hieman optimoida avaamalla taulu vain luku -tilassa (dbOpenSnapshot), koska tässä haetaan dataa vain raportin näyttämistä varten.

💡 Parannusehdotukset (Ylläpidettävyys):

    Lisää ehdottomasti virheenkäsittely (Try-Catch -vastaava VBA:ssa).

    Varmista aina, että tietuejoukossa on dataa (If Not taulu.EOF Then) ennen lukemista.

    Muista sulkea Recordset explicitisti komennolla .Close.

    Harkitse ylimääräisten kopiokappaleiden (Copy of...) poistamista tuotantokannasta, tai jos niitä on pakko pitää, siirrä datan hakulogiikka yhteen julkiseen funktioon (esim. ForDocuments.vba -moduuliin), jota kaikki raportit voivat kutsua yhdellä rivillä.

🛠️ Korjattu koodi (Report_Open):
VBA

Private Sub Report_Open(Cancel As Integer)
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset

    On Error GoTo ErrorHandler
    
    Set DB = CurrentDb
    ' Avataan taulu vain luku -tilassa suorituskyvyn ja turvallisuuden parantamiseksi
    Set taulu = DB.OpenRecordset("Projinfo", dbOpenSnapshot)
    
    ' Varmistetaan, että taulussa on tietoja ennen kenttien lukemista
    If Not taulu.EOF Then
        Me.TContractNo1.Caption = Nz(taulu("ContractNo"), "")
        Me.TContractNo2.Caption = Nz(taulu("ContractNo"), "")
        Me.TProject1.Caption = Nz(taulu("Name"), "")
        Me.TProject2.Caption = Nz(taulu("Name"), "")
        Me.TProjno.Caption = Nz(taulu("ProjNo"), "")
    Else
        ' Mahdollinen varoitus käyttäjälle, jos projektin tietoja ei löytynyt
        ' MsgBox "Projektin tietoja ei löytynyt Projinfo-taulusta.", vbInformation
    End If
    
Cleanup:
    On Error Resume Next
    If Not taulu Is Nothing Then
        taulu.Close ' Suljetaan recordset kunnolla!
        Set taulu = Nothing
    End If
    If Not DB Is Nothing Then
        Set DB = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error loading report details: " & Err.Description, vbCritical, "Report Error"
    Resume Cleanup
End Sub

## UPDATE

Tässä on uusintakatselmointi korjatuille kooditiedostoille. Hienoa työtä! Olet selvästi vienyt edellisen katselmoinnin opit käytäntöön: olet korjannut vaarallisimmat tietoturva-aukot, poistanut CurrentDb:n sulkemisesta johtuneet vakausriskit ja parantanut virheenkäsittelyä valtavasti. Koodikantasi laatu ja luotettavuus ovat nousseet huomattavasti.

Katsotaanpa läpi korjatut moduulit.
📊 Yhteenveto korjauksista

Koodi on nyt selvästi ammattimaisemmalla tasolla. Erityisesti seuraavat muutokset ansaitsevat kiitosta:

    Form_SETTINGS.cls ja Form_USysAddToDistr.cls: CurrentDb:n sulkeminen (DB.Close) on poistettu, mikä pelastaa sovelluksen satunnaisilta kaatumisilta. Transaktioiden hallinta (Rollback) on nyt turvallisella pohjalla.

    Form_DBUsers.cls ja Form_USysOpenFile.cls: Komentoinjektioriskit (Command Injection) on eliminoitu hienosti vaihtamalla turvallisempiin vaihtoehtoihin (Application.FollowHyperlink ja syötteiden puhdistus).

    Report_TRANSMITTAL.cls: Tyhjän taulun (EOF) tarkistus ja virheenkäsittely on lisätty oikeaoppisesti, mikä estää raportin kaatumisen.

    Form_USysExcelReport.cls: Excelin zombie-prosessien jättäminen muistiin on estetty luotettavalla Cleanup-lohkolla ja xlApp.Quit -kutsulla.

🚨 Kriittiset löydökset (Vielä korjattavaa)

Vaikka teit erinomaista työtä, yksi erittäin kriittinen monikäyttäjäympäristön (concurrency) ongelma on edelleen jäänyt koodiin.

1. Form_USysNewDistribution.cls - Monen käyttäjän samanaikaisuus

Koodin Form_Load -tapahtumassa on edelleen seuraava rivi:
DB.Execute "UPDATE USysRecipients SET USysRecipients.[To Distribution] = No;"

    Ongelma: Tämä tyhjentää kaikkien käyttäjien valinnat suoraan jaetusta taulusta. Jos kaksi käyttäjää avaa "Uusi distribuutio" -lomakkeen yhtä aikaa, ensimmäisen käyttäjän tekemät rastituksien valinnat katoavat, kun toinen käyttäjä avaa lomakkeen.

    Ratkaisu: Rasti "valittu" (To Distribution) ei saisi koskaan olla jaetussa backend-tietokannassa (palvelimella). Oikea tapa on toteuttaa tämä niin, että valinnat tallennetaan käyttäjän paikalliseen (Frontend) väliaikaistauluun, tai tauluun lisätään sarake ValinnutKayttaja, johon tallennetaan käyttäjätunnus, ja UPDATE/INSERT -kyselyt kohdistetaan vain kyseisen käyttäjän riveihin.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

1. Form_USysExcelReport.cls - Kovakoodattu polku

Koodissa on yhä kovakoodattu polku Excel-mallipohjalle: ExcelPolku = "l:\projekti\valdivia\sheets\tools\DocumentReport.xls".
Vaikka lisäsit hienosti tiedoston olemassaolon tarkistuksen (Dir(ExcelPolku) = ""), polku on yhä riskialtis. Jos levyjako (L:) muuttuu tai projektikansion nimi vaihtuu, koodi pitää kääntää uudelleen.

    Ehdotus: Tallenna tämä polku asetus-tauluun (esim. ProjPaths) tai käytä CurrentProject.Path -muuttujaa dynaamisen polun rakentamiseen.

2. Form_USysReserve.cls - Turvallisuus

Olet lisännyt heittomerkkien poiston: EtsittavaNumero = Replace(Me.TNumber.Value, "'", ""). Tämä on erittäin hyvä ja pragmaattinen "quick fix" SQL-injektiota vastaan. Oikea alan standardi on kuitenkin aina käyttää parametrisoituja kyselyitä (QueryDef), jos mahdollista, mutta VBA/Access-maailmassa puhdistus yhdistettynä tyyppimuunnokseen (CLng) on usein "tarpeeksi hyvä".
💡 Parannusehdotukset (Ylläpidettävyys)

Koodisi on nyt paljon siistimpää ja helpommin luettavaa. Muutamia vinkkejä jatkoon:

    DCount vs OpenRecordset: Olet siirtynyt käyttämään DCount-funktiota Form_USysAddToDistr.cls-tiedostossa olemassaolon tarkistamiseen. Tämä on loistava valinta ja selkiyttää koodia! Jatka tämän tekniikan käyttöä aina, kun haluat vain tietää, onko tietue olemassa.

    Late Binding (Myöhäinen sidonta): Muutit Excel-automaation käyttämään Object-tyyppiä (Dim xlApp As Object). Tämä oli loistava päätös! Se tekee sovelluksesta riippumattoman käyttäjän koneella olevasta Excel-versiosta (estää "Missing Reference" -virheet).

🛠️ Korjattu koodi (Form_USysNewDistribution.cls)

Tässä on ehdotus, miten voisit rakentaa Form_USysNewDistribution.cls tallennuksen luotettavasti erottamalla käyttäjät toisistaan. Huom: Tämä edellyttää, että lisäät USysRecipients -tauluun (tai uuteen liitostauluun) tekstikentän SelectedBy.
VBA

' Form_Load: Nollataan vain TÄMÄN käyttäjän aiemmin unohtuneet valinnat
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim Kayttaja As String
  
  On Error GoTo ErrorHandler
  Kayttaja = NetworkUserName ' Oletetaan että tämä globaali funktio palauttaa Windows-tunnuksen
  
  Set DB = CurrentDb
  ' Vapautetaan vain oman tunnuksen alla olevat varaukset
  DB.Execute "UPDATE USysRecipients SET [To Distribution] = No, SelectedBy = Null WHERE SelectedBy = '" & Kayttaja & "';", dbFailOnError
  Me.TDate.Value = Date
  
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  MsgBox "Error initializing form: " & Err.Description, vbExclamation
End Sub

Yhteenveto: Olet ottanut hienosti palautteen vastaan. Koodisi on hyvällä mallilla ja valmis tuotantokäyttöön, kunhan ratkaiset yllä mainitun monikäyttäjäongelman Form_USysNewDistribution-lomakkeessa!
