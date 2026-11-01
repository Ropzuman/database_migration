Option Explicit

' Päivitetty 2026-11-01: refaktorointi (taulukkokirjoitus, turvallinen vienti, eksplisiittiset taulukkoviittaukset, ääkköset)
' Päivitetty 2025-10-30: toiminnallisuuteen vaikuttamaton siivous ja pienoisoptimoint
' Päivitetty 2025-10-26: 64-bittinen yhteensopivuus, suorituskykyoptimointi, parannettu virheenkäsittely
' Excel–AutoCAD-integraatio: Tuo/vie blokkiattribuutit ja teksti-entiteetit

' ============================================================================
' AutoCAD-vakiot – tarvitaan myöhäistä sidontaa varten
' ============================================================================
' Myöhäistä sidontaa käytettäessä (Object AcadApplication, AcadDocument jne. sijaan)
' AutoCAD-tyyppikirjastoa ei viitata, joten sisäänrakennetut vakiot eivät ole käytettävissä.
' Ne on määriteltävä manuaalisesti niiden numeroarvoilla.
' Lähde: Autodesk AutoCAD ActiveX/VBA Reference Documentation
' ============================================================================

' Valintametodit (KRIITTINEN: on pysyttävä Integer-tyyppinä)
' HUOM: acSelectionSetAll-arvon on oltava 5
Public Const acSelectionSetAll As Integer = 5       ' Valitaan kaikki entiteetit (oikea arvo)
Public Const acSelectionSetPrevious As Integer = 4  ' Valitaan edellinen valinta

' Piirustuksen tallennusversiot (KRIITTINEN: on pysyttävä Integer-tyyppinä)
Public Const acNative As Integer = 60               ' Nykyinen AutoCAD-versio
Public Const ac2004_dwg As Integer = 24             ' AutoCAD 2004 -muoto
Public Const ac2007_dwg As Integer = 36             ' AutoCAD 2007 -muoto
Public Const ac2010_dwg As Integer = 48             ' AutoCAD 2010 -muoto
Public Const ac2013_dwg As Integer = 60             ' AutoCAD 2013 -muoto

' Huom: ikkunatila-, aktiivitila- ja zoomausvakiot on määritelty käyttökohtissaan (esim. DATA.bas)
' 64-bittinen yhteensopivuus: Long on turvallisempi kuin Integer (yhtenäinen DATA.bas:n kanssa)
Private Const acModelSpace As Long = 1 ' varmistaa, että valinta kohdistuu malliavaruuteen

Public oACAD As Object ' AcadApplication (myöhäinen sidonta – yhteensopivuus)
Public oDOC As Object  ' AcadDocument – myöhäinen sidonta
Public OliAuki As Boolean
Public Ver As Long 64-bittistä yhteensopivuutta varten
Public Const DEBUG_TRACE As Boolean = False ' Aseta True vianmääritykseen (hiljentymätön Immediate Window)
Public Const KAYTA_TAULUKKOKIRJOITUSTA As Boolean = True ' False = tuonti kirjoittaa solu kerrallaan (varatila)
Public Const TYHJENNYS_MERKKI As String = "<TYHJÄ>" ' Excel-solun arvo, joka tyhjentää attribuutin tarkoituksella

'' Kevyt jäljitysapuri Immediate-ikkunaan (Ctrl+G) – aktiivinen vain DEBUG_TRACE = True -tilassa
Private Sub Trace(ByVal msg As String)
    If DEBUG_TRACE Then Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
End Sub

' Rakentaa DXF-entiteetti-tyyppisuodattimet (INSERT [+ TEXT/MTEXT tarvittaessa])
' Huom: Käytetään pilkuilla erotettua yhdistelmämerkkijonoa (ei <or>-ryhmitystä),
' koska <or> voi epäonnistua AutoCAD 2019 myöhäisessä sidonnassa.
' IncludeTexts=True lisää teksti-entiteetit INSERT:n lisäksi samaan tyyppijoukkoon.
Private Sub BuildTypeFilter(ByVal IncludeTexts As Boolean, ByRef FilterType() As Integer, ByRef FilterData() As Variant)
    ReDim FilterType(0 To 0)
    ReDim FilterData(0 To 0)
    FilterType(0) = 0
    If IncludeTexts Then
        FilterData(0) = "TEXT,MTEXT,DTEXT,INSERT"
    Else
        FilterData(0) = "INSERT"
    End If
End Sub

' Julkiset wrapper-makrot, jotta TuoDATA näkyy Makrot-valintaikkunassa ilman parametreja
Public Sub TuoDATA_All()
    ' Tuo kaikki entiteetit – ei edellistä AutoCAD-valintaa
    Trace "TuoDATA_All käynnistetty"
    TuoDATA False
End Sub

Public Sub TuoDATA_Selected()
    ' Tuo vain edellinen valinta AutoCADissä
    Trace "TuoDATA_Selected käynnistetty"
    TuoDATA True
End Sub

Public Sub TuoDATA(Optional Valitut As Boolean, Optional Filtterit As String)
' Filtterit: pilkuilla erotetut blokkien nimet; jos annettu, ne ohittavat Aloitus-taulukon D7-solun
' 7.3.2003 - VG
' 27.3.2003 - VG
' 19.1.2004 - VG
' 29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
' 26.10.2025 - 64-bit compatibility, array optimization for faster import

    Dim Tyhjenna As Boolean
    Dim Listasta As Boolean
    Dim VainValitut As Boolean
    Dim Joukko As Object ' AcadSelectionSet – myöhäinen sidonta
    Dim BlockArray As Variant
    Dim i As Long, j As Long, jj As Long, k As Long
    ' KRIITTINEN: FilterType-taulukon on pysyttävä Integer-tyyppinä (ei Long)
    ' AutoCAD:n SelectionSet.Select-API vaatii Integer-taulukon DXF-suodinkoodeille
    ' Muuttaminen Longiksi aiheuttaa virheen: "Invalid argument FilterType in Select"
    Dim FilterType() As Integer ' Dynaaminen taulukko (PAKKO olla Integer)
    Dim FilterData() As Variant  ' Dynaaminen taulukko suodinarvoja varten
    ' HUOM: RemoveItems odottaa Variant-taulukkoa olioista; Variant välttää tyyppiristiriidan
    Dim Poista() As Variant ' AcadEntity-oliotaulukko Variant-muodossa
    Dim L As Long
    Dim EiPoisteta As Boolean
    Dim Nimet As String
    Dim Blokit As Variant
    Dim Blokki As Object ' AcadBlockReference – myöhäinen sidonta
    Dim DWGName As String
    Dim Hakemisto As String
    Dim Rivi As Long
    Dim oText As Object  ' AcadText – myöhäinen sidonta
    Dim oMText As Object ' AcadMText – myöhäinen sidonta
    Dim DocRivi As Long
    Dim DocMaara As Long
    Dim Loytyi As Boolean
    Dim Docmode As Boolean
    Dim StepMsg As String ' Diagnostiikkaleipimurunäyte virheen sijainniksi
    Dim IncludeTexts As Boolean ' Prosessoidaanko teksti-entiteetit käyttöliittymäävalinnan mukaan
    Dim AllowAll As Boolean     ' Käytetäänkö jokerimerkkisuodatinta *
    Dim FoundAny As Boolean     ' Löydettiinkö yhtään ehtoja vastaavaa entiteettiä
    Dim oEnt As Object          ' Nykyinen entiteetti valintajoukosta (myöhäinen sidonta)
    Dim StartBaseRow As Long    ' Ensimmäinen tyhjä rivi ennen tuontia kokonaismäärän laskemiseksi
    Dim DocStartRow As Long     ' Ensimmäinen tulostusrivi nykyiselle piirustukselle
    Dim prevCalcMode As XlCalculation ' Käyttäjän laskentatila palautetaan sellaisenaan
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim TagCol As Object        ' Välimuisti: attribuutin tagi -> sarakeindeksi
    ' Massakirjoituspuskuri (rivit x sarakkeet)
    Dim buf() As Variant
    Dim rowCap As Long, colCap As Long, rowUsed As Long, maxColUsed As Long
    Dim selCount As Long
    ' Teksti-entiteettien erillinen puskuri (teksti ei kirjoita suoraan soluihin,
    ' jotta lohkopuskurin huuhtelu ei ylikirjoita teksti-rivejä yhteisestä DocStartRow-pisteestä)
    Dim textBuf() As Variant    ' 8 saraketta x rivit (TRANSPONOITU: rivimäärä on VIIMEINEN ulottuvuus,
                                 ' jotta "ReDim Preserve" voi laajentaa sitä turvallisesti - VBA sallii
                                 ' ReDim Preserven vain moniulotteisen taulukon VIIMEISELLE ulottuvuudelle.
    Dim textBufRows As Long     ' Käytettyjen rivien määrä teksti-puskurissa
    Dim textBufCap As Long      ' Teksti-puskurin nykyinen kapasiteetti
    Dim tIdx As Long            ' Iteraattori teksti-puskurin huuhteluun
    Dim textX As Double         ' Teksti-entiteetin X-koordinaatti
    Dim textY As Double         ' Teksti-entiteetin Y-koordinaatti
    Dim textStr As String       ' Teksti-entiteetin tekstisisältö
    ' In-loop-muuttujat (VBA nostaa nämä proseduuritasolle kääntämisen yhteydessä;
    ' määritellään tässä eksplisiittisesti sekaannusten ja väärien arvojen välttämiseksi)
    Dim entHandle As Variant    ' Nykyisen entiteetin kahva (handle)
    Dim entType As String       ' Entiteettityypin nimi (TypeName tai EntityName)
    Dim isBlock As Boolean      ' Onko entiteetti blokkiviite
    Dim isText As Boolean       ' Onko entiteetti TEXT-entiteetti
    Dim isMText As Boolean      ' Onko entiteetti MTEXT-entiteetti
    Dim tmp As Variant          ' Tilapäinen varavalinta-arvo
    Dim effName As String       ' Blokin EffectiveName (dynaamisten blokkien tukemiseksi)
    Dim ip As Variant           ' Blokin InsertionPoint-koordinaattitaulukko
    Dim ipT As Variant          ' Text-entiteetin InsertionPoint-koordinaattitaulukko
    Dim ipM As Variant          ' MText-entiteetin InsertionPoint-koordinaattitaulukko
    Dim TagName As String       ' Attribuutin taginimi (isot kirjaimet)
    Dim ColIdx As Long          ' Attribuutin sarakeindeksi
    Dim preEntName As String    ' Esisuodatuksen EntityName-arvo
    Dim preEffName As String    ' Esisuodatuksen EffectiveName-arvo
    Dim haveNameFilter As Boolean ' Onko nimipohjainen suodatin käytössä
    Dim outRng As Range         ' Kohde-alue taulukkokirjoitukselle
    Dim outArr() As Variant     ' Täsmälleen oikean kokoinen kirjoitustaulukko (blokit + tekstit)
    Dim totalRows As Long, r As Long, c As Long
    Dim DocmodeTalletettu As Boolean
    
    ' Excelin asetukset talletetaan heti alussa, jotta virhekäsittelijä palauttaa ne aina oikein
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalcMode = Application.Calculation

    On Error GoTo ErrHandler
  
    ' Nykyinen-valintaruutu (auki oleva AutoCAD-ikkuna) ohittaa Lista-valinnan.
    ' Tämä korjaa tilanteen, jossa Nykyinen = True mutta Lista = True samaan aikaan.
    If Aloitus.Nykyinen.Value Then
        Listasta = False
    Else
        Listasta = Aloitus.Lista.Value
    End If
    
    If Not Listasta Then
        If Valitut Then
            VainValitut = True
        ElseIf Not Aloitus.Nykyinen.Value Then
            ' Nykyinen-tilassa haetaan aina kaikki blokit avoimesta kuvasta (ei kysytä)
            If MsgBox("Poimitaanko vain valitut kohteet?", vbYesNo, "Tuo DATA") = vbYes Then
                VainValitut = True
            End If
        End If
        ' Jos Nykyinen=True, VainValitut pysää Falsena -> acSelectionSetAll käyttöön
    End If
  
    ' Minimoidaan Excel-käyttöliittymän päivitykset ja laskentakulut tuonnin aikana
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    StepMsg = "Haetaan AutoCAD-sovellusolio"
    Trace StepMsg
    ' Yhdistetään käynnissä olevaan AutoCAD-instanssiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
        ' Palautetaan Excel-asetukset ennen poistumista (vältetään manuaalilaskentatila)
        Application.EnableEvents = prevEvents
        Application.Calculation = prevCalcMode
        Application.ScreenUpdating = prevScreen
        Exit Sub
    End If
    On Error GoTo ErrHandler
 
    ' Tyhjennetäänkö työtäulukko ennen tuontia
    If Aloitus.Tyhjenna.Value = True Then
        Tyhjenna = True
    End If
    
    StepMsg = "Valitaan DATA-taulukko"
    Trace StepMsg
    DATA.Select
    
    If Tyhjenna Then
        DATA.Cells.Clear
        ' General-muoto koko taulukolle – lausekkeet ja numerot toimivat oikein
        DATA.Cells.NumberFormat = "General"
        ' Otsikkorivin muotoilu
        DATA.Rows("1:1").Font.Bold = True
        ' Määritetään otsikot
        DATA.Cells(1, 1).Value = "PATH"
        DATA.Cells(1, 2).Value = "DWG"
        DATA.Cells(1, 3).Value = "BLOCK"
        DATA.Cells(1, 4).Value = "HANDLE"
        DATA.Cells(1, 5).Value = "XCord"
        DATA.Cells(1, 6).Value = "YCord"
        DATA.Cells(1, 7).Value = "Layer"
        ' Sarakkeiden muotoilut tyypin mukaan
        DATA.Columns("A:A").NumberFormat = "@"   ' PATH tekstityyppinen
        DATA.Columns("B:B").NumberFormat = "@"   ' DWG tekstityyppinen
        DATA.Columns("C:C").NumberFormat = "@"   ' BLOCK tekstityyppinen
        DATA.Columns("D:D").NumberFormat = "@"   ' HANDLE tekstityyppinen
        DATA.Columns("E:F").NumberFormat = "General" ' koordinaatit numeerisina
        DATA.Columns("G:G").NumberFormat = "@"   ' Layer tekstityyppinen
        ' Attribuuttisarakkeet (H eteenpäin) jätetään General-muotoon
    End If
    ' Varmistetaan koordinaattisarakkeet numeerisina myös ilman tyhjennystä
    DATA.Columns("E:F").NumberFormat = "General"
    
    StepMsg = "Lasketaan asiakirjojen määrä"
    Trace StepMsg
    If Listasta Then
        TIEDLISTA.Select
        i = 1
        Do
            i = i + 1
            If TIEDLISTA.Cells(i, 1).Value = "" Then
                DocMaara = i - 1
                Exit Do
            End If
        Loop
        DATA.Select
    Else
        DocMaara = 1
    End If
    
    ' Haetaan ensimmäinen tyhjä rivi
    Rivi = 2
    Do While DATA.Cells(Rivi, 1).Value <> ""
        Rivi = Rivi + 1
    Loop
    
    StepMsg = "Jäsennellään blokkien nimet"
    Trace StepMsg
    If Len(Trim(Filtterit)) > 0 Then
        Nimet = Filtterit
    Else
        Nimet = CStr(Aloitus.Range("D7").Value) ' säilytetään alkuperäinen kirjainkoko DXF-suodatinta varten
    End If
    Blokit = Split(Nimet, ",")
    
    For i = 0 To UBound(Blokit)
        Blokit(i) = Trim(Blokit(i))
    Next i
    
    StepMsg = "Määritetään entiteettityypit"
    Trace StepMsg
    ' Rakennetaan DXF-tyyppisuodatin apufunktiolla
    IncludeTexts = (Aloitus.Range("D5").Value = "Tekstit" Or Aloitus.Range("D5").Value = "Blokit ja tekstit")
    BuildTypeFilter IncludeTexts, FilterType, FilterData
    
    ' Talletetaan ja muutetaan hetkeksi asiakirjatila (MDI)
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    DocmodeTalletettu = True
    oACAD.Preferences.System.SingleDocumentMode = False
    
    ' Tallennetaan ensimmäinen tyhjä rivi ennen tuontia rivimäärien laskemiseksi myöhemmin
    StartBaseRow = Rivi
    ' Alustetaan tagi-sarake-välimuisti
    Set TagCol = CreateObject("Scripting.Dictionary")

    ' Käsitellään jokainen piirustus
    For DocRivi = 1 To DocMaara
        Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
        
        If Not Listasta Then
            Set oDOC = oACAD.ActiveDocument
            Loytyi = True
        Else
            StepMsg = "Selvitetään nykyinen/kohdepiirustus"
            Trace StepMsg
            Loytyi = False
            For i = 0 To oACAD.Documents.Count - 1
                If UCase(oACAD.Documents(i).Name) = UCase(TiedostoNimi(CStr(TIEDLISTA.Cells(DocRivi, 1).Value))) Then
                    Set oDOC = oACAD.Documents(i)
                    Loytyi = True
                    Exit For
                End If
            Next i
            
            If Not Loytyi Then
                StepMsg = "Avataan piirustus listalta"
                Trace StepMsg
                Set oDOC = oACAD.Documents.Open(TIEDLISTA.Cells(DocRivi, 1).Value)
            End If
        End If

        ' Haetaan piirustuksen tiedot
        DWGName = Left(oDOC.Name, Len(oDOC.Name) - 4) ' Poistetaan .dwg-pääte
        Hakemisto = oDOC.Path

        StepMsg = "Siivotaan vanha valintajoukko"
        Trace StepMsg
        ' Varmistetaan valinta malliavaruudessa (vältetään pelkkään paperitilaan kohdistuvat poiminnat)
        On Error Resume Next
        oDOC.ActiveSpace = acModelSpace
        Err.Clear
        On Error GoTo ErrHandler
        ' Käydään läpi lopusta alkuun, koska poisto siirtää jäljellä olevien indeksejä
        For i = oDOC.SelectionSets.Count - 1 To 0 Step -1
            If oDOC.SelectionSets(i).Name = "EXCELHAKU" Then
                oDOC.SelectionSets(i).Delete
            End If
        Next i
        
        Set Joukko = oDOC.SelectionSets.Add("EXCELHAKU")

        StepMsg = "Valitaan entiteetit"
        Trace StepMsg
        ' Suodatuslogiikka – sama kuin alkuperäisessä koodissa
        AllowAll = False
        For i = 0 To UBound(Blokit)
            If Blokit(i) = "*" Then AllowAll = True: Exit For
        Next i
        haveNameFilter = False
        If Not AllowAll Then
            For i = 0 To UBound(Blokit)
                If Len(Blokit(i)) > 0 Then haveNameFilter = True: Exit For
            Next i
        End If
        Trace "Filter: [" & Join(Blokit, ",") & "] AllowAll=" & AllowAll & " haveNameFilter=" & haveNameFilter

        ' Valitaan entiteetit piirustuksesta (suodatetaan tyyppi-DXF-suodattimella)
        If VainValitut Then
            Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
        Else
            Joukko.Select acSelectionSetAll, , , FilterType, FilterData
        End If
        Trace "Valinnan koko ennen nimisuodatusta: " & Joukko.Count

        ' Esisuodatus: poistetaan blokit joiden nimi ei täsmää.
        ' Käytetään suoraa olio-ominaisuuksien hakua (ei CallByName/TypeName).
        ' Esialustetaan poistolista täyteen valintakokoon – vältetään O(n²) ReDim Preserve jokaisella kierroksella
        ' Siivotaan lopuksi todelliseen kokoon ReDim Preserve:llä
        If Joukko.Count > 0 Then ReDim Poista(0 To Joukko.Count - 1)
        L = 0
        For j = 0 To Joukko.Count - 1
            On Error Resume Next
            Set oEnt = Joukko.Item(j)
            preEntName = ""
            preEntName = oEnt.EntityName
            If Err.Number <> 0 Then preEntName = "": Err.Clear
            On Error GoTo ErrHandler
            If preEntName = "AcDbBlockReference" Then
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    On Error Resume Next
                    preEffName = ""
                    preEffName = oEnt.EffectiveName
                    If Err.Number <> 0 Then preEffName = "": Err.Clear
                    On Error GoTo ErrHandler
                    For k = 0 To UBound(Blokit)
                        If UCase(preEffName) = UCase(Blokit(k)) Then
                            EiPoisteta = True: Exit For
                        End If
                    Next k
                End If
                If Not EiPoisteta Then
                    Set Poista(L) = oEnt  ' Taulukko on jo esiallokoitu – ei ReDim Preserve -kopiointia
                    L = L + 1
                End If
            End If
        Next j
        If L > 0 Then
            ' Siivotaan Poista-taulukko todelliseen kokoon ennen RemoveItems-kutsua
            ReDim Preserve Poista(0 To L - 1)
            On Error Resume Next
            Joukko.RemoveItems Poista
            If Err.Number <> 0 Then
                ' Aiemmin virhe niellittiin hiljaa eikä epäonnistumisesta jäänyt jälkeä edes DEBUG_TRACE-tilassa.
                Trace "[VAROITUS] RemoveItems epäonnistui: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        Trace "Valinnan koko nimisuodatuksen jälkeen: " & Joukko.Count
    selCount = Joukko.Count
    ' Varataan massakirjoituspuskuri valintakoon mukaan (pieni lisäkettähaarukka)
    rowCap = selCount + 8
    colCap = 40 ' 7 peruskenttää + tyypilliset attribuutit
    ReDim buf(1 To rowCap, 1 To colCap)
    rowUsed = 0
    maxColUsed = 7
    ' Alustetaan teksti-puskuri asiakirjakohtaisesti (varmistaa nollauksen moniasiakir-
    ' jaturonnissa sen sijaan että aiemman asiakirjan teksti-entiteetit kasautuisivat seuraavaan)
    textBufRows = 0
    textBufCap = 16
    ReDim textBuf(1 To 8, 1 To textBufCap)
    FoundAny = False
    DocStartRow = Rivi
        For i = 0 To Joukko.Count - 1
            ' Haetaan entiteetti eksplisiittisesti Item-metodilla välttämään myöhäissidontaambiguiteetti
            StepMsg = "Haetaan entiteetti valinnasta: indeksi=" & i
            On Error Resume Next
            Set oEnt = Joukko.Item(i)
            If Err.Number <> 0 Or oEnt Is Nothing Then
                Err.Clear
                On Error GoTo ErrHandler
                GoTo ContinueEntities
            End If
            On Error GoTo ErrHandler

            Application.StatusBar = "Luetaan tietoa: " & i + 1 & "/" & Joukko.Count & "  File: " & DWGName
            
            ' Haetaan kahva; voi epäonnistua proxy-objekteilla
            StepMsg = "Luetaan entiteetin kahva"
            entHandle = Empty  ' Nollataan ennen hakua (Dim on proseduurin alussa)
            On Error Resume Next
            entHandle = oEnt.Handle
            Err.Clear
            On Error GoTo ErrHandler
            
            StepMsg = "Tarkistetaan entiteettityyppi"
            ' Nollataan per-iteraatio-muuttujat (Dim on proseduurin alussa)
            entType = "": isBlock = False: isText = False: isMText = False: tmp = Empty
            ' Kokeillaan ensin TypeName; jos se epäonnistuu, käytetään EntityName- tai ObjectName-varavalintaa
            On Error Resume Next
            entType = ""
            entType = TypeName(oEnt) ' esim. "IAcadBlockReference", "IAcadText", "IAcadMText"
            If Err.Number <> 0 Or entType = "" Then
                Err.Clear
                StepMsg = "Tarkistetaan entiteettityyppi: EntityName-varavalinta"
                tmp = CallByName(oEnt, "EntityName", VbGet)
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                    Err.Clear
                    StepMsg = "Tarkistetaan entiteettityyppi: ObjectName-varavalinta"
                    tmp = CallByName(oEnt, "ObjectName", VbGet)
                End If
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                    ' Entiteettityyppiä ei voitu määrittää; ohitetaan
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo ContinueEntities
                End If
                entType = CStr(tmp)
            End If
            On Error GoTo ErrHandler
            If DEBUG_TRACE Then Trace "Entiteettityyppi: " & entType

            ' Tunnistus sekä TypeName- että AcDb*-arvoista
            isBlock = (InStr(1, entType, "BlockReference", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbBlockReference", vbTextCompare) > 0)
            isMText = (InStr(1, entType, "MText", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbMText", vbTextCompare) > 0)
            isText = ((InStr(1, entType, "Text", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbText", vbTextCompare) > 0)) And Not isMText

            If isBlock Then
                ' Tarkistetaan blokin nimi suodatinta vasten - alkuperäisen koodin logiikka
                On Error Resume Next
                Set Blokki = oEnt
                If Err.Number <> 0 Or Blokki Is Nothing Then
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo ContinueEntities
                End If
                On Error GoTo ErrHandler
                ' Esisuodatus on jo karsittu nimen perusteella, joten EiPoisteta = True kaikille
                ' Joukossa jäljellä oleville blokeille. Tarkistus tehdään kuitenkin EffectiveNamella
                ' varmuuden vuoksi (dynaamisten blokkien ja myöhäissidonnan virhetilanteisiin).
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    effName = ""  ' Nollataan (Dim on proseduurin alussa)
                    On Error Resume Next
                    effName = Blokki.EffectiveName
                    If Err.Number <> 0 Then effName = "": Err.Clear
                    On Error GoTo ErrHandler
                    If DEBUG_TRACE Then Trace "  EffectiveName=[" & effName & "]"
                    For k = 0 To UBound(Blokit)
                        If UCase(effName) = UCase(Blokit(k)) Then
                            EiPoisteta = True: Exit For
                        End If
                    Next k
                    ' Jos EffectiveName hakeminen epäonnistui mutta esisuodatus käsitteli sen,
                    ' hyväksytään blokki (pre-filter on jo poistanut väärät blokit)
                    If Not EiPoisteta And haveNameFilter And effName = "" Then
                        EiPoisteta = True
                        If DEBUG_TRACE Then Trace "  [VAROITUS] EffectiveName ei luettavissa, hyväksytään esisuodatuksen perusteella"
                    End If
                    If Not EiPoisteta Then
                        If DEBUG_TRACE Then Trace "  [OHITETTU] EffectiveName='" & effName & "' ei vastaa suodatinta: " & Nimet
                    End If
                End If

                If EiPoisteta Then
                    ' Lisätään rivi puskuriin
                    rowUsed = rowUsed + 1
                    If rowUsed > rowCap Then
                        ' HUOM: buf on tarkoituksella (rivit, sarakkeet) -järjestyksessä, EI transponoitu.
                        ' Tämä haara on käytännössä saavuttamaton: rowCap = selCount + 8 ja rowUsed voi
                        ' korkeintaan olla selCount, joten rajaa ei koskaan ylitetä. Transponointi TÄSTÄ
                        ' haarasta olisi rikkonut alempana olevan, aidosti käytössä olevan colCap-kasvatuksen
                        ' (uusi attribuuttisarake), koska ReDim Preserve sallii vain VIIMEISEN ulottuvuuden
                        ' kasvattamisen – jompikumpi kasvusuunta on aina "väärä päin". Jos rowCap-kaavaa
                        ' joskus muutetaan niin että tämä haara voi laueta, se on korjattava kopioimalla
                        ' data käsin uuteen, isompaan taulukkoon (ei ReDim Preserve).
                        rowCap = rowCap + 64
                        ReDim Preserve buf(1 To rowCap, 1 To colCap)
                    End If
                    On Error Resume Next
                    buf(rowUsed, 1) = Hakemisto
                    buf(rowUsed, 2) = DWGName
                    buf(rowUsed, 3) = Blokki.EffectiveName
                    buf(rowUsed, 4) = entHandle
                    ' Myöhäinen sidonta: InsertionPoint palauttaa Variant-taulukon (x,y,z). Haetaan ja indeksoidaan.
                    ip = Empty  ' Nollataan (Dim on proseduurin alussa)
                    On Error Resume Next
                    ip = CallByName(Blokki, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ip) Then
                        buf(rowUsed, 5) = CDbl(ip(0)) '' XCord
                        buf(rowUsed, 6) = CDbl(ip(1)) '' YCord
                    Else
                        ' Varavalinta: yritetään ominaisuushakua suoraan
                        On Error Resume Next
                        buf(rowUsed, 5) = CDbl(Blokki.InsertionPoint(0))
                        buf(rowUsed, 6) = CDbl(Blokki.InsertionPoint(1))
                        Err.Clear
                        On Error GoTo ErrHandler
                    End If
                    buf(rowUsed, 7) = Blokki.Layer
                    On Error GoTo ErrHandler

                    StepMsg = "Luetaan blokin attribuutit"
                    If Blokki.HasAttributes Then
                        BlockArray = Blokki.GetAttributes
                        For jj = 0 To UBound(BlockArray)
                            ' tagName, colIdx on määritelty proseduurin alussa (Dim poistettu silmukan sisältä)
                            TagName = UCase(BlockArray(jj).TagString)
                            If Not TagCol.Exists(TagName) Then
                                ' Etsitään tai luodaan sarake tälle tagille
                                ColIdx = OtsS(TagName)
                                TagCol.Add TagName, ColIdx
                                ' Lisätään otsikkoon blokin nimi kommenttina (valinnainen, virhesuojattu)
                                On Error Resume Next
                                DATA.Cells(1, ColIdx).ClearNotes
                                DATA.Cells(1, ColIdx).AddComment Blokki.EffectiveName
                                On Error GoTo ErrHandler
                            Else
                                ColIdx = CLng(TagCol(TagName))
                            End If
                            ' Varmistetaan, että puskurissa on tarpeeksi sarakkeita
                            If ColIdx > colCap Then
                                colCap = ColIdx + 8
                                ReDim Preserve buf(1 To rowCap, 1 To colCap)
                            End If
                            If ColIdx > maxColUsed Then maxColUsed = ColIdx
                            buf(rowUsed, ColIdx) = BlockArray(jj).TextString
                        Next jj
                    End If
                    FoundAny = True
                End If
            ElseIf IncludeTexts And (isText Or isMText) Then
                ' Teksti-entiteetit puskuroidaan erilliseen textBuf-taulukkoon.
                ' Aiemmin teksti kirjoitettiin suoraan DATA.Cells(Rivi,...):lle, mutta lohkopuskurin
                ' huuhtelu (outRng.Value = buf) alkaa samasta DocStartRow-pisteestä ja ylikirjoittaisi
                ' ne. Erillinen puskuri + myöhäinen huuhtelu (blokkien jälkeen) ratkaisee ristiriidan.
                textX = 0: textY = 0: textStr = ""
                If isText Then
                    On Error Resume Next
                    Set oText = oEnt
                    If Err.Number <> 0 Or oText Is Nothing Then
                        Err.Clear
                        On Error GoTo ErrHandler
                        GoTo ContinueEntities
                    End If
                    On Error GoTo ErrHandler
                    textStr = oText.TextString
                    On Error Resume Next
                    ipT = CallByName(oText, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ipT) Then
                        textX = CDbl(ipT(0))
                        textY = CDbl(ipT(1))
                    Else
                        On Error Resume Next
                        textX = CDbl(oText.InsertionPoint(0))
                        textY = CDbl(oText.InsertionPoint(1))
                        Err.Clear
                        On Error GoTo ErrHandler
                    End If
                Else
                    On Error Resume Next
                    Set oMText = oEnt
                    If Err.Number <> 0 Or oMText Is Nothing Then
                        Err.Clear
                        On Error GoTo ErrHandler
                        GoTo ContinueEntities
                    End If
                    On Error GoTo ErrHandler
                    textStr = oMText.TextString
                    On Error Resume Next
                    ipM = CallByName(oMText, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ipM) Then
                        textX = CDbl(ipM(0))
                        textY = CDbl(ipM(1))
                    Else
                        On Error Resume Next
                        textX = CDbl(oMText.InsertionPoint(0))
                        textY = CDbl(oMText.InsertionPoint(1))
                        Err.Clear
                        On Error GoTo ErrHandler
                    End If
                End If
                ' Lisätään tekstidata erilliseen puskuriin – huuhdellaan blokkien jälkeen
                ' Myös PATH, DWG, tyyppi, HANDLE ja Layer tallennetaan
                ' jotta teksti-rivit ovat käytettäviä VieDATA:ssa (lukee sarakkeen 4 HANDLen)
                textBufRows = textBufRows + 1
                If textBufRows > textBufCap Then
                    textBufCap = textBufCap + 16
                    ' ReDim Preserve sallii vain viimeisen ulottuvuuden kasvattamisen, joten rivit ovat viimeisenä
                    ReDim Preserve textBuf(1 To 8, 1 To textBufCap)
                End If
                textBuf(1, textBufRows) = Hakemisto
                textBuf(2, textBufRows) = DWGName
                textBuf(3, textBufRows) = IIf(isText, "TEXT", "MTEXT")
                textBuf(4, textBufRows) = entHandle
                textBuf(5, textBufRows) = textX
                textBuf(6, textBufRows) = textY
                ' Layer-kentän haku – virheenkäsittely, koska myöhäinen sidonta ei takaa ominaisuuden saatavuutta
                On Error Resume Next
                If isText Then
                    textBuf(7, textBufRows) = oText.Layer
                Else
                    textBuf(7, textBufRows) = oMText.Layer
                End If
                If Err.Number <> 0 Then textBuf(7, textBufRows) = "": Err.Clear
                On Error GoTo ErrHandler
                textBuf(8, textBufRows) = textStr
                FoundAny = True
            Else
                ' Muut entiteettityypit ohitetaan
            End If
ContinueEntities:
        Next i

        ' Kirjoitetaan blokki- ja tekstirivit taulukkoon täsmälleen oikean kokoisena (ei tyhjiä ylimääräisiä
        ' puskurisoluja, jotka voisivat ylikirjoittaa olemassa olevaa dataa)
        totalRows = rowUsed + textBufRows
        If totalRows > 0 Then
            If textBufRows > 0 And maxColUsed < 8 Then maxColUsed = 8
            ReDim outArr(1 To totalRows, 1 To maxColUsed)
            For r = 1 To rowUsed
                For c = 1 To maxColUsed
                    outArr(r, c) = buf(r, c)
                Next c
            Next r
            For tIdx = 1 To textBufRows
                For c = 1 To 8
                    outArr(rowUsed + tIdx, c) = textBuf(c, tIdx)
                Next c
            Next tIdx
            Set outRng = DATA.Range(DATA.Cells(DocStartRow, 1), DATA.Cells(DocStartRow + totalRows - 1, maxColUsed))
            ' Kohdealueen on oltava tyhjä: tuonti ei saa koskaan ylikirjoittaa olemassa olevia tietoja
            If Application.WorksheetFunction.CountA(outRng) > 0 Then
                Err.Raise vbObjectError + 1001, "TuoDATA", "Kohdealue DATA-taulukossa (rivit " & DocStartRow & "-" & _
                    (DocStartRow + totalRows - 1) & ") ei ole tyhjä. Tuontia ei tehty, jotta olemassa olevat tiedot säilyvät."
            End If
            ' Muodot asetetaan ennen kirjoitusta: tekstimuoto säilyttää kahvat, etunollat ja =-alkuiset arvot sellaisinaan
            outRng.NumberFormat = "@"
            DATA.Range(DATA.Cells(DocStartRow, 5), DATA.Cells(DocStartRow + totalRows - 1, 6)).NumberFormat = "General"
            If KAYTA_TAULUKKOKIRJOITUSTA Then
                outRng.Value = outArr
            Else
                For r = 1 To totalRows
                    For c = 1 To maxColUsed
                        If Not IsEmpty(outArr(r, c)) Then DATA.Cells(DocStartRow + r - 1, c).Value = outArr(r, c)
                    Next c
                Next r
            End If
            ' Tarkistus: ensimmäisen ja viimeisen rivin kahva sekä sijainti luetaan takaisin ja verrataan puskuriin
            If CStr(DATA.Cells(DocStartRow, 4).Value) <> CStr(outArr(1, 4)) Or _
               CStr(DATA.Cells(DocStartRow + totalRows - 1, 4).Value) <> CStr(outArr(totalRows, 4)) Then
                Err.Raise vbObjectError + 1002, "TuoDATA", "Taulukkoon kirjoitettu data ei vastaa luettua dataa. Tuonti keskeytettiin."
            End If
            ' Tekstirivit väritetään yhdellä kutsulla koko alueelle
            If textBufRows > 0 Then
                DATA.Range(DATA.Cells(DocStartRow + rowUsed, 1), DATA.Cells(DocStartRow + totalRows - 1, 8)).Interior.ColorIndex = 8
            End If
            Rivi = DocStartRow + totalRows
        End If
        Trace "Piirustus käsitelty: " & DWGName & ", blokkeja: " & rowUsed & ", tekstejä: " & textBufRows
        
        ' Ilmoitetaan käyttäjälle, jos mitään ei löytynyt
        If Not FoundAny Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäsivät ehdon!", vbCritical, "Tuo DATA"
        End If
        
        ' Suljetaan listalta avattu piirustus (ei tallenneta) alkuperäisen käyttäytymisen mukaisesti
        If Not Loytyi Then
            oDOC.Close False
            Set oDOC = Nothing
        End If
    Next DocRivi
    Trace "TuoDATA valmis, yht. lisättyjä rivejä: " & (Rivi - StartBaseRow)
    
Cleanup:
    On Error Resume Next
    ' Virhetilanteessa listalta avattu piirustus suljetaan tallentamatta
    If Listasta And Not Loytyi And Not oDOC Is Nothing Then oDOC.Close False
    oACAD.Visible = True
    If DocmodeTalletettu Then oACAD.Preferences.System.SingleDocumentMode = Docmode
    ' Rajataan AutoFit käytettyyn alueeseen – koko taulukon autofit on hidas (pakottaa koko layout-laskennan)
    If Not DATA.UsedRange Is Nothing Then DATA.UsedRange.Columns.AutoFit
    Application.StatusBar = False
    ' Palautetaan Excel-asetukset (laskentatila palautetaan täsmälleen alkuperäiseen arvoonsa,
    ' myös jos se oli esim. xlCalculationSemiautomatic eikä pelkkä Auto/Manual)
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = prevScreen
    ' Vapautetaan objektit
    Set Blokki = Nothing
    Set Joukko = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
          "Vaihe: " & StepMsg, vbCritical, "Tuo DATA"
    Resume Cleanup
End Sub

Public Sub VieDATA()
' Viedään DATA-taulukon tiedot takaisin AutoCADiin
' 3.1.2001 - VG
' 4.6.2002 - VG
' 7.3.2003 - VG
' 27.3.2003 - VG
' 19.1.2004 - VG
' 29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
' 26.10.2025 - 64-bit compatibility, added error handling
' 27.2.2026 - TAG-pohjainen attribuuttien päivitys (korjaa blokkien tyhjentymisbugi)
' Tyhjä Excel-solu ei koskaan korvaa AutoCADin arvoa; attribuutin voi tyhjentää tarkoituksella
' merkinnällä TYHJENNYS_MERKKI. Excelin soluja ei muokata viennin aikana.

    Dim i As Long, j As Long
    Dim oEntity As Object
    Dim oBlock As Object ' AcadBlockReference – myöhäinen sidonta
    Dim oTxt As Object   ' AcadText tai AcadMText – myöhäinen sidonta
    Dim BlockArray As Variant
    Dim entName As String
    Dim Docmode As Boolean
    Dim DocmodeTalletettu As Boolean
    Dim VirheTapahtui As Boolean
    Dim StepMsg As String
    Dim TagName As String
    Dim ColIdx As Long
    Dim NewValue As String
    Dim Odotettu As String
    Dim OldValue As String
    Dim UpdateCount As Long
    Dim SkippedCount As Long
    Dim EmptyCount As Long
    Dim DuplicateCount As Long
    Dim MismatchCount As Long
    Dim NotFoundCount As Long
    Dim HeaderMap As Object ' Otsikkosarakkeiden välimuisti: TAG -> sarakeindeksi
    Dim TagSeen As Object   ' Yhden blokin jo käsitellyt tagit (saman nimiset tagit tunnistetaan)
    Dim dataArr As Variant  ' Koko DATA-taulukon sisältö muistissa (yksi COM-kutsu)
    Dim hk As Long, hv As String
    Dim lastRow As Long, lastCol As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalcMode As XlCalculation

    ' Excelin asetukset talletetaan ennen virheenkäsittelyn alkua, jotta ne voidaan aina palauttaa
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalcMode = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    StepMsg = "VieDATA: Alustus"
    Trace StepMsg
    Set oDOC = Nothing
    OliAuki = False

    Ver = acNative ' Ver = 60

    If Sheets("Start").Ver2004.Value = True Then
        Ver = ac2004_dwg ' Ver = 24
    ElseIf Sheets("Start").Ver2007.Value = True Then
        Ver = ac2007_dwg ' Ver = 36
    ElseIf Sheets("Start").Ver2010.Value = True Then
        Ver = ac2010_dwg ' Ver = 48
    ElseIf Sheets("Start").Ver2013.Value = True Then
        Ver = ac2013_dwg ' Ver = 60
    End If

    DATA.Select

    StepMsg = "Luetaan DATA-taulukko"
    Trace StepMsg
    lastRow = DATA.Cells(DATA.Rows.Count, 4).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "DATA-taulukossa ei ole vietäviä rivejä.", vbInformation, "Vie DATA"
        GoTo Cleanup
    End If
    lastCol = DATA.Cells(1, DATA.Columns.Count).End(xlToLeft).Column
    If lastCol < 8 Then lastCol = 8
    dataArr = DATA.Range(DATA.Cells(1, 1), DATA.Cells(lastRow, lastCol)).Value2

    ' Rakennetaan otsikkokartta kerran (TAG -> sarakeindeksi). Otsikot luetaan jo muistissa olevasta
    ' taulukosta, joten attribuuttisarakkeiden määrää ei rajoiteta (aiemmin vain sarakkeisiin H-IV).
    Set HeaderMap = CreateObject("Scripting.Dictionary")
    Set TagSeen = CreateObject("Scripting.Dictionary")
    For hk = 8 To lastCol
        If IsError(dataArr(1, hk)) Then Exit For
        hv = UCase(CStr(dataArr(1, hk)))
        If hv = "" Then Exit For ' Tyhjä otsikko – loppuu tähän
        If Not HeaderMap.Exists(hv) Then HeaderMap.Add hv, hk
    Next hk
    Trace "VieDATA: otsikkokartta rakennettu, " & HeaderMap.Count & " saraketta"

    StepMsg = "Yhdistetään AutoCADiin"
    Trace StepMsg

    ' Yhdistetään käynnissä olevaan AutoCAD-instanssiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")

    If Err.Number <> 0 Then
        On Error GoTo ErrHandler
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Vie DATA"
        GoTo Cleanup
    End If
    On Error GoTo ErrHandler

    Docmode = oACAD.Preferences.System.SingleDocumentMode
    DocmodeTalletettu = True
    oACAD.Preferences.System.SingleDocumentMode = False

    For i = 2 To lastRow
        If LueSolu(dataArr(i, 4), i, 4) = "" Then Exit For ' Excelin viimeinen rivi

        If AvaaDoc(i) Then
            Application.StatusBar = "Viedään tietoa blokkiin: " & i - 1
            StepMsg = "Haetaan entiteetti kahvalla: rivi=" & i
            Set oEntity = Nothing
            On Error Resume Next
            Set oEntity = oDOC.HandleToObject(LueSolu(dataArr(i, 4), i, 4))
            If Err.Number <> 0 Then Set oEntity = Nothing: Err.Clear
            On Error GoTo ErrHandler

            If oEntity Is Nothing Then
                ' Entiteetti on poistettu piirustuksesta; ohitetaan rivi keskeyttämättä vientiä
                NotFoundCount = NotFoundCount + 1
                Trace "Kahvaa ei löytynyt: rivi=" & i
            Else
                entName = oEntity.EntityName

                If entName = "AcDbBlockReference" Then
                    Set oBlock = oEntity
                    StepMsg = "Päivitetään blokin attribuutit: rivi=" & i
                    Trace StepMsg

                    If oBlock.HasAttributes Then
                        BlockArray = oBlock.GetAttributes
                        TagSeen.RemoveAll

                        ' TAG-POHJAINEN PÄIVITYSLOGIIKKA (symmetrinen TuoDATA:n kanssa)
                        For j = 0 To UBound(BlockArray)
                            TagName = ""
                            On Error Resume Next
                            TagName = UCase(BlockArray(j).TagString)
                            OldValue = BlockArray(j).TextString
                            If Err.Number <> 0 Then
                                If DEBUG_TRACE Then Trace "VIRHE: Attribuutin " & j & " luku epäonnistui: " & Err.Description
                                Err.Clear
                                On Error GoTo ErrHandler
                                GoTo NextAttribute
                            End If
                            On Error GoTo ErrHandler

                            ' Saman nimisistä tageista päivitetään vain ensimmäinen, koska Excelissä on vain yksi sarake
                            If TagSeen.Exists(TagName) Then
                                DuplicateCount = DuplicateCount + 1
                                GoTo NextAttribute
                            End If
                            TagSeen.Add TagName, True

                            ColIdx = 0
                            If HeaderMap.Exists(TagName) Then ColIdx = CLng(HeaderMap(TagName))

                            If ColIdx > 0 Then
                                NewValue = LueSolu(dataArr(i, ColIdx), i, ColIdx)

                                If Len(NewValue) > 0 Then
                                    If NewValue = TYHJENNYS_MERKKI Then Odotettu = "" Else Odotettu = NewValue
                                    BlockArray(j).TextString = Odotettu
                                    UpdateCount = UpdateCount + 1
                                    ' Varmistus: luetaan arvo takaisin, jotta hiljaiset epäonnistumiset huomataan
                                    If BlockArray(j).TextString <> Odotettu Then
                                        MismatchCount = MismatchCount + 1
                                        Trace "  [VAROITUS] " & TagName & ": arvo ei päivittynyt, rivi=" & i
                                    End If
                                    If DEBUG_TRACE Then Trace "  [" & TagName & "] '" & OldValue & "' -> '" & Odotettu & "'"
                                Else
                                    ' Excel-solu on tyhjä – säilytetään AutoCADin nykyinen arvo
                                    EmptyCount = EmptyCount + 1
                                    If DEBUG_TRACE Then Trace "  [" & TagName & "] OHITETTU (Excel tyhjä, säilytetään '" & OldValue & "')"
                                End If
                            Else
                                SkippedCount = SkippedCount + 1
                                If DEBUG_TRACE Then Trace "  [" & TagName & "] OHITETTU (ei Excel-saraketta)"
                            End If
NextAttribute:
                        Next j
                        ' Pakotetaan AutoCAD piirtämään blokki uudelleen, jotta muutokset näkyvät heti
                        On Error Resume Next
                        oBlock.Update
                        If Err.Number <> 0 Then
                            Trace "  [VAROITUS] oBlock.Update epäonnistui: " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo ErrHandler
                    Else
                        Trace "Blokilla ei ole attribuutteja"
                    End If

                ElseIf entName = "AcDbText" Or entName = "AcDbMText" Then
                    StepMsg = "Päivitetään tekstientiteetti: rivi=" & i
                    Trace StepMsg
                    Set oTxt = oEntity
                    NewValue = LueSolu(dataArr(i, 8), i, 8)
                    If Len(NewValue) > 0 Then
                        If NewValue = TYHJENNYS_MERKKI Then Odotettu = "" Else Odotettu = NewValue
                        OldValue = oTxt.TextString
                        oTxt.TextString = Odotettu
                        UpdateCount = UpdateCount + 1
                        If oTxt.TextString <> Odotettu Then MismatchCount = MismatchCount + 1
                        If DEBUG_TRACE Then Trace "  [" & entName & "] '" & OldValue & "' -> '" & Odotettu & "'"
                    Else
                        EmptyCount = EmptyCount + 1
                        If DEBUG_TRACE Then Trace "  [" & entName & "] OHITETTU (Excel tyhjä)"
                    End If
                Else
                    Trace "Entiteettityyppiä ei tueta: " & entName & ", rivi=" & i
                End If
            End If
        End If
    Next i

    ' Viimeinen piirustus tallennetaan aina (myös jo auki ollut); edelliset on tallennettu AvaaDoc-funktiossa
    StepMsg = "Tallennetaan piirustus"
    If Not oDOC Is Nothing Then
        If TallennaDoc(oDOC) Then
            If Not OliAuki Then oDOC.Close False  ' Suljetaan vain, jos makro avasi piirustuksen
        Else
            MsgBox "Piirustuksen tallennus epäonnistui: " & oDOC.Name & vbCrLf & "Kuva saattaa olla jollakin auki.", vbCritical, "Vie DATA"
        End If
    End If

    ' Viennin yhteenveto
    Trace "VieDATA valmis: Päivitetty=" & UpdateCount & ", Ohitettu(ei saraketta)=" & SkippedCount & ", Säilytetty(tyhjä)=" & EmptyCount

Cleanup:
    On Error Resume Next
    ' Regeneroidaan aktiivinen piirustus, jotta muutokset näkyvät heti
    If Not oDOC Is Nothing Then
        oDOC.Regen 1 ' acAllViewports
        If Err.Number <> 0 Then
            Trace "[VAROITUS] oDOC.Regen epäonnistui: " & Err.Description
            Err.Clear
        End If
    End If
    If DocmodeTalletettu Then oACAD.Preferences.System.SingleDocumentMode = Docmode
    Application.StatusBar = False
    Application.Calculation = prevCalcMode
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Aloitus.Activate

    ' Vapautetaan objektit
    Set oEntity = Nothing
    Set oBlock = Nothing
    Set oTxt = Nothing
    Erase BlockArray ' Variant-taulukko, ei Object
    Set HeaderMap = Nothing
    Set TagSeen = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing

    If Not VirheTapahtui And UpdateCount + EmptyCount + SkippedCount > 0 Then
        MsgBox "Vienti valmis." & vbCrLf & _
               "Päivitetty: " & UpdateCount & vbCrLf & _
               "Säilytetty (Excel tyhjä): " & EmptyCount & vbCrLf & _
               "Ohitettu (ei saraketta): " & SkippedCount & vbCrLf & _
               "Handlea ei löytynyt: " & NotFoundCount & vbCrLf & _
               "Saman nimisiä tageja ohitettu: " & DuplicateCount & vbCrLf & _
               "Arvo ei päivittynyt: " & MismatchCount, _
               IIf(MismatchCount > 0 Or NotFoundCount > 0, vbExclamation, vbInformation), "Vie DATA"
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    ' Virhetilanteessa piirustusta ei tallenneta eikä suljeta: keskeneräiset muutokset eivät päädy levylle
    VirheTapahtui = True
    Trace "ERROR in VieDATA: " & Err.Description & " @ " & StepMsg
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
          "Vaihe: " & StepMsg & vbCrLf & _
          "Virheen hetkellä käsitelty piirustus ei tallentunut; aiemmin käsitellyt piirustukset on jo tallennettu.", vbCritical, "Vie DATA"
    Resume Cleanup
End Sub

Public Sub PoistaBlokit()
' Poistaa DATA-taulukossa valittuja rivejä vastaavat entiteetit AutoCAD-piirustuksista
    ' 26.10.2025 – 64-bittinen yhteensopivuus, lisätty virheenkäsittely
    ' Otsikkorivi (rivi 1) ohitetaan, jos se on osa valintaa
    ' Piirustus tallennetaan vain, jos jotain poistettiin eikä virheitä tapahtunut

    Dim oEntity As Object
    Dim Alue As Range
    Dim AlueRajattu As Range
    Dim Rivi As Range
    Dim RiviNo As Long
    Dim Docmode As Boolean
    Dim DocmodeTalletettu As Boolean
    Dim VirheTapahtui As Boolean
    Dim PoistettuMaara As Long
    Dim EiLoytynytMaara As Long
    ' Dictionary tarjoaa eksaktin numerotäsmäyksen ja O(1)-haun (InStr("|1|") osui myös "|10|")
    Dim Kaydyt As Object
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalcMode As XlCalculation

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalcMode = Application.Calculation

    On Error GoTo ErrHandler
    Set oDOC = Nothing
    OliAuki = False

    ' Varmistetaan, että DATA-taulukko on valittuna ja valintana on solualue
    DATA.Select
    If TypeName(Selection) <> "Range" Then
        MsgBox "Valitse poistettavat rivit DATA-taulukosta.", vbInformation, "Poista Blokit"
        Exit Sub
    End If

    Set Kaydyt = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Yhdistetään käynnissä olevaan AutoCAD-instanssiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")

    If Err.Number <> 0 Then
        On Error GoTo ErrHandler
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Poista Blokit"
        GoTo Cleanup
    End If
    On Error GoTo ErrHandler

    Docmode = oACAD.Preferences.System.SingleDocumentMode
    DocmodeTalletettu = True
    oACAD.Preferences.System.SingleDocumentMode = False

    For Each Alue In Selection.Areas
        ' Rajataan käytettyyn alueeseen, jotta koko sarakkeen valinta ei käy läpi miljoonia rivejä
        Set AlueRajattu = Intersect(Alue, DATA.UsedRange)
        If Not AlueRajattu Is Nothing Then
            For Each Rivi In AlueRajattu.Rows
                RiviNo = Rivi.Row
                If RiviNo > 1 Then ' Otsikkorivi ohitetaan
                    If Not Kaydyt.Exists(RiviNo) Then
                        Kaydyt.Add RiviNo, True
                        If AvaaDoc(RiviNo) Then
                            Application.StatusBar = "Tuhotaan objektia rivillä: " & RiviNo
                            Set oEntity = Nothing
                            On Error Resume Next
                            Set oEntity = oDOC.HandleToObject(CStr(DATA.Cells(RiviNo, 4).Value))
                            If Err.Number <> 0 Then Set oEntity = Nothing: Err.Clear
                            On Error GoTo ErrHandler
                            If oEntity Is Nothing Then
                                EiLoytynytMaara = EiLoytynytMaara + 1
                            Else
                                oEntity.Delete
                                PoistettuMaara = PoistettuMaara + 1
                            End If
                        End If
                    End If
                End If
            Next Rivi
        End If
    Next Alue

    ' Tallennetaan vain, jos jotain poistettiin; muuten pelkkä auki ollut piirustus pysyy koskemattomana
    If Not oDOC Is Nothing Then
        If PoistettuMaara > 0 Then
            If TallennaDoc(oDOC) Then
                If Not OliAuki Then oDOC.Close False ' Suljetaan vain, jos makro avasi piirustuksen
            Else
                MsgBox "Piirustuksen tallennus epäonnistui: " & oDOC.Name, vbCritical, "Poista Blokit"
            End If
        ElseIf Not OliAuki Then
            oDOC.Close False
        End If
    End If

Cleanup:
    On Error Resume Next
    If DocmodeTalletettu Then oACAD.Preferences.System.SingleDocumentMode = Docmode
    Application.StatusBar = False
    Application.Calculation = prevCalcMode
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    Set oEntity = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing
    Set Kaydyt = Nothing

    If Not VirheTapahtui And PoistettuMaara + EiLoytynytMaara > 0 Then
        MsgBox "Tuhottu: " & PoistettuMaara & vbCrLf & "Kahvaa ei löytynyt: " & EiLoytynytMaara, vbInformation, "Poista Blokit"
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    ' Virhetilanteessa piirustusta ei tallenneta eikä suljeta
    VirheTapahtui = True
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Virheen hetkellä käsitelty piirustus ei tallentunut; aiemmin käsitellyt piirustukset on jo tallennettu.", vbCritical, "Poista Blokit"
    Resume Cleanup
End Sub

' Palauttaa solun arvon merkkijonona: virhearvo ja tyhjä solu palautuvat tyhjänä
Private Function LueSolu(ByVal v As Variant, ByVal Rivi As Long, ByVal Sarake As Long) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDouble Then
        ' General-muotoisesta solusta luetaan täysi tarkkuus (.Text katkoisi desimaalit sarakkeen leveyden mukaan);
        ' erikseen muotoillun solun näkyvä muoto säilytetään. Kapea sarake voi antaa "####", jolloin käytetään arvoa
        If DATA.Cells(Rivi, Sarake).NumberFormat = "General" Then
            LueSolu = CStr(v)
        Else
            LueSolu = DATA.Cells(Rivi, Sarake).Text
            If Left(LueSolu, 1) = "#" Then LueSolu = CStr(v)
        End If
    Else
        LueSolu = CStr(v)
    End If
End Function

' Palauttaa tiedoston nimen polusta (Dir ei toimi luotettavasti puuttuvilla tai verkkopoluilla)
Private Function TiedostoNimi(ByVal Polku As String) As String
    TiedostoNimi = Mid(Polku, InStrRev(Polku, "\") + 1)
End Function

' Tallentaa piirustuksen valittuun muotoon; palauttaa False, jos tallennus epäonnistuu
Private Function TallennaDoc(ByVal Doc As Object) As Boolean
    On Error GoTo Virhe
    If Ver = 0 Then Ver = acNative
    Doc.SaveAs Doc.FullName, Ver
    TallennaDoc = True
    Exit Function
Virhe:
    TallennaDoc = False
End Function

Private Function OtsS(ByVal Nimi As String) As Long
' Etsii tai luo sarakkeen attribuutin nimelle
    Dim i As Long
    Nimi = UCase(Nimi)
    i = 7
    Do While i <= DATA.Columns.Count
        If DATA.Cells(1, i).Value = "" Then
            DATA.Cells(1, i).Value = Nimi
            OtsS = i
            Exit Function
        ElseIf DATA.Cells(1, i).Value = Nimi Then
            OtsS = i
            Exit Function
        End If
        i = i + 1
    Loop
    Err.Raise vbObjectError + 1003, "OtsS", "Taulukossa ei ole tilaa uudelle attribuuttisarakkeelle: " & Nimi
End Function

Private Function AvaaDoc(Rivi As Long) As Boolean
' Avaa AutoCAD-piirustuksen tarvittaessa. Kun käsiteltävä piirustus vaihtuu, edellinen tallennetaan
' aina ja suljetaan vain, jos makro avasi sen (aiemmin jo auki ollut piirustus jäi tallentamatta).

    Dim Doku As String
    Dim Tiedosto As String
    Dim i As Long

    Doku = CStr(DATA.Cells(Rivi, 2).Value) & ".dwg"
    Tiedosto = CStr(DATA.Cells(Rivi, 1).Value) & "\" & Doku

    If Not oDOC Is Nothing Then
        ' Haluttu piirustus on jo aktiivinen
        If LCase(oDOC.Name) = LCase(Doku) Then
            AvaaDoc = True
            Exit Function
        End If

        ' Edellinen piirustus tallennetaan ja suljetaan ennen seuraavaan siirtymistä
        If TallennaDoc(oDOC) Then
            On Error Resume Next
            If Not OliAuki Then oDOC.Close False
            Err.Clear
            On Error GoTo 0
        Else
            MsgBox "Virhe talletettaessa piirustusta: " & oDOC.Name & vbCrLf & "Kuva saattaa olla jollakin auki.", vbCritical, "Vie tiedot"
        End If
        Set oDOC = Nothing
    End If

    ' Tarkistetaan, onko piirustus avoinna AutoCADissä
    OliAuki = False
    For i = 0 To oACAD.Documents.Count - 1
        If LCase(oACAD.Documents(i).Name) = LCase(Doku) Then
            OliAuki = True
            oACAD.Documents(i).Activate
            Set oDOC = oACAD.ActiveDocument
            AvaaDoc = True
            Exit Function
        End If
    Next i

    ' Piirustus ei ollut auki AutoCADissä – avataan se
    On Error Resume Next
    Set oDOC = oACAD.Documents.Open(Tiedosto)

    If Err.Number <> 0 Then
        MsgBox "Virhe avattaessa piirustusta: " & Doku, vbCritical, "Vie tiedot"
        AvaaDoc = False
        Set oDOC = Nothing
        Err.Clear
    Else
        AvaaDoc = True
    End If
    On Error GoTo 0
End Function

Sub Numerointi()
'' Numeerinen järjestely blokeille

    Dim Alku As String
    Dim Vali As Long
    Dim i As Long, j As Long
    Dim prevScreen As Boolean
    Dim prevCalcMode As XlCalculation

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    
    '' Haetaan tiedot piirustuksesta
    TuoDATA True
    Alku = Aloitus.Range("D13").Value
    Vali = 2
    
    DATA.Cells.Sort Key1:=DATA.Range("E2"), Order1:=xlAscending, Key2:=DATA.Range("F2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ' Suorituskykysuojaus: estetään uudelleenpiirrto ja -laskenta rivittäisen kirjoituksen ajaksi
    ' Alkuperäiset asetukset talletetaan ja palautetaan sellaisinaan
    prevScreen = Application.ScreenUpdating
    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo NumCleanup
    i = 2
    j = Val(Alku)
    Do
        If DATA.Cells(i, 1).Value = "" Then Exit Do
        DATA.Cells(i, 12).Value = LNumero(j, Alku)
        If Right(CStr(j), 1) = "8" Then
            j = j + Vali
        End If
        j = j + 1
        i = i + 1
    Loop
    
NumCleanup:
    ' Palautetaan Excel-asetukset aina, myös virhetilanteessa
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Virhe Numeroinnissa: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Numerointi"
        Exit Sub
    End If
    
    Aloitus.Range("D13").Value = LNumero(j, Alku)
    VieDATA
End Sub

Private Function LNumero(No As Long, Alku As String) As String
'' Muotoilee numeron johtavilla nollilla alkumerkkijonon pituuden mukaan

    LNumero = CStr(No)
    Do
        If Len(LNumero) < Len(Alku) Then
            LNumero = "0" & LNumero
        Else
            Exit Do
        End If
    Loop
End Function

Sub RefNumerointi()
'' Viitenumerointi-työkalu
'' Suorituskykysuojaus kirjoitussilmukan ympärillä

    Dim vSivu As Long
    Dim Kirjain As String
    Dim i As Long
    Dim prevScreen As Boolean
    Dim prevCalcMode As XlCalculation

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    Kirjain = "A"
    
    '' Haetaan data piirustuksesta
    TuoDATA True, "REFERENCE"
    vSivu = CLng(Aloitus.Range("D17").Value)
  
    DATA.Cells.Sort Key1:=DATA.Range("F2"), Order1:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ' Suorituskykysuojaus: estetään uudelleenpiirrto ja -laskenta rivittäisen kirjoituksen ajaksi
    ' (vastaa Numerointi-aliohjelman toteutusta)
    prevScreen = Application.ScreenUpdating
    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo RefNumCleanup
    
    i = 2
    Do
        If DATA.Cells(i, 1).Value = "" Then Exit Do
        DATA.Cells(i, 7).Value = "(" & vSivu & ":" & Kirjain & ") To Page " & vSivu
        DATA.Cells(i, 8).Value = "To Page " & vSivu + 1 & "(" & vSivu & ":" & Kirjain & ")"
        i = i + 1
        ' Kirjaimet A-Z riittävät 26 riville; sen jälkeen Chr(Asc+1) tuottaisi virheellisen merkin
        If Kirjain = "Z" And DATA.Cells(i, 1).Value <> "" Then
            Err.Raise vbObjectError + 1004, "RefNumerointi", "Viitteitä on yli 26; kirjaimet A-Z eivät riitä."
        End If
        If Kirjain <> "Z" Then Kirjain = Chr(Asc(Kirjain) + 1)
    Loop

RefNumCleanup:
    ' Palautetaan Excel-asetukset aina, myös virhetilanteessa
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Virhe RefNumeroinnissa: " & Err.Number & vbCrLf & Err.Description, vbCritical, "RefNumerointi"
        Exit Sub
    End If
    
    Aloitus.Range("D18").Value = vSivu + 1
    VieDATA
End Sub

Function Lisaa(Nro As String, Maara As Long) As String
'' Lisää arvon numerojonoon säilyttäen muotoilun

    Dim Pit As Long
    Dim i As Long
    
    Pit = Len(Nro)
    Nro = CStr(Val(Nro) + Maara)
    
    For i = 1 To Pit - Len(Nro)
        Nro = "0" & Nro
    Next i
    
    Lisaa = Nro
End Function

Function Yhd(Alue As Range, Optional Merkki As String) As String
'' Yhdistää alueen arvot erotinmerkillä
    Dim Solu As Range
    Dim Teksti As String
    
    For Each Solu In Alue
        If Teksti = "" Then
            Teksti = Solu.Value
        Else
            Teksti = Teksti & Merkki & Solu.Value
        End If
    Next
    
    Yhd = Teksti
End Function
