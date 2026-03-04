Option Explicit

' Päivitetty 2025-10-30: toiminnallisuuteen vaikuttamaton siivous ja pienoisoptimoint
' Päivitetty 2025-10-26: 64-bittinen yhteensopivuus, suorituskykyoptimoint, parannettu virheenkäsittely
' Excel–AutoCAD-integraatio: Tuo/vie blokkiattribuutit ja teksti-entiteetit
' Muutokset: Integer → Long (64-bitti), varhainen sidonta → myöhäinen sidonta (yhteensopivuus),
' lisätty virheenkäsittelijoita, taulukko-optimoint suorituskyvyn parantamiseksi

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

'' Huom: ikkunatila-, aktiivitila- ja zoomausvakiot on määritelty käyttökohdissaan (esim. DATA.bas)
' 64-bittinen yhteensopivuus: Long on turvallisempi kuin Integer (yhtenäinen DATA.bas:n kanssa)
Private Const acModelSpace As Long = 1 ' varmistaa, että valinta kohdistuu malliavaruuteen

Public oACAD As Object ' AcadApplication (myöhäinen sidonta – yhteensopivuus)
Public oDOC As Object  ' AcadDocument – myöhäinen sidonta
Public OliAuki As Boolean
Public Ver As Long ' Muutettu Integer → Long 64-bittistä yhteensopivuutta varten
Public Const DEBUG_TRACE As Boolean = False ' Aseta True vianmääritykseen (hiljentymätön Immediate Window)

' Kevyt jäljitysapuri Immediate-ikkunaan (Ctrl+G) – aktiivinen vain DEBUG_TRACE = True -tilassa
Private Sub Trace(ByVal msg As String)
    If DEBUG_TRACE Then Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
End Sub

' Rakentaa DXF-entiteetti-tyyppisuodattimet (INSERT [+ TEXT/MTEXT tarvittaessa])
' Huom: Käytetään pilkuilla erotettua yhdistelmämerkkijonoa (ei <or>-ryhmitystä),
' koska <or> voi epäonnistua AutoCAD 2019 myöhäisessä sidonnassa.
' IncludeTexts=True lisää teksti-entiteetit INSERT:n lisäksi samaan tyyppijoukkoon.
Private Sub BuildTypeFilter(ByVal includeTexts As Boolean, ByRef FilterType() As Integer, ByRef FilterData() As Variant)
    ReDim FilterType(0 To 0)
    ReDim FilterData(0 To 0)
    FilterType(0) = 0
    If includeTexts Then
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
' Import data from AutoCAD to Excel
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
    Dim i As Long, j As Long, jj As Long, k As Long ' Muutettu Integer → Long
    ' KRIITTINEN: FilterType-taulukon on pysyttävä Integer-tyyppinä (ei Long)
    ' AutoCAD:n SelectionSet.Select-API vaatii Integer-taulukon DXF-suodinkoodeille
    ' Muuttaminen Longiksi aiheuttaa virheen: "Invalid argument FilterType in Select"
    Dim FilterType() As Integer ' Dynaaminen taulukko (PAKKO olla Integer)
    Dim FilterData() As Variant  ' Dynaaminen taulukko suodinarvoja varten
    ' HUOM: RemoveItems odottaa Variant-taulukkoa olioista; Variant välttää tyyppiristiriidan
    Dim Poista() As Variant ' AcadEntity-oliotaulukko Variant-muodossa
    Dim L As Long ' Muutettu Integer → Long
    Dim EiPoisteta As Boolean
    Dim Nimet As String
    Dim Blokit As Variant
    Dim Blokki As Object ' AcadBlockReference – myöhäinen sidonta
    Dim DWGName As String
    Dim Hakemisto As String
    Dim Rivi As Long
    Dim oText As Object  ' AcadText – myöhäinen sidonta
    Dim oMText As Object ' AcadMText – myöhäinen sidonta
    Dim DocRivi As Long  ' Muutettu Integer → Long
    Dim DocMaara As Long ' Muutettu Integer → Long
    Dim Loytyi As Boolean
    Dim Docmode As Boolean
    Dim StepMsg As String ' Diagnostiikkaleipimurunäyte virheen sijainniksi
    Dim IncludeTexts As Boolean ' Prosessoidaanko teksti-entiteetit käyttöliittymävalinnan mukaan
    Dim AllowAll As Boolean     ' Käytetäänkö jokerimerkkisuodatinta *
    Dim FoundAny As Boolean     ' Löydettiinkö yhtään ehtoja vastaavaa entiteettiä
    Dim oEnt As Object          ' Nykyinen entiteetti valintajoukosta (myöhäinen sidonta)
    Dim StartBaseRow As Long    ' Ensimmäinen tyhjä rivi ennen tuontia kokonaismäärän laskemiseksi
    Dim DocStartRow As Long     ' Ensimmäinen tulostusrivi nykyiselle piirustukselle
    Dim wasCalcAuto As Boolean  ' Oliko Automaattinen laskenta käytössä ennen ajoa
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim TagCol As Object        ' Välimuisti: attribuutin tagi → sarakeindeksi
    ' Massakirjoituspuskuri (rivit x sarakkeet)
    Dim buf() As Variant
    Dim rowCap As Long, colCap As Long, rowUsed As Long, maxColUsed As Long
    Dim selCount As Long
    ' Teksti-entiteettien erillinen puskuri (BUG 1 -korjaus: teksti ei enää kirjoita suoraan soluihin,
    ' jotta lohkopuskurin huuhtelu ei ylikirjoita teksti-rivejä yhteisestä DocStartRow-pisteestä)
    Dim textBuf() As Variant    ' Teksti-rivien puskuri (rivit x 8 saraketta)
    Dim textBufRows As Long     ' Käytettyjen rivien määrä teksti-puskurissa
    Dim textBufCap As Long      ' Teksti-puskurin nykyinen kapasiteetti
    Dim tIdx As Long            ' Iteraattori teksti-puskurin huuhteluun
    Dim textStartRow As Long    ' Ensimmäinen tulostusrivi teksti-entiteeteille (blokkien jälkeen)
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
    Dim tagName As String       ' Attribuutin taginimi (isot kirjaimet)
    Dim colIdx As Long          ' Attribuutin sarakeindeksi
    Dim preEntName As String    ' Esisuodatuksen EntityName-arvo
    Dim preEffName As String    ' Esisuodatuksen EffectiveName-arvo
    Dim haveNameFilter As Boolean ' Onko nimipohjainen suodatin käytössä
    Dim outRng As Range         ' Kohde-alue lohkopuskurin huuhteluun
    
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
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    wasCalcAuto = (Application.Calculation = xlCalculationAutomatic)
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
        If wasCalcAuto Then
            Application.Calculation = xlCalculationAutomatic
        Else
            Application.Calculation = xlCalculationManual
        End If
        Application.ScreenUpdating = prevScreen
        Exit Sub
    End If
    On Error GoTo ErrHandler
 
    ' Tyhjennetäänkö työtaulukko ennen tuontia
    If Aloitus.Tyhjenna.Value = True Then
        Tyhjenna = True
    End If
    
    StepMsg = "Valitaan DATA-taulukko"
    Trace StepMsg
    DATA.Select
    
    If Tyhjenna Then
        Cells.Clear
        ' General-muoto koko taulukolle – lausekkeet ja numerot toimivat oikein
        Cells.NumberFormat = "General"
        ' Otsikkorivin muotoilu
        Rows("1:1").Font.Bold = True
        ' Määritetään otsikot
        Cells(1, 1).Value = "PATH"
        Cells(1, 2).Value = "DWG"
        Cells(1, 3).Value = "BLOCK"
        Cells(1, 4).Value = "HANDLE"
        Cells(1, 5).Value = "XCord"
        Cells(1, 6).Value = "YCord"
        Cells(1, 7).Value = "Layer"
        ' Sarakkeiden muotoilut tyypin mukaan
        Columns("A:A").NumberFormat = "@"   ' PATH tekstityyppinen
        Columns("B:B").NumberFormat = "@"   ' DWG tekstityyppinen
        Columns("C:C").NumberFormat = "@"   ' BLOCK tekstityyppinen
        Columns("D:D").NumberFormat = "@"   ' HANDLE tekstityyppinen
        Columns("E:F").NumberFormat = "General" ' koordinaatit numeerisina
        Columns("G:G").NumberFormat = "@"   ' Layer tekstityyppinen
        ' Attribuuttisarakkeet (H eteenpäin) jätetään General-muotoon
    End If
    ' Varmistetaan koordinaattisarakkeet numeerisina myös ilman tyhjennystiä
    Columns("E:F").NumberFormat = "General"
    
    StepMsg = "Lasketaan asiakirjojen määrä"
    Trace StepMsg
    If Listasta Then
        TIEDLISTA.Select
        i = 1
        Do
            i = i + 1
            If Cells(i, 1).Value = "" Then
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
    Do While Cells(Rivi, 1).Value <> ""
        Rivi = Rivi + 1
    Loop
    
    StepMsg = "Jäsennellään blokkien nimet"
    Trace StepMsg
    Nimet = CStr(Aloitus.Range("D7").Value) ' säilytetään alkuperäinen kirjainkoko DXF-suodatinta varten
    Blokit = Split(Nimet, ",")
    
    For i = 0 To UBound(Blokit)
        Blokit(i) = Trim(Blokit(i))
    Next i
    
    StepMsg = "Määritetään entiteettityypit"
    Trace StepMsg
    ' Rakennetaan DXF-suodatin käyttämällä apufunktiota (BuildTypeFilter) – eliminoi koodin toistoa.
    ' Alkuperäinen logiikka säilyy täsmälleen samana; kutsuminen on selkeämpää kuin inline-rakentelu.
    IncludeTexts = (Aloitus.Range("D5").Value = "Tekstit" Or Aloitus.Range("D5").Value = "Blokit ja tekstit")
    BuildTypeFilter IncludeTexts, FilterType, FilterData
    
    ' Save and temporarily change document mode
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    oACAD.Preferences.System.SingleDocumentMode = False
    
    ' Tallennetaan ensimmäinen tyhjä rivi ennen tuontia rivimäärien laskemiseksi myöhemmin
    StartBaseRow = Rivi
    ' Alustetaan tagi-sarake-välimuisti
    Set TagCol = CreateObject("Scripting.Dictionary")

    ' Process each document
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
                If UCase(oACAD.Documents(i).Name) = UCase(Dir(TIEDLISTA.Cells(DocRivi, 1).Value)) Then
                    Set oDOC = oACAD.Documents(i)
                    Loytyi = True
                    Exit For
                End If
            Next i
            
            If Not Loytyi Then
                StepMsg = "Open drawing from list"
                Trace StepMsg
                Set oDOC = oACAD.Documents.Open(TIEDLISTA.Cells(DocRivi, 1).Value)
            End If
        End If

        ' Get document information
        DWGName = Left(oDOC.Name, Len(oDOC.Name) - 4) ' Remove .dwg extension
        Hakemisto = oDOC.Path

        StepMsg = "Siivotaan vanha valintajoukko"
        Trace StepMsg
        ' Varmistetaan valinta malliavaruudessa (vältetään pelkkään paperitilaan kohdistuvat poiminnat)
        On Error Resume Next
        oDOC.ActiveSpace = acModelSpace
        Err.Clear
        On Error GoTo ErrHandler
        For i = 0 To oDOC.SelectionSets.Count - 1
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
        Dim haveNameFilter As Boolean: haveNameFilter = False
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

        ' Esisuodatus: poistetaan blokit joiden nimi ei täsmää – alkuperäinen logiikka.
        ' Käytetään suoraa olio-ominaisuuksien hakua (ei CallByName/TypeName), alkuperäinen tapa.
        ' Esialustetaan poistolista täyteen valintakokoon – vältetään O(n²) ReDim Preserve jokaisella kierroksella
        ' Siivotaan lopuksi todelliseen kokoon ReDim Preserve:llä
        If Joukko.Count > 0 Then ReDim Poista(0 To Joukko.Count - 1)
        L = 0
        For j = 0 To Joukko.Count - 1
            On Error Resume Next
            Set oEnt = Joukko.Item(j)
            Dim preEntName As String: preEntName = ""
            preEntName = oEnt.EntityName
            If Err.Number <> 0 Then preEntName = "": Err.Clear
            On Error GoTo ErrHandler
            If preEntName = "AcDbBlockReference" Then
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    On Error Resume Next
                    Dim preEffName As String: preEffName = ""
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
            On Error GoTo ErrHandler
        End If
        Trace "Valinnan koko nimisuodatuksen jälkeen: " & Joukko.Count
        If Joukko.Count = 0 Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
        End If
    selCount = Joukko.Count
    ' Varataan massakirjoituspuskuri valintakoon mukaan (pieni lisikettähaarukka)
    rowCap = selCount + 8
    colCap = 40 ' 7 peruskenttää + tyypilliset attribuutit
    ReDim buf(1 To rowCap, 1 To colCap)
    rowUsed = 0
    maxColUsed = 7
        ' Alustetaan erillinen teksti-puskuri (BUG 1 -korjaus)
        textBufRows = 0
        textBufCap = 16
        ReDim textBuf(1 To textBufCap, 1 To 8)
        ' Käsitellään valintajoukon entiteetit
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
            
            ' Prepare handle; may fail for proxies
            StepMsg = "Read entity handle"
            entHandle = Empty  ' Nollataan ennen hakua (Dim on proseduurin alussa)
            On Error Resume Next
            entHandle = oEnt.Handle
            Err.Clear
            On Error GoTo ErrHandler
            
            StepMsg = "Check entity type"
            ' Nollataan per-iteraatio-muuttujat (Dim on proseduurin alussa)
            entType = "": isBlock = False: isText = False: isMText = False: tmp = Empty
            ' Try TypeName first; if it fails (rare), fall back to EntityName or ObjectName via CallByName
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
            Trace "Entiteettityyppi: " & entType

            ' Normalize checks for both interface TypeName and AcDb* values
            isBlock = (InStr(1, entType, "BlockReference", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbBlockReference", vbTextCompare) > 0)
            isMText = (InStr(1, entType, "MText", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbMText", vbTextCompare) > 0)
            isText = ((InStr(1, entType, "Text", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbText", vbTextCompare) > 0)) And Not isMText

            If isBlock Then
                ' Tarkistetaan blokin nimi suodatinta vasten - alkuperaisen koodin logiikka
                On Error Resume Next
                Set Blokki = oEnt
                If Err.Number <> 0 Or Blokki Is Nothing Then
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo ContinueEntities
                End If
                On Error GoTo ErrHandler
                ' Esisuodatus on jo karsittu nimen perusteella, joten EiPoisteta = True kaikille
                ' Joukossa jaljella oleville blokeille. Tarkistus tehdaan kuitenkin EffectiveNamella
                ' varmuuden vuoksi (dynaamisten blokkien ja myohaissidonnan virhetilanteisiin).
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    effName = ""  ' Nollataan (Dim on proseduurin alussa)
                    On Error Resume Next
                    effName = Blokki.EffectiveName
                    If Err.Number <> 0 Then effName = "": Err.Clear
                    On Error GoTo ErrHandler
                    Trace "  EffectiveName=[" & effName & "]"
                    For k = 0 To UBound(Blokit)
                        If UCase(effName) = UCase(Blokit(k)) Then
                            EiPoisteta = True: Exit For
                        End If
                    Next k
                    ' Jos EffectiveName hakeminen epaonnistui mutta esisuodatus kasitteli sen,
                    ' hyvaksytaan blokki (pre-filter on jo poistanut vaarat blokit)
                    If Not EiPoisteta And haveNameFilter And effName = "" Then
                        EiPoisteta = True
                        Trace "  [VAROITUS] EffectiveName ei luettavissa, hyvaksytaan esisuodatuksen perusteella"
                    End If
                    If Not EiPoisteta Then
                        Trace "  [OHITETTU] EffectiveName='" & effName & "' ei vastaa suodatinta: " & Nimet
                    End If
                End If

                If EiPoisteta Then
                    ' Lisätään rivi puskuriin
                    rowUsed = rowUsed + 1
                    If rowUsed > rowCap Then
                        ' Laajennetaan puskuririvejä tarvittaessa
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
                            tagName = UCase(BlockArray(jj).TagString)
                            If Not TagCol.Exists(tagName) Then
                                ' Find or create a column for this tag
                                colIdx = OtsS(tagName)
                                TagCol.Add tagName, colIdx
                                ' Annotate header with block name (optional, error-safe)
                                On Error Resume Next
                                Cells(1, colIdx).ClearNotes
                                Cells(1, colIdx).AddComment Blokki.EffectiveName
                                On Error GoTo ErrHandler
                            Else
                                colIdx = CLng(TagCol(tagName))
                            End If
                            ' Ensure buffer has enough columns
                            If colIdx > colCap Then
                                colCap = colIdx + 8
                                ReDim Preserve buf(1 To rowCap, 1 To colCap)
                            End If
                            If colIdx > maxColUsed Then maxColUsed = colIdx
                            buf(rowUsed, colIdx) = BlockArray(jj).TextString
                        Next jj
                    End If
                    FoundAny = True
                End If
            ElseIf IncludeTexts And (isText Or isMText) Then
                ' BUG 1 -KORJAUS: Teksti-entiteetit puskuroidaan erilliseen textBuf-taulukkoon.
                ' Aiemmin teksti kirjoitettiin suoraan Cells(Rivi,...):lle, mutta lohkopuskurin
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
                ' Lisätään tekstidataa erilliseen puskuriin – huuhdellaan blokkien jälkeen
                textBufRows = textBufRows + 1
                If textBufRows > textBufCap Then
                    textBufCap = textBufCap + 16
                    ReDim Preserve textBuf(1 To textBufCap, 1 To 8)
                End If
                textBuf(textBufRows, 5) = textX
                textBuf(textBufRows, 6) = textY
                textBuf(textBufRows, 8) = textStr
                FoundAny = True
            Else
                ' Skip other entity types
            End If
ContinueEntities:
        Next i

        ' Huuhdellaan lohkopuskuroidut rivit taulukkoon yhdellä kirjoituksella
        If rowUsed > 0 Then
            Set outRng = Range(Cells(DocStartRow, 1), Cells(DocStartRow + rowUsed - 1, maxColUsed))
            outRng.Value = buf
            Rivi = DocStartRow + rowUsed
        End If
        ' Huuhtelun jälkeen pakotetaan koordinaattisarakkeet numeerisiksi (korjaa mahdolliset tekstijääntämät)
        If rowUsed > 0 Then
            With Range(Cells(DocStartRow, 5), Cells(DocStartRow + rowUsed - 1, 6))
                .NumberFormat = "General"
                .Value = .Value
            End With
        End If
        ' BUG 1 -KORJAUS: Huuhdellaan teksti-entiteetit blokkien jälkeen omille riveilleen.
        ' Värikoodi (ColorIndex=8) lisätään solukohtaisesti, koska Array→Range-dump ei tue solun muotoilua.
        If textBufRows > 0 Then
            textStartRow = Rivi
            For tIdx = 1 To textBufRows
                Cells(textStartRow + tIdx - 1, 5).Value = textBuf(tIdx, 5)
                Cells(textStartRow + tIdx - 1, 6).Value = textBuf(tIdx, 6)
                Cells(textStartRow + tIdx - 1, 8).Value = textBuf(tIdx, 8)
                Range(Cells(textStartRow + tIdx - 1, 1), Cells(textStartRow + tIdx - 1, 8)).Interior.ColorIndex = 8
            Next tIdx
            Rivi = textStartRow + textBufRows
        End If
        Trace "Piirustus käsitelty: " & DWGName & ", blokkeja: " & rowUsed & ", tekstejä: " & textBufRows
        
        ' Ilmoitetaan käyttäjälle, jos mitään ei löytynyt
        If Not FoundAny Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
        End If
        
        ' Suljetaan listalta avattu piirustus (ei tallenneta) alkuperäisen käyttäytymisen mukaisesti
        If Not Loytyi Then
            oDOC.Close False
        End If
    Next DocRivi
    Trace "TuoDATA valmis, yht. lisättyjä rivejä: " & (Rivi - StartBaseRow)
    
Cleanup:
    On Error Resume Next
    oACAD.Visible = True
    oACAD.Preferences.System.SingleDocumentMode = Docmode
    ' Rajataan AutoFit käytettyyn alueeseen – koko taulukon autofit on hidas (pakottaa koko layout-laskennan)
    If Not DATA.UsedRange Is Nothing Then DATA.UsedRange.Columns.AutoFit
    Application.StatusBar = False
    ' Palautetaan Excel-asetukset
    Application.EnableEvents = prevEvents
    If wasCalcAuto Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.ScreenUpdating = prevScreen
    ' Pakotetaan laskenta äänekään tulkintalipukkeiden tyhjentämiseksi muuttamatta käyttäjän asetusta
    If Application.Calculation = xlCalculationManual Then
        Application.CalculateFullRebuild
    Else
        Application.CalculateFull
    End If
    
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
' Export data from Excel back to AutoCAD
' 3.1.2001 - VG
' 4.6.2002 - VG
' 7.3.2003 - VG
' 27.3.2003 - VG
' 19.1.2004 - VG
' 29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
' 26.10.2025 - 64-bit compatibility, added error handling
' 27.2.2026 - CRITICAL FIX: TAG-pohjainen attribuuttien päivitys (korjaa blokkien tyhjentymisbugi)

    Dim i As Long, j As Long, k As Long ' Muutettu Integer → Long
    Dim oEntity As Object
    Dim oBlock As Object ' AcadBlockReference – myöhäinen sidonta
    Dim BlockArray As Variant
    Dim BlockNimi As String
    Dim DWGName As String
    Dim oText As Object  ' AcadText – myöhäinen sidonta
    Dim oMText As Object ' AcadMText – myöhäinen sidonta
    Dim Docmode As Boolean
    Dim StepMsg As String
    Dim TagName As String
    Dim ColIdx As Long
    Dim NewValue As String
    Dim OldValue As String
    Dim UpdateCount As Long
    Dim SkippedCount As Long
    Dim EmptyCount As Long
    Dim HeaderMap As Object ' Otsikkosarakkeiden välimuisti: TAG -> sarakeindeksi
    
    StepMsg = "VieDATA: Alustus"
    Trace StepMsg
    UpdateCount = 0
    SkippedCount = 0
    EmptyCount = 0
    
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
    
    ' Varmista että DATA-taulukko on aktiivinen
    DATA.Select

    ' Rakennetaan otsikkokartta kerran (TAG -> sarakeindeksi) suorituskyvyn parantamiseksi.
    ' Näin jokainen attribuutti ei vaadi erillistä k=8..256-silmukkaa per rivi.
    Set HeaderMap = CreateObject("Scripting.Dictionary")
    Dim hk As Long
    For hk = 8 To 256
        Dim hv As String
        hv = UCase(Cells(1, hk).Value)
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
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Vie DATA"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    oACAD.Preferences.System.SingleDocumentMode = False
    
    i = 1
    Do
        i = i + 1
        If Cells(i, 4).Value = "" Then ' Last row in Excel
            ' Tallennetaan aina riippumatta OliAuki-lipusta – muutoin jo auki oleva piirustus jää tallentamatta
            If Not oDOC Is Nothing Then
                oDOC.SaveAs oDOC.FullName, Ver
                If Not OliAuki Then oDOC.Close False  ' Suljetaan vain, jos macro avasi piirustuksen
            End If
            Exit Do
        Else
            If AvaaDoc(i) Then
                Application.StatusBar = "Viedään tietoa blokkiin: " & i - 1
                Set oEntity = oDOC.HandleToObject(Cells(i, 4).Text)
                
                If oEntity.EntityName = "AcDbBlockReference" Then ' Block
                    Set oBlock = oEntity
                    StepMsg = "Update block attributes: row=" & i
                    Trace StepMsg
                    
                    If oBlock.HasAttributes Then
                        BlockArray = oBlock.GetAttributes
                        Trace "Blokilla " & (UBound(BlockArray) + 1) & " attribuuttia"
                        
                        ' TAG-POHJAINEN PÄIVITYSLOGIIKKA (symmetrinen TuoDATA:n kanssa)
                        ' Käydään läpi jokainen blokin attribuutti
                        For j = 0 To UBound(BlockArray)
                            On Error Resume Next
                            TagName = UCase(BlockArray(j).TagString)
                            OldValue = BlockArray(j).TextString
                            If Err.Number <> 0 Then
                                Trace "VIRHE: Attribuutin " & j & " luku epäonnistui: " & Err.Description
                                Err.Clear
                                GoTo NextAttribute
                            End If
                            On Error GoTo ErrHandler
                            
                            ' Haetaan sarakeindeksi välimuistista (nopeampi kuin k=8..256-skannaus)
                            ColIdx = 0
                            If HeaderMap.Exists(TagName) Then ColIdx = CLng(HeaderMap(TagName))
                            
                            If ColIdx > 0 Then
                                ' Sarake löytyi – tarkistetaan, onko Excel-arvo ei-tyhjä
                                NewValue = CStr(Cells(i, ColIdx).Text)
                                
                                If Len(NewValue) > 0 Then
                                    ' Päivitetään attribuutti vain, jos Excelissä on arvo
                                    BlockArray(j).TextString = NewValue
                                    UpdateCount = UpdateCount + 1
                                    Trace "  [" & TagName & "] '" & OldValue & "' -> '" & NewValue & "'"
                                Else
                                    ' Excel-solu on tyhjä – säilytetään AutoCAD:n nykyinen arvo
                                    EmptyCount = EmptyCount + 1
                                    Trace "  [" & TagName & "] OHITETTU (Excel tyhjä, säilytetään '" & OldValue & "')"
                                End If
                            Else
                                ' Vastaavaa Excel-saraketta ei löytynyt
                                SkippedCount = SkippedCount + 1
                                Trace "  [" & TagName & "] OHITETTU (ei Excel-saraketta)"
                            End If
NextAttribute:
                        Next j
                        ' Pakota AutoCAD piirtämään blokin uudelleen - ilman tätä
                        ' attribuuttimuutokset eivat näy näytolla ennen tallennusta.
                        On Error Resume Next
                        oBlock.Update
                        If Err.Number <> 0 Then
                            Trace "  [VAROITUS] oBlock.Update epaonnistui: " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo ErrHandler
                    Else
                        Trace "Block has no attributes"
                    End If
                Else
                    ' Text or MText entity
                    StepMsg = "Update text entity: row=" & i
                    Trace StepMsg
                    
                    If oEntity.EntityName = "AcDbText" Then
                        Set oText = oEntity
                        NewValue = CStr(Cells(i, 8).Value)
                        If Len(NewValue) > 0 Then
                            OldValue = oText.TextString
                            oText.TextString = NewValue
                            UpdateCount = UpdateCount + 1
                            Trace "  [TEXT] '" & OldValue & "' -> '" & NewValue & "'"
                        Else
                            EmptyCount = EmptyCount + 1
                            Trace "  [TEXT] SKIPPED (Excel empty)"
                        End If
                    Else
                        Set oMText = oEntity
                        NewValue = CStr(Cells(i, 8).Value)
                        If Len(NewValue) > 0 Then
                            OldValue = oMText.TextString
                            oMText.TextString = NewValue
                            UpdateCount = UpdateCount + 1
                            Trace "  [MTEXT] '" & OldValue & "' -> '" & NewValue & "'"
                        Else
                            EmptyCount = EmptyCount + 1
                            Trace "  [MTEXT] SKIPPED (Excel empty)"
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    ' Viennin yhteenveto
    Trace "VieDATA valmis: Päivitetty=" & UpdateCount & ", Ohitettu(ei saraketta)=" & SkippedCount & ", Säilytetty(tyhjä)=" & EmptyCount
  
Cleanup:
    On Error Resume Next
    ' Regeneroi aktiivinen piirustus ettei muutokset jääkään näytolla nahtaviksi
    ' vasta tallennuksen jalkeen. acAllViewports = 1.
    If Not oDOC Is Nothing Then
        oDOC.Regen 1
        If Err.Number <> 0 Then
            Trace "[VAROITUS] oDOC.Regen epaonnistui: " & Err.Description
            Err.Clear
        End If
    End If
    Aloitus.Activate
    If Not oACAD Is Nothing Then
        oACAD.Preferences.System.SingleDocumentMode = Docmode
    End If
    Application.StatusBar = False

    ' Vapauta objektit
    Set oEntity = Nothing
    Set oBlock = Nothing
    Erase BlockArray  ' BlockArray on Variant-taulukko, ei Object – Set Nothing aiheuttaisi runtime-virheen
    Set HeaderMap = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    Trace "ERROR in VieDATA: " & Err.Description & " @ " & StepMsg
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
          "Vaihe: " & StepMsg, vbCritical, "Vie DATA"
    Resume Cleanup
End Sub

Public Sub PoistaBlokit()
' Poistaa valitut blokit AutoCAD-piirustuksesta
' 26.10.2025 – 64-bittinen yhteensopivuus, lisätty virheenkäsittely

    Dim i As Long, j As Long ' Muutettu Integer → Long
    Dim Docmode As Boolean
    Dim oEntity As Object
    Dim DWGName As String
    Dim Rivi As Range
    Dim RiviNo As Long
    ' KORJATTU: Kaydyt-merkkijono → Dictionary – InStr("|1|") osui myös "|10|", "|21|" jne.
    ' Dictionary tarjoaa O(1)-haun ja eksaktin numerotäsmäyksen
    Dim Kaydyt As Object

    On Error GoTo ErrHandler
    
    ' Alustetaan käytyjen rivien seurantasanakirja
    Set Kaydyt = CreateObject("Scripting.Dictionary")
    
    ' Varmistetaan, että datasivutaulukko on valittuna
    DATA.Select
    
    ' Yhdistetään käynnissä olevaan AutoCAD-instanssiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Poista Blokit"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    oACAD.Preferences.System.SingleDocumentMode = False
  
    For Each Rivi In Selection.Rows
        If Not Kaydyt.Exists(Rivi.Row) Then
            RiviNo = Rivi.Row
            Kaydyt.Add RiviNo, True  ' Merkitään rivi käydyksi – Dictionary-haku on O(1) ja eksakti
            If AvaaDoc(RiviNo) Then
                Application.StatusBar = "Tuhotaan objektia rivillä: " & Rivi.Row
                Set oEntity = oDOC.HandleToObject(Cells(Rivi.Row, 4).Text)
                oEntity.Delete
            End If
        End If
    Next
  
Cleanup:
    On Error Resume Next
    If Not OliAuki Then
        If Not oDOC Is Nothing Then
            If Ver = 0 Then Ver = acNative
            oDOC.SaveAs oDOC.FullName, Ver
        End If
    End If
    If Not oACAD Is Nothing Then
        oACAD.Preferences.System.SingleDocumentMode = Docmode
    End If
    Application.StatusBar = False
    
    ' Vapautetaan objektit
    Set oEntity = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing
    
    MsgBox "Valitut objektit tuhottiin", vbInformation, "Poista Blokit"
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Poista Blokit"
    Resume Cleanup
End Sub

Private Function OtsS(Nimi As String) As Long '' Muutettu Integer → Long
'' Etsii tai luo sarakkeen attribuutin nimelle
    Dim i As Long '' Muutettu Integer → Long
    Nimi = UCase(Nimi)
    i = 7
    Do
        If Cells(1, i).Value = "" Then
            Cells(1, i).Value = Nimi
            OtsS = i
            Exit Do
        ElseIf Cells(1, i).Value = Nimi Then
            OtsS = i
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Private Function AvaaDoc(Rivi As Long) As Boolean
'' Avaa AutoCAD-piirustuksen tarvittaessa
'' 26.10.2025 – Integer muutettu Longiksi

    Dim Doku As String
    Dim EdDoku As String
    Dim Tiedosto As String
    Dim i As Long '' Muutettu Integer → Long
    
    Doku = (Cells(Rivi, 2).Value) & ".dwg"
    EdDoku = (Cells(Rivi - 1, 2).Value) & ".dwg"
    Tiedosto = Cells(Rivi, 1).Value & "\" & Doku
    
    '' Tarkistetaan, onko haluttu piirustus jo aktiivinen
    If Not oDOC Is Nothing Then '' Jokin piirustus on auki
        If LCase(oDOC.Name) = LCase(Doku) Then '' Piirustus on käsiteltävä
            AvaaDoc = True
            Exit Function
        ElseIf LCase(oDOC.Name) = LCase(EdDoku) Then '' Edellinen piirustus on auki
            If Not OliAuki Then
                On Error Resume Next
                oDOC.Close True
                If Err.Number <> 0 Then
                    Err.Clear
                    MsgBox "Virhe talletettaessa piirustusta: " & oDOC.Name & vbCrLf & "Kuva saattaa olla jollakin auki.", vbCritical, "Vie tiedot"
                End If
                On Error GoTo 0
            End If
        End If
    End If
    
    '' Haluttua piirustusta ei käsitelty eikä se ole aktiivinen
    '' Tarkistetaan, onko piirustus avoinna AutoCADissä
    OliAuki = False
    For i = 0 To oACAD.Documents.Count - 1
        If LCase(oACAD.Documents(i).Name) = LCase(Doku) Then '' Piirustus on auki, asetetaan aktiiviseksi
            OliAuki = True
            oACAD.Documents(i).Activate
            Set oDOC = oACAD.ActiveDocument
            AvaaDoc = True
            Exit Function
        End If
    Next i
    
    '' Piirustus ei ollut käsittelyssä eikä auki AutoCADissä – avataan se
    On Error Resume Next
    Set oDOC = oACAD.Documents.Open(Tiedosto)
    
    If Err.Number <> 0 Then
        MsgBox "Virhe avattaessa piirustusta: " & Doku, vbCritical, "Vie tiedot"
        AvaaDoc = False
        Err.Clear
    Else
        AvaaDoc = True
    End If
    On Error GoTo 0
End Function

Sub Numerointi()
'' Numeerinen järjestely blokeille
'' 26.10.2025 – Integer muutettu Longiksi

    Dim Alku As String
    Dim Jakso As Long '' Muutettu Integer → Long
    Dim Vali As Long '' Muutettu Integer → Long
    Dim i As Long, j As Long '' Muutettu Integer → Long

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    
    '' Fetch from drawing
    TuoDATA True
    Alku = Aloitus.Range("D13").Value
    Vali = 2
    
    Cells.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("F2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ' Suorituskykysuojaus: estetään uudelleenpiirrto ja -laskenta rivittäisen kirjoituksen ajaksi
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo NumCleanup
    i = 2
    j = Val(Alku)
    Do
        If Cells(i, 1).Value = "" Then Exit Do
        Cells(i, 12).Value = LNumero(j, Alku)
        If Right(CStr(j), 1) = "8" Then
            j = j + Vali
        End If
        j = j + 1
        i = i + 1
    Loop
    
NumCleanup:
    ' Palautetaan Excel-asetukset aina, myös virhetilanteessa
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Virhe Numeroinnissa: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Numerointi"
        Exit Sub
    End If
    
    Aloitus.Range("D13").Value = LNumero(j, Alku)
    VieDATA
End Sub

Private Function LNumero(No As Long, Alku As String) As String
'' Muotoilee numeron johtavilla nollilla alkumerkkijonon pituuden mukaan
'' 26.10.2025 – Integer muutettu Longiksi

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
'' Viitenuerointi-työkalu
'' 26.10.2025 – Integer muutettu Longiksi

    Dim vSivu As Long '' Muutettu Integer → Long
    Dim Kirjain As String
    Dim i As Long, j As Long '' Muutettu Integer → Long

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    Kirjain = "A"
    
    '' Fetch from drawing
    TuoDATA True, "REFERENCE"
    vSivu = CLng(Aloitus.Range("D17").Value) '' Changed from CInt to CLng
  
    Cells.Sort Key1:=Range("F2"), Order1:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    i = 2
    Do
        If Cells(i, 1).Value = "" Then Exit Do
        Cells(i, 7).Value = "(" & vSivu & ":" & Kirjain & ") To Page " & vSivu
        Cells(i, 8).Value = "To Page " & vSivu + 1 & "(" & vSivu & ":" & Kirjain & ")"
        i = i + 1
        Kirjain = Chr(Asc(Kirjain) + 1)
    Loop
    
    Aloitus.Range("D18").Value = vSivu + 1
    VieDATA
End Sub

Function Lisaa(Nro As String, Maara As Long) As String
'' Lisää arvon numerojonoon säilyttäen muotoilun
'' 26.10.2025 – Integer muutettu Longiksi

    Dim Pit As Long '' Muutettu Integer → Long
    Dim i As Long   '' Muutettu Integer → Long
    
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

