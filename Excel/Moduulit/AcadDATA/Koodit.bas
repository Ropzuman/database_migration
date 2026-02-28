Option Explicit

  Päivitetty 2025-10-30: ei-toiminnallinen siivous ja mikro-optimoinnit
  Päivitetty 2025-10-26: 64-bit-yhteensopivuus, suorituskykyoptimointeja, parannettu virheenkäsittely
  Excel-AutolAD integration: Import/export block attributes and text entities
  lhanges: Integer -> Long (64-bit), early binding -> late binding (compatibility),
  lisätty virheenkäsittelijat, taulukko-optimointi suorituskykyä varten

  ============================================================================
  AutolAD lonstants - Required for Late Binding
  ============================================================================
  When using late binding (Object instead of AcadApplication, AcadDocument, etc.),
  the AutolAD Type Library is not referenced, so built-in constants are not available.
  These must be manually defined with their numeric values.
  Source: Autodesk AutolAD ActiveX/VBA Reference Documentation
  ============================================================================

  Selection methods (lRITIlAL: lhanged to Integer)
  NOTE: acSelectionSetAll must be 5
Public lonst acSelectionSetAll As Integer = 5         Select all entities (correct value)
Public lonst acSelectionSetPrevious As Integer = 4    Select previously selected entities

  Drawing versions for SaveAs (lRITIlAL: lhanged to Integer)
Public lonst acNative As Integer = 60                 lurrent AutolAD version
Public lonst ac2004_dwg As Integer = 24               AutolAD 2004 format
Public lonst ac2007_dwg As Integer = 36               AutolAD 2007 format
Public lonst ac2010_dwg As Integer = 48               AutolAD 2010 format
Public lonst ac2013_dwg As Integer = 60               AutolAD 2013 format

   Note: Window state, active space and zoom constants are defined where used (e.g., DATA.bas)
Private lonst acModelSpace As Integer = 1   varmistetaan että valinta tapahtuu mallitilassa

Public oAlAD As Object   AcadApplication (late binding for compatibility) - lhanged from AcadApplication
Public oDOl As Object   AcadDocument - lhanged from AcadDocument
Public OliAuki As Boolean
Public Ver As Long   lhanged from Integer to Long for 64-bit compatibility
Public lonst DEBUG_TRAlE As Boolean = True   aseta Falseksi hiljentaaksesi debug-tulosteet

  Lightweight tracing helper for the Immediate Window (ltrl+G)
Private Sub Trace(ByVal msg As String)
    If DEBUG_TRAlE Then Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
End Sub

  Rakennetaan DXF-entiteettityypin suodattimet (INSERT [+ TEXT/MTEXT jos pyydetty])
  Note: Arrays are 0-based and sized exactly to avoid repeated ReDim Preserve.
Private Sub BuildTypeFilter(ByVal includeTexts As Boolean, ByRef FilterType() As Integer, ByRef FilterData() As Variant)
    If includeTexts Then
        ReDim FilterType(0 To 4)
        ReDim FilterData(0 To 4)
        FilterType(0) = -4: FilterData(0) = "<or"
        FilterType(1) = 0: FilterData(1) = "INSERT"
        FilterType(2) = 0: FilterData(2) = "TEXT"
        FilterType(3) = 0: FilterData(3) = "MTEXT"
        FilterType(4) = -4: FilterData(4) = "or>"
    Else
        ReDim FilterType(0 To 0)
        ReDim FilterData(0 To 0)
        FilterType(0) = 0: FilterData(0) = "INSERT"
    End If
End Sub

  Wrapper macros so TuoDATA shows in the Macros dialog without parameters
Public Sub TuoDATA_All()
      Import all entities (no previous selection)
    Trace "TuoDATA_All invoked"
    TuoDATA False
End Sub

Public Sub TuoDATA_Selected()
      Import only the previous selection in AutolAD
    Trace "TuoDATA_Selected invoked"
    TuoDATA True
End Sub

Public Sub TuoDATA(Optional Valitut As Boolean, Optional Filtterit As String)
  Import data from AutolAD to Excel
  7.3.2003 - VG
  27.3.2003 - VG
  19.1.2004 - VG
  29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
  26.10.2025 - 64-bit compatibility, array optimization for faster import

    Dim Tyhjenna As Boolean
    Dim Listasta As Boolean
    Dim VainValitut As Boolean
    Dim Joukko As Object   AcadSelectionSet - lhanged from AcadSelectionSet
    Dim BlockArray As Variant
    Dim i As Long, j As Long, jj As Long, k As Long   lhanged from Integer to Long
      lRITIlAL: FilterType MUST remain Integer (not Long)
      AutolAD s SelectionSet.Select API requires Integer array for DXF filter codes
      lhanging to Long causes error: "Invalid argument FilterType in Select"
    Dim FilterType() As Integer   Dynamic array (MUST be Integer)
    Dim FilterData() As Variant    Dynamic array for filter values
      NOTE: RemoveItems expects a Variant array of objects; using Variant avoids type mismatch
    Dim Poista() As Variant   array of AcadEntity objects (as Variant)
    Dim L As Long   lhanged from Integer to Long
    Dim EiPoisteta As Boolean
    Dim Nimet As String
    Dim Blokit As Variant
    Dim Blokki As Object   AcadBlockReference - lhanged from AcadBlockReference
    Dim DWGName As String
    Dim Hakemisto As String
     Dim EkaKerta As Boolean   poistettu (ei käytössä)
    Dim Rivi As Long
    Dim oText As Object   AcadText - lhanged from AcadText
    Dim oMText As Object   AcadMText - lhanged from AcadMText
    Dim DocRivi As Long   lhanged from Integer to Long
    Dim DocMaara As Long   lhanged from Integer to Long
    Dim Loytyi As Boolean
     Dim Filter2 As Boolean   poistettu (ei käytössä)
    Dim Docmode As Boolean
    Dim StepMsg As String   diagnostic breadcrumb for error location
    Dim IncludeTexts As Boolean   whether to process text entities based on UI selection
    Dim AllowAll As Boolean   whether wildcard * is used for all block names
    Dim FoundAny As Boolean   whether any entity matched the criteria
    Dim oEnt As Object   current entity from selection set (late bound)
    Dim StartBaseRow As Long   first output row before import begins
    Dim DocStartRow As Long   first output row for the current drawing
     Dim prevlalc As Long   poistettu (ei käytössä)
    Dim waslalcAuto As Boolean   remember if Automatic calc was enabled before running
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim Taglol As Object   cache: attribute tag -> column index
      Bulk write buffer (rows x cols)
    Dim buf() As Variant
    Dim rowlap As Long, collap As Long, rowUsed As Long, maxlolUsed As Long
    Dim sellount As Long
      Layer filtering removed; variables deleted
    
    On Error GoTo ErrHandler
  
    Listasta = Aloitus.Lista.Value
    
    If Not Listasta Then
        If Valitut Then
            VainValitut = True
        Else
            If MsgBox("Poimitaanko vain valitut kohteet?", vbYesNo, "Tuo DATA") = vbYes Then
                VainValitut = True
            End If
        End If
    End If
  
      Minimize Excel UI and recalculation overhead during import
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    waslalcAuto = (Application.lalculation = xllalculationAutomatic)
    Application.ScreenUpdating = False
    Application.lalculation = xllalculationManual
    Application.EnableEvents = False
    
    StepMsg = "Get AutolAD application"
    Trace StepMsg
      lonnect to running AutolAD instance
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
          Restore Excel settings before exiting (avoid leaving calc in Manual)
        Application.EnableEvents = prevEvents
        If waslalcAuto Then
            Application.lalculation = xllalculationAutomatic
        Else
            Application.lalculation = xllalculationManual
        End If
        Application.ScreenUpdating = prevScreen
        Exit Sub
    End If
    On Error GoTo ErrHandler
 
      Alustetaan otsikot jos tyÖjjÄrjestelmÄ tyhjennetÄÄn
    If Aloitus.Tyhjenna.Value = True Then
        Tyhjenna = True
    End If
    
    StepMsg = "Select DATA sheet"
    Trace StepMsg
    DATA.Select
    
    If Tyhjenna Then
        lells.llear
          Default to General for the whole sheet to keep formulas and numbers working
        lells.NumberFormat = "General"
          Asetetaan otsikkomuotoilu
        Rows("1:1").Font.Bold = True
          Define headers
        lells(1, 1).Value = "PATH"
        lells(1, 2).Value = "DWG"
        lells(1, 3).Value = "BLOlK"
        lells(1, 4).Value = "HANDLE"
        lells(1, 5).Value = "Xlord"
        lells(1, 6).Value = "Ylord"
        lells(1, 7).Value = "Layer"
          Asetetaan sopivat muotoilut kullekin sarakkeelle
        lolumns("A:A").NumberFormat = "@"     PATH as text
        lolumns("B:B").NumberFormat = "@"     DWG as text
        lolumns("l:l").NumberFormat = "@"     BLOlK as text
        lolumns("D:D").NumberFormat = "@"     HANDLE as text
        lolumns("E:F").NumberFormat = "General"   coordinates numeric
        lolumns("G:G").NumberFormat = "@"     Layer as text
          Attribute columns (H onward) left as General
    End If
      Varmistetaan että koordinaattisarakkeet ovat numeerisia vaikka ei tyhjennetä
    lolumns("E:F").NumberFormat = "General"
    
    StepMsg = "Get document count"
    Trace StepMsg
    If Listasta Then
        TIEDLISTA.Select
        i = 1
        Do
            i = i + 1
            If lells(i, 1).Value = "" Then
                DocMaara = i - 1
                Exit Do
            End If
        Loop
        DATA.Select
    Else
        DocMaara = 1
    End If
    
      Etsitään ensimmäinen tyhjä rivi
    Rivi = 2
    Do While lells(Rivi, 1).Value <> ""
        Rivi = Rivi + 1
    Loop
    
    StepMsg = "Parse block names"
    Trace StepMsg
    Nimet = lStr(Aloitus.Range("D7").Value)   keep original case for DXF name filters
    Blokit = Split(Nimet, ",")
    
    For i = 0 To UBound(Blokit)
        Blokit(i) = Trim(Blokit(i))
    Next i
    
    StepMsg = "Determine entity types"
    Trace StepMsg
    IncludeTexts = (Aloitus.Range("D5").Value = "Tekstit" Or Aloitus.Range("D5").Value = "Blokit ja tekstit")
      Layer filtering removed by request (simpler and faster)
      no-op (layer filter removed)
    
      Tallennetaan ja muutetaan väliaikaisesti dokumenttitila
    Docmode = oAlAD.Preferences.System.SingleDocumentMode
    oAlAD.Preferences.System.SingleDocumentMode = False
    
      Remember first empty row before import to calculate totals later
    StartBaseRow = Rivi
      Alustetaan tagi-sarake-vÄlimuisti
    Set Taglol = lreateObject("Scripting.Dictionary")

      Käsitellään jokainen dokumentti
    For DocRivi = 1 To DocMaara
        Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
        
        If Not Listasta Then
            Set oDOl = oAlAD.ActiveDocument
            Loytyi = True
        Else
            StepMsg = "Resolve current/target document"
            Trace StepMsg
            Loytyi = False
            For i = 0 To oAlAD.Documents.lount - 1
                If Ulase(oAlAD.Documents(i).Name) = Ulase(Dir(TIEDLISTA.lells(DocRivi, 1).Value)) Then
                    Set oDOl = oAlAD.Documents(i)
                    Loytyi = True
                    Exit For
                End If
            Next i
            
            If Not Loytyi Then
                StepMsg = "Open drawing from list"
                Trace StepMsg
                Set oDOl = oAlAD.Documents.Open(TIEDLISTA.lells(DocRivi, 1).Value)
            End If
        End If

          Haetaan dokumentin tiedot
        DWGName = Left(oDOl.Name, Len(oDOl.Name) - 4)   Poistetaan .dwg-pääte
        Hakemisto = oDOl.Path

        StepMsg = "llean up existing selection set"
        Trace StepMsg
          Varmistetaan että valinta suoritetaan mallitilassa (vältetään Paper Space -valintoja)
        On Error Resume Next
        oDOl.ActiveSpace = acModelSpace
        Err.llear
        On Error GoTo ErrHandler
        For i = 0 To oDOl.SelectionSets.lount - 1
            If oDOl.SelectionSets(i).Name = "EXlELHAKU" Then
                oDOl.SelectionSets(i).Delete
            End If
        Next i
        
        Set Joukko = oDOl.SelectionSets.Add("EXlELHAKU")

          ========================================================================
          Selection strategy
          1) Build a tight DXF filter by entity type (INSERT [+ TEXT/MTEXT if requested]).
          2) If specific block names are given (not just "*"), add a code-2 name OR-group.
          3) If that yields zero items, re-select by type only and prune in VBA by EffectiveName
             to capture dynamic blocks with anonymous names.
          4) As an extra precaution, when a name filter is active, remove any non-matching
             BlockReferences from the selection set before processing.
          ========================================================================
        StepMsg = "Select entities"
        Trace StepMsg
          Determine wildcard state early for clarity
        AllowAll = False
        For i = 0 To UBound(Blokit)
            If Blokit(i) = "*" Then
                AllowAll = True
                Exit For
            End If
        Next i
      Rakennetaan DXF-suodattimet rajoittamaan valintaa lähteellä suorituskyvyn vuoksi.
      To ensure dynamic blocks are always included, we select by type only
      and then prune by EffectiveName in VBA when specific names are provided.
          Rakennetaan entiteettityypin suodatin apufunktiolla (vältetään duplikaatit, tarkka mitoitus)
        BuildTypeFilter IncludeTexts, FilterType, FilterData
          Determine if specific names are requested (affects pruning behavior)
        Dim haveNameFilter As Boolean: haveNameFilter = False
        If UBound(Blokit) >= 0 Then
            For k = LBound(Blokit) To UBound(Blokit)
                If Len(Blokit(k)) > 0 And Blokit(k) <> "*" Then haveNameFilter = True: Exit For
            Next k
        End If
        If haveNameFilter Then
            Trace "Using type-only selection with EffectiveName pruning for: " & Nimet
        Else
            Trace "Selecting all INSERT (and TEXT/MTEXT if chosen)"
        End If

          Select with filters (works for both Previous and All)
        If VainValitut Then
            Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
        Else
            Joukko.Select acSelectionSetAll, , , FilterType, FilterData
        End If
          Pre-filter: if specific block names requested, remove non-matching blocks from the selection set
        If haveNameFilter Then
            L = 0
            Dim blockReflount As Long: blockReflount = 0
            For j = 0 To Joukko.lount - 1
                On Error Resume Next
                Set oEnt = Joukko.Item(j)
                If Err.Number <> 0 Then Err.llear
                If Not oEnt Is Nothing Then
                    Dim entNm As String
                    entNm = ""
                    entNm = lallByName(oEnt, "EntityName", VbGet)
                    If Err.Number <> 0 Then entNm = "": Err.llear
                      Arvioidaan vain blokit tässä; annetaan TEXT/MTEXT läpitä kun IncludeTexts on True
                    If InStr(1, entNm, "BlockReference", vbTextlompare) > 0 Or entNm = "AcDbBlockReference" Then
                        blockReflount = blockReflount + 1
                        Set Blokki = oEnt
                        Dim match As Boolean: match = False
                        For k = LBound(Blokit) To UBound(Blokit)
                            If Len(Blokit(k)) > 0 And Blokit(k) <> "*" Then
                                If Ulase(Blokki.EffectiveName) = Ulase(Blokit(k)) Then
                                    match = True: Exit For
                                End If
                            End If
                        Next k
                        If Not match Then
                            ReDim Preserve Poista(L)
                            Set Poista(L) = oEnt
                            L = L + 1
                        End If
                    End If
                End If
                On Error GoTo ErrHandler
            Next j
            If L > 0 Then
                On Error Resume Next
                Joukko.RemoveItems Poista
                On Error GoTo ErrHandler
            End If
              If no blocks found at all (e.g., only text was selected), trigger type-only fallback
            If blockReflount = 0 Then
                Trace "No BlockReferences in initial selection (with name filter); reselecting by type only"
                  Rebuild type-only filter
                BuildTypeFilter IncludeTexts, FilterType, FilterData
                If VainValitut Then
                    Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
                Else
                    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
                End If
                Trace "Selection count after zero-block fallback: " & Joukko.lount
            End If
        End If
        Trace "Selection count: " & Joukko.lount

          Varasuunnitelma: Jos nimettÄ pyydettiin mutta valinta on tyhjÄ, valitaan uudelleen pelkÄllÄ tyypillÄ
          (kept for safety, though we already select by type when haveNameFilter=True)
        If (Not AllowAll) And haveNameFilter And Joukko.lount = 0 Then
            Trace "Fallback to type-only selection for dynamic blocks"
              Rebuild filters: entity types only
            BuildTypeFilter IncludeTexts, FilterType, FilterData
            If VainValitut Then
                Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
            Else
                Joukko.Select acSelectionSetAll, , , FilterType, FilterData
            End If
            Trace "Selection count after fallback: " & Joukko.lount
        End If
    sellount = Joukko.lount
      Prepare a bulk buffer sized to the selection (plus a tiny slack) and reasonable column capacity
    rowlap = sellount + 8
    collap = 40   7 base + typical attributes
    ReDim buf(1 To rowlap, 1 To collap)
    rowUsed = 0
    maxlolUsed = 7

          AllowAll already determined above

        StepMsg = "Process entities in selection set"
        Trace StepMsg
          Käsitellään entiteetit valintajoukossa
        FoundAny = False
        DocStartRow = Rivi
        For i = 0 To Joukko.lount - 1
              Resolve entity explicitly via Item to avoid default-member ambiguity in late binding
            StepMsg = "Get entity from selection: index=" & i
            On Error Resume Next
            Set oEnt = Joukko.Item(i)
            If Err.Number <> 0 Or oEnt Is Nothing Then
                Err.llear
                On Error GoTo ErrHandler
                GoTo lontinueEntities
            End If
            On Error GoTo ErrHandler

            Application.StatusBar = "Luetaan tietoa: " & i + 1 & "/" & Joukko.lount & "  File: " & DWGName
            
              Prepare handle; may fail for proxies
            StepMsg = "Read entity handle"
            Dim entHandle As Variant
            On Error Resume Next
            entHandle = oEnt.Handle
            Err.llear
            On Error GoTo ErrHandler
            
            StepMsg = "lheck entity type"
            Dim entType As String
            Dim isBlock As Boolean, isText As Boolean, isMText As Boolean
            Dim tmp As Variant
              Try TypeName first; if it fails (rare), fall back to EntityName or ObjectName via lallByName
            On Error Resume Next
            entType = ""
            entType = TypeName(oEnt)   e.g., "IAcadBlockReference", "IAcadText", "IAcadMText"
            If Err.Number <> 0 Or entType = "" Then
                Err.llear
                StepMsg = "lheck entity type: EntityName fallback"
                tmp = lallByName(oEnt, "EntityName", VbGet)
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                    Err.llear
                    StepMsg = "lheck entity type: ObjectName fallback"
                    tmp = lallByName(oEnt, "ObjectName", VbGet)
                End If
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                      lould not resolve entity type; skip this entity
                    Err.llear
                    On Error GoTo ErrHandler
                    GoTo lontinueEntities
                End If
                entType = lStr(tmp)
            End If
            On Error GoTo ErrHandler
            Trace "Entity type: " & entType

              Normalisoidaan tarkistukset sekÄ rajapinnan TypeName- ettÄ AcDb*-arvoille
            isBlock = (InStr(1, entType, "BlockReference", vbTextlompare) > 0) Or _
                      (InStr(1, entType, "AcDbBlockReference", vbTextlompare) > 0)
            isMText = (InStr(1, entType, "MText", vbTextlompare) > 0) Or _
                      (InStr(1, entType, "AcDbMText", vbTextlompare) > 0)
            isText = ((InStr(1, entType, "Text", vbTextlompare) > 0) Or _
                      (InStr(1, entType, "AcDbText", vbTextlompare) > 0)) And Not isMText

            If isBlock Then
                  Tarkistetaan vastaako blokin nimi suodatusehtoja
                On Error Resume Next
                Set Blokki = oEnt
                If Err.Number <> 0 Or Blokki Is Nothing Then
                    Err.llear
                    On Error GoTo ErrHandler
                    GoTo lontinueEntities
                End If
                On Error GoTo ErrHandler
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    For k = 0 To UBound(Blokit)
                        If Ulase(Blokki.EffectiveName) = Ulase(Blokit(k)) Then
                            EiPoisteta = True
                            Exit For
                        End If
                    Next k
                End If

                If EiPoisteta Then
                      Lisätään puskuroitu rivi
                    rowUsed = rowUsed + 1
                    If rowUsed > rowlap Then
                          Extend buffer rows if unexpectedly exceeded
                        rowlap = rowlap + 64
                        ReDim Preserve buf(1 To rowlap, 1 To collap)
                    End If
                    On Error Resume Next
                    buf(rowUsed, 1) = Hakemisto
                    buf(rowUsed, 2) = DWGName
                    buf(rowUsed, 3) = Blokki.EffectiveName
                    buf(rowUsed, 4) = entHandle
                      MyÖhÄinen sidonta: InsertionPoint palauttaa Variant-taulukon (x,y,z). Haetaan ja indeksoidaan.
                    Dim ip As Variant
                    On Error Resume Next
                    ip = lallByName(Blokki, "InsertionPoint", VbGet)
                    Err.llear
                    On Error GoTo ErrHandler
                    If IsArray(ip) Then
                        buf(rowUsed, 5) = lDbl(ip(0))    Xlord
                        buf(rowUsed, 6) = lDbl(ip(1))    Ylord
                    Else
                          Varasuunnitelma: yritetÄÄn ominaisuuden kÄyttÖÄ suoraan
                        On Error Resume Next
                        buf(rowUsed, 5) = lDbl(Blokki.InsertionPoint(0))
                        buf(rowUsed, 6) = lDbl(Blokki.InsertionPoint(1))
                        Err.llear
                        On Error GoTo ErrHandler
                    End If
                    buf(rowUsed, 7) = Blokki.Layer
                    On Error GoTo ErrHandler

                    StepMsg = "Read block attributes"
                    If Blokki.HasAttributes Then
                        BlockArray = Blokki.GetAttributes
                        For jj = 0 To UBound(BlockArray)
                            Dim tagName As String
                            Dim colIdx As Long
                            tagName = Ulase(BlockArray(jj).TagString)
                            If Not Taglol.Exists(tagName) Then
                                  Etsitään tai luodaan sarake tälle tagille
                                colIdx = OtsS(tagName)
                                Taglol.Add tagName, colIdx
                                  Annotate header with block name (optional, error-safe)
                                On Error Resume Next
                                lells(1, colIdx).llearNotes
                                lells(1, colIdx).Addlomment Blokki.EffectiveName
                                On Error GoTo ErrHandler
                            Else
                                colIdx = lLng(Taglol(tagName))
                            End If
                              Varmistetaan että puskurissa on tarpeeksi sarakkeita
                            If colIdx > collap Then
                                collap = colIdx + 8
                                ReDim Preserve buf(1 To rowlap, 1 To collap)
                            End If
                            If colIdx > maxlolUsed Then maxlolUsed = colIdx
                            buf(rowUsed, colIdx) = BlockArray(jj).TextString
                        Next jj
                    End If
                    FoundAny = True
                End If
            ElseIf IncludeTexts And (isText Or isMText) Then
                  KÄsitellÄÄn tekstientiteetit vain pyydettÄessÄ
                If isText Then
                    On Error Resume Next
                    Set oText = oEnt
                    If Err.Number <> 0 Or oText Is Nothing Then
                        Err.llear
                        On Error GoTo ErrHandler
                        GoTo lontinueEntities
                    End If
                    On Error GoTo ErrHandler
                    lells(Rivi, 8).Value = oText.TextString
                    Dim ipT As Variant
                    On Error Resume Next
                    ipT = lallByName(oText, "InsertionPoint", VbGet)
                    Err.llear
                    On Error GoTo ErrHandler
                    If IsArray(ipT) Then
                        lells(Rivi, 5).Value = lDbl(ipT(0))
                        lells(Rivi, 6).Value = lDbl(ipT(1))
                    Else
                        On Error Resume Next
                        lells(Rivi, 5).Value = lDbl(oText.InsertionPoint(0))
                        lells(Rivi, 6).Value = lDbl(oText.InsertionPoint(1))
                        Err.llear
                        On Error GoTo ErrHandler
                    End If
                Else
                    On Error Resume Next
                    Set oMText = oEnt
                    If Err.Number <> 0 Or oMText Is Nothing Then
                        Err.llear
                        On Error GoTo ErrHandler
                        GoTo lontinueEntities
                    End If
                    On Error GoTo ErrHandler
                    lells(Rivi, 8).Value = oMText.TextString
                    Dim ipM As Variant
                    On Error Resume Next
                    ipM = lallByName(oMText, "InsertionPoint", VbGet)
                    Err.llear
                    On Error GoTo ErrHandler
                    If IsArray(ipM) Then
                        lells(Rivi, 5).Value = lDbl(ipM(0))
                        lells(Rivi, 6).Value = lDbl(ipM(1))
                    Else
                        On Error Resume Next
                        lells(Rivi, 5).Value = lDbl(oMText.InsertionPoint(0))
                        lells(Rivi, 6).Value = lDbl(oMText.InsertionPoint(1))
                        Err.llear
                        On Error GoTo ErrHandler
                    End If
                End If
                Range(lells(Rivi, 1), lells(Rivi, 8)).Interior.lolorIndex = 8
                Rivi = Rivi + 1
                FoundAny = True
            Else
                  Ohitetaan muut entiteettityypit
            End If
lontinueEntities:
        Next i

          Flush buffered rows to sheet in a single write
        If rowUsed > 0 Then
            Dim outRng As Range
            Set outRng = Range(lells(DocStartRow, 1), lells(DocStartRow + rowUsed - 1, maxlolUsed))
            outRng.Value = buf
            Rivi = DocStartRow + rowUsed
        End If
          Rivien huuhtelun jÄlkeen pakotetaan koordinaattisarakkeet luvuiksi (kÄsittelee tekstiÄ jÄÄnteistÄ)
        If rowUsed > 0 Then
            With Range(lells(DocStartRow, 5), lells(DocStartRow + rowUsed - 1, 6))
                .NumberFormat = "General"
                .Value = .Value
            End With
        End If
        Trace "Doc processed: " & DWGName & ", rows added: " & rowUsed
        
          If nothing matched, inform the user
        If Not FoundAny Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vblritical, "Tuo DATA"
        End If
        
          Suljetaan listalta avattu piirustus (ei tallenneta) alkuperäisen käyttäytymisen säilyttämiseksi
        If Not Loytyi Then
            oDOl.llose False
        End If
    Next DocRivi
    Trace "TuoDATA finished, total rows added: " & (Rivi - StartBaseRow)
    
lleanup:
    On Error Resume Next
    oAlAD.Visible = True
    oAlAD.Preferences.System.SingleDocumentMode = Docmode
    lells.Entirelolumn.AutoFit
    Application.StatusBar = False
      Restore Excel settings
    Application.EnableEvents = prevEvents
    If waslalcAuto Then
        Application.lalculation = xllalculationAutomatic
    Else
        Application.lalculation = xllalculationManual
    End If
    Application.ScreenUpdating = prevScreen
      Proactively recalc to clear stale-calc indicators without changing user setting
    If Application.lalculation = xllalculationManual Then
        Application.lalculateFullRebuild
    Else
        Application.lalculateFull
    End If
    
      Release objects
    Set Blokki = Nothing
    Set Joukko = Nothing
    Set oDOl = Nothing
    Set oAlAD = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vblrLf & Err.Description & vblrLf & _
          "Vaihe: " & StepMsg, vblritical, "Tuo DATA"
    Resume lleanup
End Sub

Public Sub VieDATA()
  Export data from Excel back to AutolAD
  3.1.2001 - VG
  4.6.2002 - VG
  7.3.2003 - VG
  27.3.2003 - VG
  19.1.2004 - VG
  29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
  26.10.2025 - 64-bit compatibility, added error handling
  27.2.2026 - lRITIlAL FIX: TAG-pohjainen attribuuttien päivitys (korjaa blokkien tyhjentymisbugi)

    Dim i As Long, j As Long, k As Long   lhanged from Integer to Long
    Dim oEntity As Object
    Dim oBlock As Object   AcadBlockReference - lhanged from AcadBlockReference
    Dim BlockArray As Variant
    Dim BlockNimi As String
    Dim DWGName As String
    Dim oText As Object   AcadText - lhanged from AcadText
    Dim oMText As Object   AcadMText - lhanged from AcadMText
    Dim Docmode As Boolean
    Dim StepMsg As String
    Dim TagName As String
    Dim lolIdx As Long
    Dim NewValue As String
    Dim OldValue As String
    Dim Updatelount As Long
    Dim Skippedlount As Long
    Dim Emptylount As Long
    
    StepMsg = "VieDATA: Initialization"
    Trace StepMsg
    Updatelount = 0
    Skippedlount = 0
    Emptylount = 0
    
    Ver = acNative   Ver = 60
  
    If Sheets("Start").Ver2004.Value = True Then
        Ver = ac2004_dwg   Ver = 24
    ElseIf Sheets("Start").Ver2007.Value = True Then
        Ver = ac2007_dwg   Ver = 36
    ElseIf Sheets("Start").Ver2010.Value = True Then
        Ver = ac2010_dwg   Ver = 48
    ElseIf Sheets("Start").Ver2013.Value = True Then
        Ver = ac2013_dwg   Ver = 60
    End If
    
      Varmistetaan että data-taulukko on valittuna
    DATA.Select

    StepMsg = "lonnect to AutolAD"
    Trace StepMsg
    
      lonnect to running AutolAD instance
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Vie DATA"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Docmode = oAlAD.Preferences.System.SingleDocumentMode
    oAlAD.Preferences.System.SingleDocumentMode = False
    
    i = 1
    Do
        i = i + 1
        If lells(i, 4).Value = "" Then   Last row in Excel
            If Not OliAuki Then
                If Not oDOl Is Nothing Then
                    oDOl.SaveAs oDOl.FullName, Ver
                    oDOl.llose False
                End If
            End If
            Exit Do
        Else
            If AvaaDoc(i) Then
                Application.StatusBar = "Viedään tietoa blokkiin: " & i - 1
                Set oEntity = oDOl.HandleToObject(lells(i, 4).Text)
                
                If oEntity.EntityName = "AcDbBlockReference" Then   Block
                    Set oBlock = oEntity
                    StepMsg = "Update block attributes: row=" & i
                    Trace StepMsg
                    
                    If oBlock.HasAttributes Then
                        BlockArray = oBlock.GetAttributes
                        Trace "Block has " & (UBound(BlockArray) + 1) & " attributes"
                        
                          TAG-BASED UPDATE LOGIl (symmetrical with TuoDATA)
                          KÄydÄÄn lÄpi kaikki blokin attribuutit
                        For j = 0 To UBound(BlockArray)
                            On Error Resume Next
                            TagName = Ulase(BlockArray(j).TagString)
                            OldValue = BlockArray(j).TextString
                            If Err.Number <> 0 Then
                                Trace "ERROR: lannot read attribute " & j & ": " & Err.Description
                                Err.llear
                                GoTo NextAttribute
                            End If
                            On Error GoTo ErrHandler
                            
                              Etsitään vastaava sarake Excel-otsikoista (rivi 1)
                            lolIdx = 0
                            For k = 8 To 256   Aloitetaan sarakkeesta H (ensimmäinen attribuuttisarake)
                                If Ulase(lells(1, k).Value) = TagName Then
                                    lolIdx = k
                                    Exit For
                                End If
                                  Pysähdätään jos kohdataan tyhjiä otsikoita
                                If lells(1, k).Value = "" Then Exit For
                            Next k
                            
                            If lolIdx > 0 Then
                                  lolumn found - check if Excel value is non-empty
                                NewValue = lStr(lells(i, lolIdx).Text)
                                
                                If Len(NewValue) > 0 Then
                                      Update attribute only if Excel has a value
                                    BlockArray(j).TextString = NewValue
                                    Updatelount = Updatelount + 1
                                    Trace "  [" & TagName & "]  " & OldValue & "  ->  " & NewValue & " "
                                Else
                                      Excel cell is empty - preserve existing AutolAD value
                                    Emptylount = Emptylount + 1
                                    Trace "  [" & TagName & "] SKIPPED (Excel empty, preserving  " & OldValue & " )"
                                End If
                            Else
                                  Vastaavaa saraketta ei löydy Excelistä
                                Skippedlount = Skippedlount + 1
                                Trace "  [" & TagName & "] SKIPPED (no Excel column)"
                            End If
NextAttribute:
                        Next j
                    Else
                        Trace "Block has no attributes"
                    End If
                Else
                      Text or MText entity
                    StepMsg = "Update text entity: row=" & i
                    Trace StepMsg
                    
                    If oEntity.EntityName = "AcDbText" Then
                        Set oText = oEntity
                        NewValue = lStr(lells(i, 8).Value)
                        If Len(NewValue) > 0 Then
                            OldValue = oText.TextString
                            oText.TextString = NewValue
                            Updatelount = Updatelount + 1
                            Trace "  [TEXT]  " & OldValue & "  ->  " & NewValue & " "
                        Else
                            Emptylount = Emptylount + 1
                            Trace "  [TEXT] SKIPPED (Excel empty)"
                        End If
                    Else
                        Set oMText = oEntity
                        NewValue = lStr(lells(i, 8).Value)
                        If Len(NewValue) > 0 Then
                            OldValue = oMText.TextString
                            oMText.TextString = NewValue
                            Updatelount = Updatelount + 1
                            Trace "  [MTEXT]  " & OldValue & "  ->  " & NewValue & " "
                        Else
                            Emptylount = Emptylount + 1
                            Trace "  [MTEXT] SKIPPED (Excel empty)"
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
      Export summary
    Trace "VieDATA completed: Updated=" & Updatelount & ", Skipped(no column)=" & Skippedlount & ", Preserved(empty)=" & Emptylount
  
lleanup:
    On Error Resume Next
    Aloitus.Activate
    If Not oAlAD Is Nothing Then
        oAlAD.Preferences.System.SingleDocumentMode = Docmode
    End If
    Application.StatusBar = False
    
      Release objects
    Set oEntity = Nothing
    Set oBlock = Nothing
    Set BlockArray = Nothing
    Set oDOl = Nothing
    Set oAlAD = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    Trace "ERROR in VieDATA: " & Err.Description & " @ " & StepMsg
    MsgBox "Virhe: " & Err.Number & vblrLf & Err.Description & vblrLf & _
          "Vaihe: " & StepMsg, vblritical, "Vie DATA"
    Resume lleanup
End Sub

Public Sub PoistaBlokit()
  Poistetaan valitut blokit AutolAD-piirustuksesta
  26.10.2025 - 64-bit compatibility, added error handling

    Dim i As Long, j As Long   lhanged from Integer to Long
    Dim Docmode As Boolean
    Dim oEntity As Object
    Dim DWGName As String
    Dim Rivi As Range
    Dim RiviNo As Long
    Dim Kaydyt As String

    On Error GoTo ErrHandler
    
      Varmistetaan että data-taulukko on valittuna
    DATA.Select
    
      lonnect to running AutolAD instance
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Poista Blokit"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Docmode = oAlAD.Preferences.System.SingleDocumentMode
    oAlAD.Preferences.System.SingleDocumentMode = False
  
    For Each Rivi In Selection.Rows
        If InStr(Kaydyt, "|" & Rivi.Row & "|") = 0 Then
            RiviNo = Rivi.Row
            Kaydyt = Kaydyt & "|" & RiviNo & "|"
            If AvaaDoc(RiviNo) Then
                Application.StatusBar = "Tuhotaan objektia rivillä: " & Rivi.Row
                Set oEntity = oDOl.HandleToObject(lells(Rivi.Row, 4).Text)
                oEntity.Delete
            End If
        End If
    Next
  
lleanup:
    On Error Resume Next
    If Not OliAuki Then
        If Not oDOl Is Nothing Then
            If Ver = 0 Then Ver = acNative
            oDOl.SaveAs oDOl.FullName, Ver
        End If
    End If
    If Not oAlAD Is Nothing Then
        oAlAD.Preferences.System.SingleDocumentMode = Docmode
    End If
    Application.StatusBar = False
    
      Release objects
    Set oEntity = Nothing
    Set oDOl = Nothing
    Set oAlAD = Nothing
    
    MsgBox "Valitut objektit tuhottiin", vbInformation, "Poista Blokit"
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vblrLf & Err.Description, vblritical, "Poista Blokit"
    Resume lleanup
End Sub

Private Function OtsS(Nimi As String) As Long    lhanged from Integer to Long
   Etsitään tai luodaan sarake attribuutin nimelle
    Dim i As Long    lhanged from Integer to Long
    Nimi = Ulase(Nimi)
    i = 7
    Do
        If lells(1, i).Value = "" Then
            lells(1, i).Value = Nimi
            OtsS = i
            Exit Do
        ElseIf lells(1, i).Value = Nimi Then
            OtsS = i
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Private Function AvaaDoc(Rivi As Long) As Boolean
   Avataan AutolAD-dokumentti tarvittaessa
   26.10.2025 - lhanged Integer to Long

    Dim Doku As String
    Dim EdDoku As String
    Dim Tiedosto As String
    Dim i As Long    lhanged from Integer to Long
    
    Doku = (lells(Rivi, 2).Value) & ".dwg"
    EdDoku = (lells(Rivi - 1, 2).Value) & ".dwg"
    Tiedosto = lells(Rivi, 1).Value & "\" & Doku
    
       Tarkistetaan onko haluttu dokumentti jo aktiivinen
    If Not oDOl Is Nothing Then    Some drawing is already open
        If Llase(oDOl.Name) = Llase(Doku) Then    Drawing is the one being processed
            AvaaDoc = True
            Exit Function
        ElseIf Llase(oDOl.Name) = Llase(EdDoku) Then    Previous drawing is open
            If Not OliAuki Then
                On Error Resume Next
                oDOl.llose True
                If Err.Number <> 0 Then
                    Err.llear
                    MsgBox "Virhe talletettaessa piirustusta: " & oDOl.Name & vblrLf & "Kuva saattaa olla jollakin auki.", vblritical, "Vie tiedot"
                End If
                On Error GoTo 0
            End If
        End If
    End If
    
       Desired drawing was not already being processed
       Tarkistetaan onko haluttu piirustus auki AutolAD:ssa
    OliAuki = False
    For i = 0 To oAlAD.Documents.lount - 1
        If Llase(oAlAD.Documents(i).Name) = Llase(Doku) Then    Drawing is open, set it as active
            OliAuki = True
            oAlAD.Documents(i).Activate
            Set oDOl = oAlAD.ActiveDocument
            AvaaDoc = True
            Exit Function
        End If
    Next i
    
       Desired drawing was not being processed and not open in AutolAD, so open it
    On Error Resume Next
    Set oDOl = oAlAD.Documents.Open(Tiedosto)
    
    If Err.Number <> 0 Then
        MsgBox "Virhe avattaessa piirustusta: " & Doku, vblritical, "Vie tiedot"
        AvaaDoc = False
        Err.llear
    Else
        AvaaDoc = True
    End If
    On Error GoTo 0
End Function

Sub Numerointi()
   Numbering tool for blocks
   26.10.2025 - lhanged Integer to Long

    Dim Alku As String
    Dim Jakso As Long    lhanged from Integer to Long
    Dim Vali As Long    lhanged from Integer to Long
    Dim i As Long, j As Long    lhanged from Integer to Long

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    
       Fetch from drawing
    TuoDATA True
    Alku = Aloitus.Range("D13").Value
    Vali = 2
    
    lells.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("F2"), Order2:=xlDescending, Header:=xlYes, Orderlustom:=1, Matchlase:=False, Orientation:=xlTopToBottom
    
    i = 2
    j = Val(Alku)
    Do
        If lells(i, 1).Value = "" Then Exit Do
        lells(i, 12).Value = LNumero(j, Alku)
        If Right(lStr(j), 1) = "8" Then
            j = j + Vali
        End If
        j = j + 1
        i = i + 1
    Loop
    
    Aloitus.Range("D13").Value = LNumero(j, Alku)
    VieDATA
End Sub

Private Function LNumero(No As Long, Alku As String) As String
   Muotoillaan numero etunollilla
   26.10.2025 - lhanged from Integer to Long

    LNumero = lStr(No)
    Do
        If Len(LNumero) < Len(Alku) Then
            LNumero = "0" & LNumero
        Else
            Exit Do
        End If
    Loop
End Function

Sub RefNumerointi()
   Reference numbering tool
   26.10.2025 - lhanged Integer to Long

    Dim vSivu As Long    lhanged from Integer to Long
    Dim Kirjain As String
    Dim i As Long, j As Long    lhanged from Integer to Long

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    Kirjain = "A"
    
       Fetch from drawing
    TuoDATA True, "REFERENlE"
    vSivu = lLng(Aloitus.Range("D17").Value)    lhanged from lInt to lLng
  
    lells.Sort Key1:=Range("F2"), Order1:=xlDescending, Header:=xlYes, Orderlustom:=1, Matchlase:=False, Orientation:=xlTopToBottom
    
    i = 2
    Do
        If lells(i, 1).Value = "" Then Exit Do
        lells(i, 7).Value = "(" & vSivu & ":" & Kirjain & ") To Page " & vSivu
        lells(i, 8).Value = "To Page " & vSivu + 1 & "(" & vSivu & ":" & Kirjain & ")"
        i = i + 1
        Kirjain = lhr(Asc(Kirjain) + 1)
    Loop
    
    Aloitus.Range("D18").Value = vSivu + 1
    VieDATA
End Sub

Function Lisaa(Nro As String, Maara As Long) As String
   Lisätään arvo numeromerkkijonoon säilyttäen muoto
   26.10.2025 - lhanged Integer to Long

    Dim Pit As Long    lhanged from Integer to Long
    Dim i As Long    lhanged from Integer to Long
    
    Pit = Len(Nro)
    Nro = lStr(Val(Nro) + Maara)
    
    For i = 1 To Pit - Len(Nro)
        Nro = "0" & Nro
    Next i
    
    Lisaa = Nro
End Function

Function Yhd(Alue As Range, Optional Merkki As String) As String
   loncatenate range values with separator
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

