Option Explicit

' Updated 2025-10-30: non-functional cleanup and micro-optimizations
' Updated 2025-10-26: 64-bit compatibility, performance optimizations, improved error handling
' Excel-AutoCAD integration: Import/export block attributes and text entities
' Changes: Integer -> Long (64-bit), early binding -> late binding (compatibility),
' added error handlers, array optimization for performance

' ============================================================================
' AutoCAD Constants - Required for Late Binding
' ============================================================================
' When using late binding (Object instead of AcadApplication, AcadDocument, etc.),
' the AutoCAD Type Library is not referenced, so built-in constants are not available.
' These must be manually defined with their numeric values.
' Source: Autodesk AutoCAD ActiveX/VBA Reference Documentation
' ============================================================================

' Selection methods (CRITICAL: Changed to Integer)
' NOTE: acSelectionSetAll must be 5
Public Const acSelectionSetAll As Integer = 5       ' Select all entities (correct value)
Public Const acSelectionSetPrevious As Integer = 4  ' Select previously selected entities

' Drawing versions for SaveAs (CRITICAL: Changed to Integer)
Public Const acNative As Integer = 60               ' Current AutoCAD version
Public Const ac2004_dwg As Integer = 24             ' AutoCAD 2004 format
Public Const ac2007_dwg As Integer = 36             ' AutoCAD 2007 format
Public Const ac2010_dwg As Integer = 48             ' AutoCAD 2010 format
Public Const ac2013_dwg As Integer = 60             ' AutoCAD 2013 format

'' Note: Window state, active space and zoom constants are defined where used (e.g., DATA.bas)
Private Const acModelSpace As Integer = 1 ' ensure selection happens in Model Space

Public oACAD As Object ' AcadApplication (late binding for compatibility) - Changed from AcadApplication
Public oDOC As Object ' AcadDocument - Changed from AcadDocument
Public OliAuki As Boolean
Public Ver As Long ' Changed from Integer to Long for 64-bit compatibility
Public Const DEBUG_TRACE As Boolean = True ' set False to silence debug prints

' Lightweight tracing helper for the Immediate Window (Ctrl+G)
Private Sub Trace(ByVal msg As String)
    If DEBUG_TRACE Then Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
End Sub

' Build DXF entity-type filters (INSERT [+ TEXT/MTEXT if requested])
' Note: Arrays are 0-based and sized exactly to avoid repeated ReDim Preserve.
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

' Wrapper macros so TuoDATA shows in the Macros dialog without parameters
Public Sub TuoDATA_All()
    ' Import all entities (no previous selection)
    Trace "TuoDATA_All invoked"
    TuoDATA False
End Sub

Public Sub TuoDATA_Selected()
    ' Import only the previous selection in AutoCAD
    Trace "TuoDATA_Selected invoked"
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
    Dim Joukko As Object ' AcadSelectionSet - Changed from AcadSelectionSet
    Dim BlockArray As Variant
    Dim i As Long, j As Long, jj As Long, k As Long ' Changed from Integer to Long
    ' CRITICAL: FilterType MUST remain Integer (not Long)
    ' AutoCAD's SelectionSet.Select API requires Integer array for DXF filter codes
    ' Changing to Long causes error: "Invalid argument FilterType in Select"
    Dim FilterType() As Integer ' Dynamic array (MUST be Integer)
    Dim FilterData() As Variant  ' Dynamic array for filter values
    ' NOTE: RemoveItems expects a Variant array of objects; using Variant avoids type mismatch
    Dim Poista() As Variant ' array of AcadEntity objects (as Variant)
    Dim L As Long ' Changed from Integer to Long
    Dim EiPoisteta As Boolean
    Dim Nimet As String
    Dim Blokit As Variant
    Dim Blokki As Object ' AcadBlockReference - Changed from AcadBlockReference
    Dim DWGName As String
    Dim Hakemisto As String
    'Dim EkaKerta As Boolean ' removed (unused)
    Dim Rivi As Long
    Dim oText As Object ' AcadText - Changed from AcadText
    Dim oMText As Object ' AcadMText - Changed from AcadMText
    Dim DocRivi As Long ' Changed from Integer to Long
    Dim DocMaara As Long ' Changed from Integer to Long
    Dim Loytyi As Boolean
    'Dim Filter2 As Boolean ' removed (unused)
    Dim Docmode As Boolean
    Dim StepMsg As String ' diagnostic breadcrumb for error location
    Dim IncludeTexts As Boolean ' whether to process text entities based on UI selection
    Dim AllowAll As Boolean ' whether wildcard * is used for all block names
    Dim FoundAny As Boolean ' whether any entity matched the criteria
    Dim oEnt As Object ' current entity from selection set (late bound)
    Dim StartBaseRow As Long ' first output row before import begins
    Dim DocStartRow As Long ' first output row for the current drawing
    'Dim prevCalc As Long ' removed (unused)
    Dim wasCalcAuto As Boolean ' remember if Automatic calc was enabled before running
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim TagCol As Object ' cache: attribute tag -> column index
    ' Bulk write buffer (rows x cols)
    Dim buf() As Variant
    Dim rowCap As Long, colCap As Long, rowUsed As Long, maxColUsed As Long
    Dim selCount As Long
    ' Layer filtering removed; variables deleted
    
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
  
    ' Minimize Excel UI and recalculation overhead during import
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    wasCalcAuto = (Application.Calculation = xlCalculationAutomatic)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    StepMsg = "Get AutoCAD application"
    Trace StepMsg
    ' Connect to running AutoCAD instance
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
        ' Restore Excel settings before exiting (avoid leaving calc in Manual)
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
 
    ' Initialize headers if clearing worksheet
    If Aloitus.Tyhjenna.Value = True Then
        Tyhjenna = True
    End If
    
    StepMsg = "Select DATA sheet"
    Trace StepMsg
    DATA.Select
    
    If Tyhjenna Then
        Cells.Clear
        ' Default to General for the whole sheet to keep formulas and numbers working
        Cells.NumberFormat = "General"
        ' Set header styling
        Rows("1:1").Font.Bold = True
        ' Define headers
        Cells(1, 1).Value = "PATH"
        Cells(1, 2).Value = "DWG"
        Cells(1, 3).Value = "BLOCK"
        Cells(1, 4).Value = "HANDLE"
        Cells(1, 5).Value = "XCord"
        Cells(1, 6).Value = "YCord"
        Cells(1, 7).Value = "Layer"
        ' Set appropriate formats per column
        Columns("A:A").NumberFormat = "@"   ' PATH as text
        Columns("B:B").NumberFormat = "@"   ' DWG as text
        Columns("C:C").NumberFormat = "@"   ' BLOCK as text
        Columns("D:D").NumberFormat = "@"   ' HANDLE as text
        Columns("E:F").NumberFormat = "General" ' coordinates numeric
        Columns("G:G").NumberFormat = "@"   ' Layer as text
        ' Attribute columns (H onward) left as General
    End If
    ' Ensure coordinate columns are numeric even when not clearing
    Columns("E:F").NumberFormat = "General"
    
    StepMsg = "Get document count"
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
    
    ' Find first empty row
    Rivi = 2
    Do While Cells(Rivi, 1).Value <> ""
        Rivi = Rivi + 1
    Loop
    
    StepMsg = "Parse block names"
    Trace StepMsg
    Nimet = CStr(Aloitus.Range("D7").Value) ' keep original case for DXF name filters
    Blokit = Split(Nimet, ",")
    
    For i = 0 To UBound(Blokit)
        Blokit(i) = Trim(Blokit(i))
    Next i
    
    StepMsg = "Determine entity types"
    Trace StepMsg
    IncludeTexts = (Aloitus.Range("D5").Value = "Tekstit" Or Aloitus.Range("D5").Value = "Blokit ja tekstit")
    ' Layer filtering removed by request (simpler and faster)
    ' no-op (layer filter removed)
    
    ' Save and temporarily change document mode
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    oACAD.Preferences.System.SingleDocumentMode = False
    
    ' Remember first empty row before import to calculate totals later
    StartBaseRow = Rivi
    ' Initialize tag-to-column cache
    Set TagCol = CreateObject("Scripting.Dictionary")

    ' Process each document
    For DocRivi = 1 To DocMaara
        Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
        
        If Not Listasta Then
            Set oDOC = oACAD.ActiveDocument
            Loytyi = True
        Else
            StepMsg = "Resolve current/target document"
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

        StepMsg = "Clean up existing selection set"
        Trace StepMsg
        ' Ensure selection is executed in Model Space (avoid Paper Space-only picks)
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

        ' ========================================================================
        ' Selection strategy
        ' 1) Build a tight DXF filter by entity type (INSERT [+ TEXT/MTEXT if requested]).
        ' 2) If specific block names are given (not just "*"), add a code-2 name OR-group.
        ' 3) If that yields zero items, re-select by type only and prune in VBA by EffectiveName
        '    to capture dynamic blocks with anonymous names.
        ' 4) As an extra precaution, when a name filter is active, remove any non-matching
        '    BlockReferences from the selection set before processing.
        ' ========================================================================
        StepMsg = "Select entities"
        Trace StepMsg
        ' Determine wildcard state early for clarity
        AllowAll = False
        For i = 0 To UBound(Blokit)
            If Blokit(i) = "*" Then
                AllowAll = True
                Exit For
            End If
        Next i
    ' Build DXF filters to limit selection at source for performance.
    ' To ensure dynamic blocks are always included, we select by type only
    ' and then prune by EffectiveName in VBA when specific names are provided.
        ' Build entity type filter using helper (avoids duplication, exact sizing)
        BuildTypeFilter IncludeTexts, FilterType, FilterData
        ' Determine if specific names are requested (affects pruning behavior)
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

        ' Select with filters (works for both Previous and All)
        If VainValitut Then
            Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
        Else
            Joukko.Select acSelectionSetAll, , , FilterType, FilterData
        End If
        ' Pre-filter: if specific block names requested, remove non-matching blocks from the selection set
        If haveNameFilter Then
            L = 0
            Dim blockRefCount As Long: blockRefCount = 0
            For j = 0 To Joukko.Count - 1
                On Error Resume Next
                Set oEnt = Joukko.Item(j)
                If Err.Number <> 0 Then Err.Clear
                If Not oEnt Is Nothing Then
                    Dim entNm As String
                    entNm = ""
                    entNm = CallByName(oEnt, "EntityName", VbGet)
                    If Err.Number <> 0 Then entNm = "": Err.Clear
                    ' Only evaluate blocks here; allow TEXT/MTEXT to pass when IncludeTexts is True
                    If InStr(1, entNm, "BlockReference", vbTextCompare) > 0 Or entNm = "AcDbBlockReference" Then
                        blockRefCount = blockRefCount + 1
                        Set Blokki = oEnt
                        Dim match As Boolean: match = False
                        For k = LBound(Blokit) To UBound(Blokit)
                            If Len(Blokit(k)) > 0 And Blokit(k) <> "*" Then
                                If UCase(Blokki.EffectiveName) = UCase(Blokit(k)) Then
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
            ' If no blocks found at all (e.g., only text was selected), trigger type-only fallback
            If blockRefCount = 0 Then
                Trace "No BlockReferences in initial selection (with name filter); reselecting by type only"
                ' Rebuild type-only filter
                BuildTypeFilter IncludeTexts, FilterType, FilterData
                If VainValitut Then
                    Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
                Else
                    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
                End If
                Trace "Selection count after zero-block fallback: " & Joukko.Count
            End If
        End If
        Trace "Selection count: " & Joukko.Count

        ' Fallback: If specific names were requested and the selection is empty, re-select by type only
        ' (kept for safety, though we already select by type when haveNameFilter=True)
        If (Not AllowAll) And haveNameFilter And Joukko.Count = 0 Then
            Trace "Fallback to type-only selection for dynamic blocks"
            ' Rebuild filters: entity types only
            BuildTypeFilter IncludeTexts, FilterType, FilterData
            If VainValitut Then
                Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
            Else
                Joukko.Select acSelectionSetAll, , , FilterType, FilterData
            End If
            Trace "Selection count after fallback: " & Joukko.Count
        End If
    selCount = Joukko.Count
    ' Prepare a bulk buffer sized to the selection (plus a tiny slack) and reasonable column capacity
    rowCap = selCount + 8
    colCap = 40 ' 7 base + typical attributes
    ReDim buf(1 To rowCap, 1 To colCap)
    rowUsed = 0
    maxColUsed = 7

        ' AllowAll already determined above

        StepMsg = "Process entities in selection set"
        Trace StepMsg
        ' Process entities in selection set
        FoundAny = False
        DocStartRow = Rivi
        For i = 0 To Joukko.Count - 1
            ' Resolve entity explicitly via Item to avoid default-member ambiguity in late binding
            StepMsg = "Get entity from selection: index=" & i
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
            Dim entHandle As Variant
            On Error Resume Next
            entHandle = oEnt.Handle
            Err.Clear
            On Error GoTo ErrHandler
            
            StepMsg = "Check entity type"
            Dim entType As String
            Dim isBlock As Boolean, isText As Boolean, isMText As Boolean
            Dim tmp As Variant
            ' Try TypeName first; if it fails (rare), fall back to EntityName or ObjectName via CallByName
            On Error Resume Next
            entType = ""
            entType = TypeName(oEnt) ' e.g., "IAcadBlockReference", "IAcadText", "IAcadMText"
            If Err.Number <> 0 Or entType = "" Then
                Err.Clear
                StepMsg = "Check entity type: EntityName fallback"
                tmp = CallByName(oEnt, "EntityName", VbGet)
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                    Err.Clear
                    StepMsg = "Check entity type: ObjectName fallback"
                    tmp = CallByName(oEnt, "ObjectName", VbGet)
                End If
                If Err.Number <> 0 Or IsEmpty(tmp) Then
                    ' Could not resolve entity type; skip this entity
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo ContinueEntities
                End If
                entType = CStr(tmp)
            End If
            On Error GoTo ErrHandler
            Trace "Entity type: " & entType

            ' Normalize checks for both interface TypeName and AcDb* values
            isBlock = (InStr(1, entType, "BlockReference", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbBlockReference", vbTextCompare) > 0)
            isMText = (InStr(1, entType, "MText", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbMText", vbTextCompare) > 0)
            isText = ((InStr(1, entType, "Text", vbTextCompare) > 0) Or _
                      (InStr(1, entType, "AcDbText", vbTextCompare) > 0)) And Not isMText

            If isBlock Then
                ' Check block name matches filter criteria
                On Error Resume Next
                Set Blokki = oEnt
                If Err.Number <> 0 Or Blokki Is Nothing Then
                    Err.Clear
                    On Error GoTo ErrHandler
                    GoTo ContinueEntities
                End If
                On Error GoTo ErrHandler
                EiPoisteta = AllowAll
                If Not EiPoisteta Then
                    For k = 0 To UBound(Blokit)
                        If UCase(Blokki.EffectiveName) = UCase(Blokit(k)) Then
                            EiPoisteta = True
                            Exit For
                        End If
                    Next k
                End If

                If EiPoisteta Then
                    ' Add a buffered row
                    rowUsed = rowUsed + 1
                    If rowUsed > rowCap Then
                        ' Extend buffer rows if unexpectedly exceeded
                        rowCap = rowCap + 64
                        ReDim Preserve buf(1 To rowCap, 1 To colCap)
                    End If
                    On Error Resume Next
                    buf(rowUsed, 1) = Hakemisto
                    buf(rowUsed, 2) = DWGName
                    buf(rowUsed, 3) = Blokki.EffectiveName
                    buf(rowUsed, 4) = entHandle
                    ' Late binding: InsertionPoint returns a Variant array (x,y,z). Retrieve then index.
                    Dim ip As Variant
                    On Error Resume Next
                    ip = CallByName(Blokki, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ip) Then
                        buf(rowUsed, 5) = CDbl(ip(0)) '' XCord
                        buf(rowUsed, 6) = CDbl(ip(1)) '' YCord
                    Else
                        ' Fallback: attempt property access directly
                        On Error Resume Next
                        buf(rowUsed, 5) = CDbl(Blokki.InsertionPoint(0))
                        buf(rowUsed, 6) = CDbl(Blokki.InsertionPoint(1))
                        Err.Clear
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
                ' Handle text entities only when requested
                If isText Then
                    On Error Resume Next
                    Set oText = oEnt
                    If Err.Number <> 0 Or oText Is Nothing Then
                        Err.Clear
                        On Error GoTo ErrHandler
                        GoTo ContinueEntities
                    End If
                    On Error GoTo ErrHandler
                    Cells(Rivi, 8).Value = oText.TextString
                    Dim ipT As Variant
                    On Error Resume Next
                    ipT = CallByName(oText, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ipT) Then
                        Cells(Rivi, 5).Value = CDbl(ipT(0))
                        Cells(Rivi, 6).Value = CDbl(ipT(1))
                    Else
                        On Error Resume Next
                        Cells(Rivi, 5).Value = CDbl(oText.InsertionPoint(0))
                        Cells(Rivi, 6).Value = CDbl(oText.InsertionPoint(1))
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
                    Cells(Rivi, 8).Value = oMText.TextString
                    Dim ipM As Variant
                    On Error Resume Next
                    ipM = CallByName(oMText, "InsertionPoint", VbGet)
                    Err.Clear
                    On Error GoTo ErrHandler
                    If IsArray(ipM) Then
                        Cells(Rivi, 5).Value = CDbl(ipM(0))
                        Cells(Rivi, 6).Value = CDbl(ipM(1))
                    Else
                        On Error Resume Next
                        Cells(Rivi, 5).Value = CDbl(oMText.InsertionPoint(0))
                        Cells(Rivi, 6).Value = CDbl(oMText.InsertionPoint(1))
                        Err.Clear
                        On Error GoTo ErrHandler
                    End If
                End If
                Range(Cells(Rivi, 1), Cells(Rivi, 8)).Interior.ColorIndex = 8
                Rivi = Rivi + 1
                FoundAny = True
            Else
                ' Skip other entity types
            End If
ContinueEntities:
        Next i

        ' Flush buffered rows to sheet in a single write
        If rowUsed > 0 Then
            Dim outRng As Range
            Set outRng = Range(Cells(DocStartRow, 1), Cells(DocStartRow + rowUsed - 1, maxColUsed))
            outRng.Value = buf
            Rivi = DocStartRow + rowUsed
        End If
        ' After flushing rows, coerce coordinate columns to numbers (handles any text leftovers)
        If rowUsed > 0 Then
            With Range(Cells(DocStartRow, 5), Cells(DocStartRow + rowUsed - 1, 6))
                .NumberFormat = "General"
                .Value = .Value
            End With
        End If
        Trace "Doc processed: " & DWGName & ", rows added: " & rowUsed
        
        ' If nothing matched, inform the user
        If Not FoundAny Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
        End If
        
        ' Close drawing opened from list (do not save) to match original behavior
        If Not Loytyi Then
            oDOC.Close False
        End If
    Next DocRivi
    Trace "TuoDATA finished, total rows added: " & (Rivi - StartBaseRow)
    
Cleanup:
    On Error Resume Next
    oACAD.Visible = True
    oACAD.Preferences.System.SingleDocumentMode = Docmode
    Cells.EntireColumn.AutoFit
    Application.StatusBar = False
    ' Restore Excel settings
    Application.EnableEvents = prevEvents
    If wasCalcAuto Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.ScreenUpdating = prevScreen
    ' Proactively recalc to clear stale-calc indicators without changing user setting
    If Application.Calculation = xlCalculationManual Then
        Application.CalculateFullRebuild
    Else
        Application.CalculateFull
    End If
    
    ' Release objects
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

    Dim i As Long, j As Long, k As Long ' Changed from Integer to Long
    Dim oEntity As Object
    Dim oBlock As Object ' AcadBlockReference - Changed from AcadBlockReference
    Dim BlockArray As Variant
    Dim BlockNimi As String
    Dim DWGName As String
    Dim oText As Object ' AcadText - Changed from AcadText
    Dim oMText As Object ' AcadMText - Changed from AcadMText
    Dim Docmode As Boolean
    Dim StepMsg As String
    Dim TagName As String
    Dim ColIdx As Long
    Dim NewValue As String
    Dim OldValue As String
    Dim UpdateCount As Long
    Dim SkippedCount As Long
    Dim EmptyCount As Long
    
    StepMsg = "VieDATA: Initialization"
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
    
    ' Ensure data sheet is selected
    DATA.Select

    StepMsg = "Connect to AutoCAD"
    Trace StepMsg
    
    ' Connect to running AutoCAD instance
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
            If Not OliAuki Then
                If Not oDOC Is Nothing Then
                    oDOC.SaveAs oDOC.FullName, Ver
                    oDOC.Close False
                End If
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
                        Trace "Block has " & (UBound(BlockArray) + 1) & " attributes"
                        
                        ' TAG-BASED UPDATE LOGIC (symmetrical with TuoDATA)
                        ' Loop through each attribute in the block
                        For j = 0 To UBound(BlockArray)
                            On Error Resume Next
                            TagName = UCase(BlockArray(j).TagString)
                            OldValue = BlockArray(j).TextString
                            If Err.Number <> 0 Then
                                Trace "ERROR: Cannot read attribute " & j & ": " & Err.Description
                                Err.Clear
                                GoTo NextAttribute
                            End If
                            On Error GoTo ErrHandler
                            
                            ' Find matching column in Excel header (row 1)
                            ColIdx = 0
                            For k = 8 To 256 ' Start from column H (first attribute column)
                                If UCase(Cells(1, k).Value) = TagName Then
                                    ColIdx = k
                                    Exit For
                                End If
                                ' Stop if we hit empty headers
                                If Cells(1, k).Value = "" Then Exit For
                            Next k
                            
                            If ColIdx > 0 Then
                                ' Column found - check if Excel value is non-empty
                                NewValue = CStr(Cells(i, ColIdx).Text)
                                
                                If Len(NewValue) > 0 Then
                                    ' Update attribute only if Excel has a value
                                    BlockArray(j).TextString = NewValue
                                    UpdateCount = UpdateCount + 1
                                    Trace "  [" & TagName & "] '" & OldValue & "' -> '" & NewValue & "'"
                                Else
                                    ' Excel cell is empty - preserve existing AutoCAD value
                                    EmptyCount = EmptyCount + 1
                                    Trace "  [" & TagName & "] SKIPPED (Excel empty, preserving '" & OldValue & "')"
                                End If
                            Else
                                ' No matching column found in Excel
                                SkippedCount = SkippedCount + 1
                                Trace "  [" & TagName & "] SKIPPED (no Excel column)"
                            End If
NextAttribute:
                        Next j
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
    
    ' Export summary
    Trace "VieDATA completed: Updated=" & UpdateCount & ", Skipped(no column)=" & SkippedCount & ", Preserved(empty)=" & EmptyCount
  
Cleanup:
    On Error Resume Next
    Aloitus.Activate
    If Not oACAD Is Nothing Then
        oACAD.Preferences.System.SingleDocumentMode = Docmode
    End If
    Application.StatusBar = False
    
    ' Release objects
    Set oEntity = Nothing
    Set oBlock = Nothing
    Set BlockArray = Nothing
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
' Delete selected blocks from AutoCAD drawing
' 26.10.2025 - 64-bit compatibility, added error handling

    Dim i As Long, j As Long ' Changed from Integer to Long
    Dim Docmode As Boolean
    Dim oEntity As Object
    Dim DWGName As String
    Dim Rivi As Range
    Dim RiviNo As Long
    Dim Kaydyt As String

    On Error GoTo ErrHandler
    
    ' Ensure data sheet is selected
    DATA.Select
    
    ' Connect to running AutoCAD instance
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
        If InStr(Kaydyt, "|" & Rivi.Row & "|") = 0 Then
            RiviNo = Rivi.Row
            Kaydyt = Kaydyt & "|" & RiviNo & "|"
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
    
    ' Release objects
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

Private Function OtsS(Nimi As String) As Long '' Changed from Integer to Long
'' Find or create column for attribute name
    Dim i As Long '' Changed from Integer to Long
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
'' Open AutoCAD document if needed
'' 26.10.2025 - Changed Integer to Long

    Dim Doku As String
    Dim EdDoku As String
    Dim Tiedosto As String
    Dim i As Long '' Changed from Integer to Long
    
    Doku = (Cells(Rivi, 2).Value) & ".dwg"
    EdDoku = (Cells(Rivi - 1, 2).Value) & ".dwg"
    Tiedosto = Cells(Rivi, 1).Value & "\" & Doku
    
    '' Check if desired document is already active
    If Not oDOC Is Nothing Then '' Some drawing is already open
        If LCase(oDOC.Name) = LCase(Doku) Then '' Drawing is the one being processed
            AvaaDoc = True
            Exit Function
        ElseIf LCase(oDOC.Name) = LCase(EdDoku) Then '' Previous drawing is open
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
    
    '' Desired drawing was not already being processed
    '' Check if desired drawing is open in AutoCAD
    OliAuki = False
    For i = 0 To oACAD.Documents.Count - 1
        If LCase(oACAD.Documents(i).Name) = LCase(Doku) Then '' Drawing is open, set it as active
            OliAuki = True
            oACAD.Documents(i).Activate
            Set oDOC = oACAD.ActiveDocument
            AvaaDoc = True
            Exit Function
        End If
    Next i
    
    '' Desired drawing was not being processed and not open in AutoCAD, so open it
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
'' Numbering tool for blocks
'' 26.10.2025 - Changed Integer to Long

    Dim Alku As String
    Dim Jakso As Long '' Changed from Integer to Long
    Dim Vali As Long '' Changed from Integer to Long
    Dim i As Long, j As Long '' Changed from Integer to Long

    Aloitus.Tyhjenna.Value = True
    Aloitus.Nykyinen.Value = True
    
    '' Fetch from drawing
    TuoDATA True
    Alku = Aloitus.Range("D13").Value
    Vali = 2
    
    Cells.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("F2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
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
    
    Aloitus.Range("D13").Value = LNumero(j, Alku)
    VieDATA
End Sub

Private Function LNumero(No As Long, Alku As String) As String
'' Format number with leading zeros
'' 26.10.2025 - Changed from Integer to Long

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
'' Reference numbering tool
'' 26.10.2025 - Changed Integer to Long

    Dim vSivu As Long '' Changed from Integer to Long
    Dim Kirjain As String
    Dim i As Long, j As Long '' Changed from Integer to Long

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
'' Add value to number string maintaining format
'' 26.10.2025 - Changed Integer to Long

    Dim Pit As Long '' Changed from Integer to Long
    Dim i As Long '' Changed from Integer to Long
    
    Pit = Len(Nro)
    Nro = CStr(Val(Nro) + Maara)
    
    For i = 1 To Pit - Len(Nro)
        Nro = "0" & Nro
    Next i
    
    Lisaa = Nro
End Function

Function Yhd(Alue As Range, Optional Merkki As String) As String
'' Concatenate range values with separator
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

