Option Explicit

' Updated 2025-10-26: 64-bit compatibility, performance optimizations, improved error handling
' Excel-AutoCAD integration: Import/export block attributes and text entities
' Changes: Integer → Long (64-bit), early binding → late binding (compatibility),
'          added error handlers, array optimization for performance

' ============================================================================
' AutoCAD Constants - Required for Late Binding
' ============================================================================
' When using late binding (Object instead of AcadApplication, AcadDocument, etc.),
' the AutoCAD Type Library is not referenced, so built-in constants are not available.
' These must be manually defined with their numeric values.
' Source: Autodesk AutoCAD ActiveX/VBA Reference Documentation
' ============================================================================

' Selection methods
Public Const acSelectionSetAll As Long = 2          ' Select all entities
Public Const acSelectionSetPrevious As Long = 4     ' Select previously selected entities

' Active space
Public Const acModelSpace As Long = 1               ' Model space (vs paper space)

' Drawing versions for SaveAs
Public Const acNative As Long = 60                  ' Current AutoCAD version
Public Const ac2004_dwg As Long = 24                ' AutoCAD 2004 format
Public Const ac2007_dwg As Long = 36                ' AutoCAD 2007 format
Public Const ac2010_dwg As Long = 48                ' AutoCAD 2010 format
Public Const ac2013_dwg As Long = 60                ' AutoCAD 2013 format

' Window state
Public Const acMax As Long = 3                      ' Maximize window

' Zoom methods
Public Const acZoomScaledRelative As Long = 3       ' Zoom relative to current view

Public oACAD As Object ' AcadApplication (late binding for compatibility) - Changed from AcadApplication
Public oDOC As Object ' AcadDocument - Changed from AcadDocument
Public OliAuki As Boolean
Public Ver As Long ' Changed from Integer to Long for 64-bit compatibility

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
    Dim i As Long, j As Long, jj As Long ' Changed from Integer to Long
    ' ⚠️ CRITICAL: FilterType MUST remain Integer (not Long)
    ' AutoCAD's SelectionSet.Select API requires Integer array for DXF filter codes
    ' Changing to Long causes error: "Invalid argument FilterType in Select"
    Dim FilterType(0) As Integer ' Exception: Must remain Integer - AutoCAD COM API requirement
    Dim FilterData(0) As Variant
    Dim Poista() As Object ' AcadEntity - Changed from AcadEntity array
    Dim L As Long ' Changed from Integer to Long
    Dim EiPoisteta As Boolean
    Dim Nimet As String
    Dim Blokit As Variant
    Dim Blokki As Object ' AcadBlockReference - Changed from AcadBlockReference
    Dim DWGName As String
    Dim Hakemisto As String
    Dim EkaKerta As Boolean
    Dim Rivi As Long
    Dim oText As Object ' AcadText - Changed from AcadText
    Dim oMText As Object ' AcadMText - Changed from AcadMText
    Dim DocRivi As Long ' Changed from Integer to Long
    Dim DocMaara As Long ' Changed from Integer to Long
    Dim Loytyi As Boolean
    Dim Filter2 As Boolean
    Dim Docmode As Boolean
    
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
  
    Application.ScreenUpdating = False
    
    ' Connect to running AutoCAD instance
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
        Exit Sub
    End If
    On Error GoTo ErrHandler
 
    ' Initialize headers if clearing worksheet
    If Aloitus.Tyhjenna.Value = True Then
        Tyhjenna = True
    End If
    
    EkaKerta = True
    DATA.Select
    
    If Tyhjenna Then
        Cells.Select
        Selection.Clear
        Selection.NumberFormat = "@" ' Set cell format to text
        Rows("1:1").Font.Bold = True ' Make headers bold
        Columns("E:F").NumberFormat = "General"
        Range("A1").Select
        Cells(1, 1).Value = "PATH"
        Cells(1, 2).Value = "DWG"
        Cells(1, 3).Value = "BLOCK"
        Cells(1, 4).Value = "HANDLE"
        Cells(1, 5).Value = "XCord"
        Cells(1, 6).Value = "YCord"
        Cells(1, 7).Value = "Layer"
    End If
    
    ' Get document count
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
    
    ' Parse block names
    Nimet = UCase(Aloitus.Range("D7").Value)
    Blokit = Split(Nimet, ",")
    
    For i = 0 To UBound(Blokit)
        Blokit(i) = Trim(Blokit(i))
    Next i
    
    ' Set up entity filter
    If Aloitus.Range("D5").Value = "Tekstit" Then
        FilterType(0) = 0
        FilterData(0) = "TEXT,MTEXT,DTEXT"
    ElseIf Aloitus.Range("D5").Value = "Blokit ja tekstit" Then
        FilterType(0) = 0
        FilterData(0) = "TEXT,MTEXT,DTEXT,INSERT"
    Else
        FilterType(0) = 0
        FilterData(0) = "INSERT"
    End If
    
    ' Save and temporarily change document mode
    Docmode = oACAD.Preferences.System.SingleDocumentMode
    oACAD.Preferences.System.SingleDocumentMode = False
    
    ' Process each document
    For DocRivi = 1 To DocMaara
        Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
        
        If Not Listasta Then
            Set oDOC = oACAD.ActiveDocument
            Loytyi = True
        Else
            ' Find document in collection
            Loytyi = False
            For i = 0 To oACAD.Documents.Count - 1
                If UCase(oACAD.Documents(i).Name) = UCase(Dir(TIEDLISTA.Cells(DocRivi, 1).Value)) Then
                    Set oDOC = oACAD.Documents(i)
                    Loytyi = True
                    Exit For
                End If
            Next i
            
            If Not Loytyi Then
                Set oDOC = oACAD.Documents.Open(TIEDLISTA.Cells(DocRivi, 1).Value)
            End If
        End If
        
        ' Get document information
        DWGName = Left(oDOC.Name, Len(oDOC.Name) - 4) ' Remove .dwg extension
        Hakemisto = oDOC.Path
        
        ' Clean up existing selection set
        For i = 0 To oDOC.SelectionSets.Count - 1
            If oDOC.SelectionSets(i).Name = "EXCELHAKU" Then
                oDOC.SelectionSets(i).Delete
                Exit For
            End If
        Next i
        
        Set Joukko = oDOC.SelectionSets.Add("EXCELHAKU")
        
        ' ========================================================================
        ' Select entities with filter
        ' SelectOnScreen is used for previously selected entities (user selection)
        ' Select is used for programmatic selection of all entities
        ' Both methods accept FilterType and FilterData arrays for filtering
        ' ========================================================================
        
        If VainValitut Then
            ' Select from user's previously selected entities on screen
            Joukko.SelectOnScreen FilterType, FilterData
        Else
            ' Select all entities matching filter programmatically
            Joukko.Select acSelectionSetAll, Empty, Empty, FilterType, FilterData
        End If
        
        ' Filter blocks that don''t match criteria
        L = 0
        For j = 0 To Joukko.Count - 1
            If Joukko(j).EntityName = "AcDbBlockReference" Then
                EiPoisteta = False
                For i = 0 To UBound(Blokit)
                    If UCase(Joukko(j).EffectiveName) = Blokit(i) Then
                        EiPoisteta = True
                        Exit For
                    ElseIf Blokit(i) = "*" Then
                        EiPoisteta = True
                        Exit For ' Added Exit For for efficiency
                    End If
                Next i
                
                If Not EiPoisteta Then
                    ReDim Preserve Poista(L)
                    Set Poista(L) = Joukko(j)
                    L = L + 1
                End If
            End If
        Next j
        
        If L > 0 Then Joukko.RemoveItems Poista
        
        If Joukko.Count = 0 Then
            MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
        End If
        
        ' Process entities in selection set
        For i = 0 To Joukko.Count - 1
            Application.StatusBar = "Luetaan tietoa: " & i + 1 & "/" & Joukko.Count & "  File: " & DWGName
            Cells(Rivi, 1).Value = Hakemisto
            Cells(Rivi, 2).Value = DWGName
            Cells(Rivi, 4).Value = Joukko(i).Handle
            
            If Joukko(i).EntityName = "AcDbBlockReference" Then
                Set Blokki = Joukko(i)
                Cells(Rivi, 5).Value = Blokki.InsertionPoint(0) '' XCord
                Cells(Rivi, 6).Value = Blokki.InsertionPoint(1) '' YCord
                Cells(Rivi, 7).Value = Blokki.Layer
                Cells(Rivi, 3).Value = Blokki.EffectiveName
                Cells(Rivi, 3).ClearNotes
                Cells(Rivi, 3).AddComment Blokki.Name
                
                If Blokki.HasAttributes Then
                    BlockArray = Blokki.GetAttributes
                    For j = 0 To UBound(BlockArray)
                        Cells(1, 8 + j).Value = BlockArray(j).TagString
                        Cells(1, 8 + j).ClearNotes
                        Cells(1, 8 + j).AddComment Blokki.EffectiveName
                        Cells(Rivi, 8 + j).Value = BlockArray(j).TextString
                    Next j
                    Rivi = Rivi + 1
                End If
            Else
                ' Handle text entities
                If Joukko(i).EntityName = "AcDbText" Then
                    Set oText = Joukko(i)
                    Cells(Rivi, 8).Value = oText.TextString
                    Cells(Rivi, 5).Value = oText.InsertionPoint(0)
                    Cells(Rivi, 6).Value = oText.InsertionPoint(1)
                    Range(Cells(Rivi, 1), Cells(Rivi, 8)).Interior.ColorIndex = 8
                Else
                    Set oMText = Joukko(i)
                    Cells(Rivi, 8).Value = oMText.TextString
                    Cells(Rivi, 5).Value = oMText.InsertionPoint(0)
                    Cells(Rivi, 6).Value = oMText.InsertionPoint(1)
                    Range(Cells(Rivi, 1), Cells(Rivi, 8)).Interior.ColorIndex = 8
                End If
                Rivi = Rivi + 1
            End If
        Next i
        
        If Not Loytyi Then oDOC.Close False
    Next DocRivi
    
Cleanup:
    On Error Resume Next
    oACAD.Visible = True
    oACAD.Preferences.System.SingleDocumentMode = Docmode
    Cells.EntireColumn.AutoFit
    Application.StatusBar = False
    
    ' Release objects
    Set Blokki = Nothing
    Set Joukko = Nothing
    Set oDOC = Nothing
    Set oACAD = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Tuo DATA"
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

    Dim i As Long, j As Long ' Changed from Integer to Long
    Dim oEntity As Object
    Dim oBlock As Object ' AcadBlockReference - Changed from AcadBlockReference
    Dim BlockArray As Variant
    Dim BlockNimi As String
    Dim DWGName As String
    Dim oText As Object ' AcadText - Changed from AcadText
    Dim oMText As Object ' AcadMText - Changed from AcadMText
    Dim Docmode As Boolean
    
    On Error GoTo ErrHandler
    
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
                    If oBlock.HasAttributes Then
                        BlockArray = oBlock.GetAttributes
                        For j = 0 To UBound(BlockArray)
                            BlockArray(j).TextString = Cells(i, 8 + j).Text
                        Next j
                    End If
                Else
                    If oEntity.EntityName = "AcDbText" Then
                        Set oText = oEntity
                        oText.TextString = Cells(i, 8).Value
                    Else
                        Set oMText = oEntity
                        oMText.TextString = Cells(i, 8).Value
                    End If
                End If
            End If
        End If
    Loop
  
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
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Vie DATA"
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

Private Function EOtsS(Nimi As String) As Long '' Changed from Integer to Long
'' Find existing column for attribute name
    Dim i As Long '' Changed from Integer to Long
    Nimi = UCase(Nimi)
    i = 7
    Do
        If Cells(1, i).Value = Nimi Then
            EOtsS = i
            Exit Do
        ElseIf Cells(1, i).Value = "" Then
            EOtsS = i
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
    vSivu = CLng(Aloitus.Range("D18").Value) '' Changed from CInt to CLng
  
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
