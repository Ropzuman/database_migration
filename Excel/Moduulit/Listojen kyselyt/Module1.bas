Public CheckOK As Boolean
Public PHStart As Long
Public PHEnd As Long
Public PFStart As Long
Public PFEnd As Long
Public DocStart As Long
Public DocEnd As Long
Public Sarakkeita As Long
Public RMAX As Long
Public POSheet As String
Public HideLINKING As Boolean
Public AddFooter As Boolean
Public DIContract As String
Public DIMill As String
Public DIDepartName As String
Public DICustomer As String
Public DIProject As String
Public DIProjNo As String
Public DIProjName As String
Public DIMunit As String
Public DIManager As String
Public DIDocNo As String
Public DIMetsoDocNo As String
Public DIDocName As String
Public DIDocName1 As String
Public DIDocName2 As String
Public DIDocName3 As String
Public DIPath As String
Public DIFile As String
Public DIDate As String
Public DIRev As String
Public DIRevArr() As String
Public DIRevID As String
Public DIRevDate As String
Public DIStatus As String

'''
' Module1.vba - Main logic for Kytkentälista Excel macro system.
' Handles data fetching from Access, checkouts and printout generation.
'''

' Performance/UX state (to minimize screen flashing and speed up macros.)
Private prevScreenUpdating As Boolean
Private prevCalculation As XlCalculation
Private prevEnableEvents As Boolean
Private prevDisplayAlerts As Boolean
Private prevDisplayStatusBar As Boolean

'''
' BeginFastMode: Temporarily disables Excel UI updates, events, and sets calculation to manual
' to speed up macro execution and prevent screen flicker.
'''
Private Sub BeginFastMode()
  prevScreenUpdating = Application.ScreenUpdating
  prevCalculation = Application.Calculation
  prevEnableEvents = Application.EnableEvents
  prevDisplayAlerts = Application.DisplayAlerts
  prevDisplayStatusBar = Application.DisplayStatusBar
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  On Error Resume Next
  Application.DisplayStatusBar = False
  Application.AskToUpdateLinks = False
  On Error GoTo 0
End Sub

'''
' EndFastMode: Restores Excel UI and calculation settings to their previous state.
'''
Private Sub EndFastMode()
  On Error Resume Next
  Application.ScreenUpdating = prevScreenUpdating
  Application.Calculation = prevCalculation
  Application.EnableEvents = prevEnableEvents
  Application.DisplayAlerts = prevDisplayAlerts
  Application.DisplayStatusBar = prevDisplayStatusBar
  Application.AskToUpdateLinks = True
  On Error GoTo 0
End Sub

'''
' HaeData: Fetches data from Access database using ODBC and SQL queries defined in the faceplate.
' Populates DB1 (main body data) and DB2 (document metadata) sheets.
' QueryTable lifecycle: Create → Refresh → Delete (no persistent connections left behind).
' Diagnostics: Row counts displayed in StatusBar and Immediate Window after each query.
'''
Sub HaeData()
Dim Kanta As String
Dim sSQL(1 To 2) As String
Dim Valinta As Long
Dim i As Long
Dim TAULUKKO As QueryTable
Dim Yhteys As String
Dim ws As Worksheet
Dim rc As Long

  ' Select which SQL query to use from Main sheet
  On Error Resume Next
  If Sheets("Main").Valinta1.Value = True Then
    Valinta = 0
  ElseIf Sheets("Main").Valinta2.Value = True Then
    Valinta = 1
  Else
    Valinta = 2
  End If
  On Error GoTo 0
  
  Kanta = Sheets("Main").Range("C6").Value
  sSQL(1) = Sheets("Main").Cells(8 + Valinta, 3).Value
  sSQL(2) = Sheets("Main").Cells(12 + Valinta, 3).Value
  
  BeginFastMode
  
' Verify database file exists
  If Dir(Kanta) = "" Then
    MsgBox "Database file not found: " & Kanta, vbCritical, "Database Error"
    EndFastMode
    Exit Sub
  End If
  
  Yhteys = "ODBC;DBQ=" & Kanta & ";Driver={Microsoft Access Driver (*.mdb, *.accdb)}"
  
  On Error GoTo ErrorHandler
  
  For i = 1 To 2
    Set ws = ThisWorkbook.Sheets("DB" & i)
    ws.Cells.Clear
    
    If sSQL(i) <> "" Then
      ' Create QueryTable, refresh, then delete (no persistent connections)
      Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
      With TAULUKKO
        .Sql = sSQL(i)
        .FieldNames = True
        .RefreshStyle = xlInsertDeleteCells
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .HasAutoFormat = True
        .SaveData = True
        .BackgroundQuery = False
        .Refresh
        .Delete
      End With
      Set TAULUKKO = Nothing
      
      ' Report row count for diagnostics
      rc = 0
      On Error Resume Next
      rc = ws.UsedRange.Rows.Count
      On Error GoTo 0
      Application.StatusBar = "DB" & i & " rows: " & rc
      Debug.Print "HaeData: DB" & i & " rows=", rc
    End If
  Next i
  EndFastMode
  MsgBox "Data brought successfully!", vbOKOnly, "Ready"
  Sheets("Main").Select
  Exit Sub
  
ErrorHandler:
  EndFastMode
  MsgBox "ODBC Error: " & Err.Description & vbCrLf & vbCrLf & _
         "Database: " & Kanta & vbCrLf & _
         "SQL Query " & i & ": " & sSQL(i), vbCritical, "Database Connection Error"
  Err.Clear
  Sheets("Main").Select
End Sub
'''
' GenPrintout: Generates a new printout workbook using TEMPLATE and data from DB1.
' Uses template-driven population: copies TEMPLATE blocks per data group (RMAX rows),
' then maps values from LINKING sheet via comment-based markers (VaihdaLinkit).
' This preserves the template's layout, formatting, and linking logic.
'''
Sub GenPrintout()
  If CheckOK = False Then
    MsgBox "Check data first!", vbCritical, "Error!"
    Exit Sub
  End If
  
  ' Variable declarations
  Dim srcWB As Workbook
  Dim destWB As Workbook
  Dim destSheet As Worksheet
  Dim wsDB1 As Worksheet
  Dim ViimRivi As Long
  Dim Recordeja As Long
  Dim Kerta As Long
  Dim Tiedosto As String
  Dim i As Long
  Dim Oletus As String
  Dim lastCol As Long
  Dim c As Range
  Dim Riveja As Long
  
  On Error GoTo GenPrintoutError
  BeginFastMode
  
  ' Get POSheet name from faceplate
  POSheet = Sheets("Main").Range("C16").Value
  If Trim(POSheet) = "" Then POSheet = "Printout" ' Default name if not set
  
  ' Ensure document info is current (path/name from DB2)
  On Error Resume Next
  If Trim(DIPath) = "" Or Trim(DIFile) = "" Then HaeDocTiedot
  On Error GoTo GenPrintoutError
  
  Application.StatusBar = "Initializing printout generation..."
    
  ' Get user options from faceplate
  AddFooter = Sheets("Main").AddFooter.Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo GenPrintoutError
    
  ' Set workbook references
  Set srcWB = ThisWorkbook
  Set wsDB1 = srcWB.Sheets("DB1")
  
  Application.StatusBar = "Reading data from DB1..."
  
  ' Find last used row in DB1 - with null checking
  Dim lastCell As Range
  Set lastCell = wsDB1.Cells.Find(What:="*", _
                                  After:=wsDB1.Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
  
  If lastCell Is Nothing Then
    EndFastMode
    MsgBox "DB1 sheet is empty! Please click 'Get Data' first to load data from database.", vbCritical, "No Data Error"
    Exit Sub
  End If
  
  Recordeja = lastCell.Row
  
  Application.StatusBar = "Creating new workbook..."
  
  ' Create new workbook by copying Info sheet
  srcWB.Sheets("Info").Copy
  Set destWB = ActiveWorkbook
  destWB.Sheets(1).Cells.ClearComments
  
  ' Copy TEMPLATE, Legend, and Revisions to new workbook
  srcWB.Sheets("TEMPLATE").Copy After:=destWB.Sheets(1)
  destWB.Sheets(2).Name = POSheet
  Set destSheet = destWB.Sheets(POSheet)
  
  srcWB.Sheets("Legend").Copy After:=destWB.Sheets(2)
  srcWB.Sheets("Revisions").Copy After:=destWB.Sheets(1)
  
  Application.StatusBar = "Populating revisions..."
  PopulateRevisionsSimple
  
  ' Clear POSheet and remove shapes
  Application.StatusBar = "Preparing printout sheet..."
  destSheet.Cells.Clear
  If destSheet.Shapes.Count > 0 Then
    For i = destSheet.Shapes.Count To 1 Step -1
      destSheet.Shapes(i).Delete
    Next i
  End If
  
  ' Copy header rows from TEMPLATE to POSheet
  Application.StatusBar = "Copying headers..."
  ViimRivi = 1
  srcWB.Sheets("TEMPLATE").Rows(PHStart & ":" & PHEnd).Copy _
      Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PHEnd - PHStart)
  Application.CutCopyMode = False
  
  ' Set print title rows and freeze panes
  destSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  destSheet.Activate
  destSheet.Cells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
  
  ' Set up footers for first three sheets (Info, POSheet, Legend)
  Application.StatusBar = "Setting up footers..."
  For i = 1 To 3
    With destWB.Sheets(i).PageSetup
      ' Null-safe string concatenation (empty values are OK - some equipment lacks certain attributes)
      .LeftFooter = "&8Document: " & (DIMetsoDocNo & "") & Chr(10) _
                  & "&8Revision: " & (DIRevID & "") & " - " & (DIRevDate & "") & Chr(10) _
                  & "&8Status: " & (DIStatus & "")
      .CenterFooter = "&8 " & (DICustomer & "") & Chr(10) _
                    & "&8 " & (DIMill & "") & Chr(10) _
                    & "&8 " & (DIDepartName & "") & Chr(10) _
                    & "&8 " & (DIDocName2 & "")
      .RightFooter = "&8Project: " & (DIProjNo & "") & Chr(10) _
                   & "&8File: &F" & Chr(10) _
                   & "&8Page &P(&N)"
    End With
  Next i
  
  ' Create LINKING sheet and copy DB1 data
  Application.StatusBar = "Creating LINKING sheet..."
  With destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
    .Name = "LINKING"
    wsDB1.Cells.Copy Destination:=.Range("A1")
  End With
  Application.CutCopyMode = False
  
  ' Initial linking for header area
  Kerta = 0
  VaihdaLinkit destSheet, 1, ViimRivi, Kerta
  
  ' Template-driven population: copy TEMPLATE blocks and map values via VaihdaLinkit
  Application.StatusBar = "Copying data to printout using template blocks..."
  Riveja = DocEnd - DocStart
  If RMAX <= 0 Then RMAX = 1
  ' Iterate DB1 data rows in groups of RMAX, copying TEMPLATE rows each time
  Kerta = 0
  For i = 2 To Recordeja Step RMAX
    ' Copy TEMPLATE block to destination
    srcWB.Sheets("TEMPLATE").Rows(DocStart & ":" & DocEnd).Copy _
        Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + Riveja)
    Application.CutCopyMode = False ' Clear clipboard after copy
    ' Apply alternating shading per block
    If ((i - 2) \ RMAX) Mod 2 = 1 Then
      With destSheet.Range(destSheet.Cells(ViimRivi, 1), destSheet.Cells(ViimRivi + Riveja, Sarakkeita)).Interior
        .ColorIndex = 19
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
      End With
    End If
    ' Map values from LINKING to the template area via comment markers
    VaihdaLinkit destSheet, ViimRivi, ViimRivi + Riveja, Kerta
    ' Advance to next block
    ViimRivi = ViimRivi + Riveja + 1
    Kerta = Kerta + 1
  Next i
  
  ' Delete extra columns beyond Sarakkeita
  lastCol = destSheet.Cells(1, destSheet.Columns.Count).End(xlToLeft).Column
  If lastCol > Sarakkeita Then
    destSheet.Columns(Sarakkeita + 1 & ":" & lastCol).Delete
  End If
  
  ' Add footer block if requested
  If AddFooter = True Then
    Application.StatusBar = "Adding footer..."
    srcWB.Sheets("TEMPLATE").Rows(PFStart & ":" & PFEnd).Copy _
        Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PFEnd - PFStart)
    Application.CutCopyMode = False
    
    ' Set up formulas for footer summary cells
    For Each c In destSheet.Range(destSheet.Cells(ViimRivi, 1), destSheet.Cells(ViimRivi + PFEnd - PFStart, Sarakkeita))
      If Len(CStr(c.Value)) >= 2 Then
        If Left(CStr(c.Value), 2) = "&&" Then
          c.Formula = "=SUM(" & destSheet.Cells(PHEnd, c.Column).Address(False, False) & ":" & destSheet.Cells(ViimRivi - 1, c.Column).Address(False, False) & ")"
        End If
      End If
    Next c
  End If
  
  ' Clear comments and add LINKING comments
  Application.StatusBar = "Finalizing..."
  destSheet.Cells.ClearComments
  TeeLinkingKommentit
  
  ' Handle LINKING sheet visibility/deletion
  On Error Resume Next
  If HideLINKING Then
    destWB.Sheets("LINKING").Visible = False
  Else
    Application.DisplayAlerts = False
    destWB.Sheets("LINKING").Delete
    Application.DisplayAlerts = True
  End If
  On Error GoTo GenPrintoutError
  
  destSheet.Activate
  Application.StatusBar = False
  EndFastMode
  
  ' Prompt for file name and save
  ' Build a robust default path+file for the save dialog
  Dim defPath As String, defName As String
  
  ' Safe handling of potentially empty/null DIPath and DIFile
  On Error Resume Next
    defPath = Trim(DIPath & "")
    defName = Trim(DIFile & "")
  On Error GoTo GenPrintoutError
  
  If defPath = "" Then defPath = ThisWorkbook.Path & Application.PathSeparator
  If Right$(defPath, 1) <> "\\" And Right$(defPath, 1) <> "/" Then defPath = defPath & Application.PathSeparator
  
  ' Per requirement: file name should come from DB2 "File" column
  If defName = "" Then defName = POSheet ' fallback only if DB2 didn't provide a name
  If defName = "" Then defName = "Printout" ' Ultimate fallback
  
  ' Ensure .xlsx extension if none provided
  If InStrRev(defName, ".") = 0 Then defName = defName & ".xlsx"
  
  Oletus = defPath & defName
  Tiedosto = InputBox("Give The File Name", "Save File", Oletus)
  If Tiedosto <> "" Then
    destWB.BuiltinDocumentProperties("Title").Value = POSheet
    destWB.SaveAs Tiedosto, xlOpenXMLWorkbook
  End If
  Exit Sub

GenPrintoutError:
  Application.StatusBar = False
  EndFastMode
  
  ' Enhanced error handler with context-specific messages
  Dim errMsg As String
  errMsg = "Error in GenPrintout: " & Err.Description & " (Error " & Err.Number & ")"
  
  ' Add context based on error number
  Select Case Err.Number
    Case 91
      errMsg = errMsg & vbCrLf & vbCrLf & "Object variable not set." & vbCrLf & _
               "This usually means a sheet or range is missing." & vbCrLf & _
               "Check that all required sheets exist (TEMPLATE, DB1, DB2, Info, Legend, Revisions)."
    Case 1004
      errMsg = errMsg & vbCrLf & vbCrLf & "Application-defined or object-defined error." & vbCrLf & _
               "This often means a copy/paste operation failed or a sheet name is invalid."
    Case 9
      errMsg = errMsg & vbCrLf & vbCrLf & "Subscript out of range." & vbCrLf & _
               "A sheet with the specified name doesn't exist."
    Case 13
      errMsg = errMsg & vbCrLf & vbCrLf & "Type mismatch." & vbCrLf & _
               "Trying to use incompatible data types (e.g., text where number expected)."
  End Select
  
  MsgBox errMsg, vbCritical, "Printout Generation Error"
  
  ' Log to Immediate Window for debugging
  Debug.Print "GenPrintout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  Debug.Print "  Source: " & Err.Source
  
  Err.Clear
  On Error GoTo 0
End Sub
'''
' Checkout: Validates TEMPLATE structure and DB1 headers.
' - Finds area markers (&&PAGE_HEADER_START, &&DOC_DATA_START, etc.) in TEMPLATE
' - Validates row markers (££ for single-row, £1/2/3 for multi-row groups)
' - Creates comments in TEMPLATE cells to link to DB1 columns (via EtsiOts)
' - Populates Info sheet with document metadata from DB2
' - Sets CheckOK flag if validation passes, otherwise reports errors to ERRORS sheet
'''
Sub Checkout()
Dim i As Long
Dim j As Long
Dim Arvo As String
Dim Virhe As Boolean
Dim wsTemplate As Worksheet
Dim wsErrors As Worksheet

  On Error GoTo CheckoutError
  
  CheckOK = False
  RMAX = 0
  Virhe = False
  Application.ScreenUpdating = False
  
  Set wsErrors = Sheets("ERRORS")
  Set wsTemplate = Sheets("TEMPLATE")
  
  ' Clear ERRORS sheet
  wsErrors.Cells.Clear
  
  ' Get constants from faceplate
  POSheet = Sheets("Main").Range("C16").Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo CheckoutError
  
  ' Clear comments from TEMPLATE
  wsTemplate.Cells.ClearComments
  
  ' Find area markers in TEMPLATE - with null checking
  Dim foundCell As Range
  With wsTemplate
    Set foundCell = .Cells.Find(What:="&&PAGE_HEADER_START")
    If foundCell Is Nothing Then
      wsErrors.Range("A1").Value = "TEMPLATE ERROR: Marker &&PAGE_HEADER_START not found!"
      wsErrors.Range("A1").Font.Bold = True
      wsErrors.Range("A1").Font.ColorIndex = 3
      wsErrors.Activate
      Application.ScreenUpdating = True
      MsgBox "TEMPLATE is missing required markers! See ERRORS sheet.", vbCritical, "Template Error"
      Exit Sub
    End If
    PHStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&PAGE_HEADER_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_HEADER_END not found"
    PHEnd = foundCell.Row - 1
    
    Set foundCell = .Cells.Find(What:="&&DOC_DATA_START")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOC_DATA_START not found"
    DocStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&DOC_DATA_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOC_DATA_END not found"
    DocEnd = foundCell.Row - 1
    
    Set foundCell = .Cells.Find(What:="&&END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&END not found"
    Sarakkeita = foundCell.Column
    
    Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_START")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_FOOTER_START not found"
    PFStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_FOOTER_END not found"
    PFEnd = foundCell.Row - 1
  End With
  
  ' Fetch document info from DB2 sheet
  HaeDocTiedot
  
  ' Check if data was loaded from DB2
  If DIProject = "" And DIDocNo = "" And DIProjNo = "" And DIMetsoDocNo = "" Then
    wsErrors.Range("A1").Value = "WARNING: No document metadata found in DB2 sheet!"
    wsErrors.Range("A2").Value = "Please click 'Get Data' button first to load data from database."
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A1").Font.ColorIndex = 3 ' Red
    Debug.Print "Checkout: No data found in DB2 - Info sheet will be empty"
  End If
  
  VaihdaInfo   'Populate document info to Info sheet only (not Revisions during checkout)
  
  ' Search for row markers in TEMPLATE
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      Arvo = wsTemplate.Cells(i, j).Value
      If Len(Arvo) > 2 Then 'Cell has data
        If Left(Arvo, 2) = "££" Then
          If RMAX > 1 Then Virhe = True
          RMAX = 1
        ElseIf Left(Arvo, 1) = "£" Then
          If RMAX <> 0 And RMAX <> CInt(Mid(Arvo, 4, 1)) Then Virhe = True
          RMAX = CInt(Mid(Arvo, 4, 1))
        End If
      End If
    Next j
  Next i
  
  If Virhe Then
    wsErrors.Cells.Clear
    wsErrors.Range("A1").Value = "Error on declaration!"
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A2").Value = "- You cannot have ££ and £1/2 links on same template."
    wsErrors.Range("A3").Value = "- Neither you can have £1/2 and £1/3 links on same template."
    wsErrors.Range("A4").Value = "- Please correct these errors and try again!"
    wsErrors.Activate
    Application.ScreenUpdating = True
    MsgBox "There where errors on template, see ERRORS sheet!", vbCritical, "Error!"
    Exit Sub
  End If
  
  ' Row markers were correct, now searching for headers
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      Arvo = wsTemplate.Cells(i, j).Value
      If Len(Arvo) > 2 Then 'Cell has data
        If Left(Arvo, 2) = "££" Then
          If EtsiOts(Mid(Arvo, 3), i, j, 1) = False Then Virhe = True
        ElseIf Left(Arvo, 1) = "£" Then
          If EtsiOts(Mid(Arvo, 5), i, j, CInt(Mid(Arvo, 2, 1))) = False Then Virhe = True
        End If
      End If
    Next j
  Next i
  
  If Virhe Then
    wsErrors.Activate
    Application.ScreenUpdating = True
    MsgBox "There were errors on the template! See ERRORS sheet.", vbCritical, "Error!"
  Else
    Sheets("Main").Activate
    Application.ScreenUpdating = True
    CheckOK = True
    MsgBox "Check OK!", vbOKOnly, "OK!"
  End If
  Exit Sub

CheckoutError:
  Application.ScreenUpdating = True
  
  Dim errMsg As String
  errMsg = "Error in Checkout: " & Err.Description & " (Error " & Err.Number & ")"
  
  ' Add helpful context for common errors
  If Err.Number = vbObjectError + 1 Then
    errMsg = errMsg & vbCrLf & vbCrLf & "TEMPLATE sheet is missing required marker(s)." & vbCrLf & _
             "Please ensure TEMPLATE contains all markers:" & vbCrLf & _
             "&&PAGE_HEADER_START, &&PAGE_HEADER_END" & vbCrLf & _
             "&&DOC_DATA_START, &&DOC_DATA_END" & vbCrLf & _
             "&&PAGE_FOOTER_START, &&PAGE_FOOTER_END" & vbCrLf & _
             "&&END"
  ElseIf Err.Number = 91 Then
    errMsg = errMsg & vbCrLf & vbCrLf & "This usually means a required sheet or object is missing." & vbCrLf & _
             "Ensure TEMPLATE, DB1, DB2, ERRORS, and Info sheets exist."
  End If
  
  MsgBox errMsg, vbCritical, "Checkout Error"
  
  ' Log to Immediate Window
  Debug.Print "Checkout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  
  Err.Clear
  On Error GoTo 0
End Sub

