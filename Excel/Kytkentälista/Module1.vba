Public CheckOK As Boolean
Public PHStart As Long
Public PHEnd As Long
Public PFStart As Long
Public PFEnd As Long
Public DocStart As Long
Public DocEnd As Long
Public Sarakkeita As Long
Public RMAX As Integer
Public FPSheet As String
Public POSheet As String
Public GenFP As Boolean
Public GenHeader As Boolean
Public GenDocHeader As Boolean
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
' Module1.vba - Main logic for Kytkentälista Excel macro system
' Handles data fetching from Access, printout generation, and performance optimizations.
'''

' Performance/UX state (to minimize screen flashing and speed up macros)
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
' Populates DB1 and DB2 sheets with the results. Uses fast mode for performance.
'''
Sub HaeData()
Dim Kanta As String
Dim sSQL(2) As String
Dim Valinta As Integer
Dim i As Long
Dim TAULUKKO As QueryTable
Dim Yhteys As String

'Tuhotaan DB-sheetit, jotta ne saadaan varmasti tyhjäksi
'Sheets("DB1").Delete
'Sheets("DB2").Delete
'Sheets.Add After:=Sheets("Main")
'ActiveSheet.Name = "DB1"
'Sheets.Add After:=Sheets("DB1")
'ActiveSheet.Name = "DB2"

  ' Get Valinta selection
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
  Dim ws As Worksheet
  
  On Error GoTo ErrorHandler
  
  For i = 1 To 2
    ' Clear previous data and run query for each DB sheet
    Set ws = ThisWorkbook.Sheets("DB" & i)
    ws.Cells.Clear
    ' Skip Excel-based queries (_qryForExcel) - only process Access database queries
    If sSQL(i) <> "" And InStr(1, sSQL(i), "_qryForExcel", vbTextCompare) = 0 Then
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
        .Delete ' Remove query after refresh to avoid leaving connections
      End With
      Set TAULUKKO = Nothing
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
' Copies headers, footers, and main data body, applies formatting, and saves the result.
' Uses array-based transfer for main data for speed. All formatting and linking logic preserved.
'''
Sub GenPrintout()
  If CheckOK = False Then
    MsgBox "Check data first!", vbCritical, "Error!"
  Else ' Data has been checked and is ready for printout generation.
    Dim MacroWB As String
    Dim UusiWB As String
    Dim ViimRivi As Long
    Dim Riveja As Long
    Dim Recordeja As Long, Recordeja2 As Long
    Dim Kerta As Long
    Dim Alku As Long
    Dim Apu As Long
    Dim Sivunvaihtoja As Long
    Dim Tiedosto As String
    Dim i As Long
    Dim Oletus As String
    Dim Sarjoja As Long
    
  BeginFastMode
    
  AddFooter = Sheets("Main").AddFooter.Value ' User option from faceplate
    
  Riveja = DocEnd - DocStart ' Number of rows in the document body
  Sheets("DB1").Select
  ' Find last used row in DB1
  Recordeja = Cells.Find(What:="*", _
          After:=Range("A1"), _
          LookAt:=xlPart, _
          LookIn:=xlFormulas, _
          SearchOrder:=xlByRows, _
          SearchDirection:=xlPrevious, _
          MatchCase:=False).Row
    
  MacroWB = ActiveWorkbook.Name ' Store macro workbook name
  ' Copy Info sheet to new workbook (creates new workbook)
  Sheets("Info").Copy
  UusiWB = ActiveWorkbook.Name ' Store new workbook name
  Cells.ClearComments ' Remove comments from new Info sheet
  ' Copy TEMPLATE to new workbook and rename as POSheet
  Windows(MacroWB).Activate
  Sheets("TEMPLATE").Copy After:=Workbooks(UusiWB).Sheets(1)
  ActiveSheet.Name = POSheet
    
    
    'Kopioidaan sitten Legend ja Revisions sheetit uuten työkirjaan
    Windows(MacroWB).Activate
  Sheets("Legend").Copy After:=Workbooks(UusiWB).Sheets(2) ' Copy Legend
  Windows(MacroWB).Activate
  Sheets("Revisions").Copy After:=Workbooks(UusiWB).Sheets(1) ' Copy Revisions
    Sheets(POSheet).Select
    ' Clear all data from printout sheet (keeps column widths)
    Cells.Clear
    ' Remove all shapes/images if present
    If ActiveSheet.Shapes.Count > 0 Then
      For i = 1 To ActiveSheet.Shapes.Count
        ActiveSheet.Shapes(1).Delete
      Next i
    End If
    
    
    'Kopioidaan mahdolliset otsikkotiedot Prinout sheetille
  ViimRivi = 1
  ' Copy header rows from TEMPLATE to printout
  Windows(MacroWB).Activate
  Sheets("TEMPLATE").Rows(PHStart & ":" & PHEnd).Copy
  Windows(UusiWB).Activate
  Rows(ViimRivi & ":" & ViimRivi + PHEnd - PHStart).Select
  ActiveSheet.Paste
  ' Set print title rows and freeze panes below header
  ActiveSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  Cells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
'---------- [ ALATUNNISTEET ] --------------------------
    ' Set up footers for first three sheets (Info, POSheet, Legend)
    For i = 1 To 3
      Sheets(i).Select
      Application.StatusBar = "Prosessing Footers: " & i & "/3"
      ' Left footer: document info
      ActiveSheet.PageSetup.LeftFooter = "&8Document: " & DIMetsoDocNo & Chr(10) _
                                       & "&8Revision: " & DIRevID & " - " & DIRevDate & Chr(10) _
                                       & "&8Status: " & DIStatus
      ' Center footer: customer and mill info
      ActiveSheet.PageSetup.CenterFooter = "&8 " & DICustomer & Chr(10) _
                                         & "&8 " & DIMill & Chr(10) _
                                         & "&8 " & DIDepartName & Chr(10) _
                                         & "&8 " & DIDocName2
      ' Right footer: project and file info
      ActiveSheet.PageSetup.RightFooter = "&8Project: " & DIProjNo & Chr(10) _
                                        & "&8File: &F" & Chr(10) _
                                        & "&8Page &P(&N)"
    Next i
    Sheets(POSheet).Select
'---------- [ ALATUNNISTEET ] --------------------------
    Kerta = 0
    VaihdaLinkit 1, ViimRivi, Kerta

    'ActiveSheet.HPageBreaks(1).Location.Row
    Sivunvaihtoja = ActiveSheet.HPageBreaks.Count
    Sarjoja = 0
  ' --- Optimized: Bulk copy DB1 data to POSheet using array ---
  ' This block loads all DB1 data into a 2D array and writes it in one operation for speed.
    Dim dbData As Variant
    Dim dataRows As Long, dataCols As Long
    Dim destStartRow As Long, destEndRow As Long
    Dim destSheet As Worksheet
    Set destSheet = Sheets(POSheet)
  ' Find data range in DB1 (from row 2 to last row, all columns)
    With Sheets("DB1")
        dataRows = .Cells(.Rows.Count, 1).End(xlUp).Row
        dataCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        If dataRows >= 2 And dataCols >= 1 Then
            dbData = .Range(.Cells(2, 1), .Cells(dataRows, dataCols)).Value
        Else
            dbData = Empty
        End If
    End With
  ' Write data array to POSheet at the correct position (after header rows)
    destStartRow = ViimRivi
    destEndRow = destStartRow + UBound(dbData, 1) - 1
    If Not IsEmpty(dbData) Then
        destSheet.Range(destSheet.Cells(destStartRow, 1), destSheet.Cells(destEndRow, dataCols)).Value = dbData
  ' Optional: Apply alternating row color for every other block (as before)
  ' (keeps the same look as the original macro)
        For i = 0 To UBound(dbData, 1) - 1 Step RMAX
            If (i \ RMAX) Mod 2 = 1 Then
                With destSheet.Range(destSheet.Cells(destStartRow + i, 1), destSheet.Cells(destStartRow + i + RMAX - 1, Sarakkeita)).Interior
                    .ColorIndex = 19
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                End With
            End If
        Next i
  ' Call VaihdaLinkit for each block as before (handles linking logic)
        Kerta = 0
        For i = 0 To UBound(dbData, 1) - 1 Step RMAX
            VaihdaLinkit destStartRow + i, destStartRow + i + RMAX - 1, Kerta
            Kerta = Kerta + 1
        Next i
        ViimRivi = destEndRow + 1
    End If
  ' --- End optimized block ---
  If AddFooter = True Then
    ' Copy footer rows from TEMPLATE to printout
    Windows(MacroWB).Activate
    Sheets("TEMPLATE").Rows(PFStart & ":" & PFEnd).Copy
    Windows(UusiWB).Activate
    Rows(ViimRivi & ":" & ViimRivi + PFEnd - PFStart).Select
    ActiveSheet.Paste
    ' Set up formulas for footer summary cells
    Dim c As Range
    For Each c In Selection
      If Left(c.Value, 2) = "&&" Then
        c.Formula = "=sum(" & Cells(PHEnd, c.Column).Address & ":" & Cells(ViimRivi - 1, c.Column).Address & ")"
      End If
    Next c
  End If
  Cells.ClearComments ' Remove all comments from printout
  TeeLinkingKommentit ' Add comments to LINKING sheet for traceability
'    If HideLINKING Then
'      Sheets("LINKING").Visible = False
'    End If
  Application.DisplayAlerts = False
  Sheets("LINKING").Delete ' Remove LINKING sheet if user requested
  Application.DisplayAlerts = True
    Windows(UusiWB).Activate
    Worksheets(POSheet).Activate
    Application.StatusBar = False
    EndFastMode
    ' Prompt user for file name and save as .xlsx
    Oletus = DIPath & DIFile
    Tiedosto = InputBox("Give The File Name", "Save File", Oletus)
    If Tiedosto <> "" Then
      ActiveWorkbook.BuiltinDocumentProperties("Title").Value = POSheet
      ActiveWorkbook.SaveAs Tiedosto, xlOpenXMLWorkbook
    End If
End Sub
'''
' Checkout: Validates that all required headers and row markers exist in the TEMPLATE and DB1 sheets.
' Reports errors to the ERRORS sheet and sets CheckOK flag.
'''
Sub Checkout() 'Tämä tarkistaa löytyvätkö kaikki otsikot datasta
Dim i As Long
Dim j As Long
Dim Arvo As String
Dim Virhe As Boolean
Dim Apu As Long
RMAX = 0
Virhe = False
   Application.ScreenUpdating = False
   Sheets("ERRORS").Select
   Cells.Select
   Selection.Clear
   Range("A1").Select
   'Vakiot
   POSheet = Sheets("Main").Range("C16").Value
'   HideLINKING = Sheets("Main").HLINKING.Value
   
   Sheets("TEMPLATE").Select
   'Etsitään ensin alueet
   PHStart = Cells.Find(What:="&&PAGE_HEADER_START").Row + 1
   PHEnd = Cells.Find(What:="&&PAGE_HEADER_END").Row - 1
   DocStart = Cells.Find(What:="&&DOC_DATA_START").Row + 1
   DocEnd = Cells.Find(What:="&&DOC_DATA_END").Row - 1
   Sarakkeita = Cells.Find(What:="&&END").Column
   PFStart = Cells.Find(What:="&&PAGE_FOOTER_START").Row + 1
   PFEnd = Cells.Find(What:="&&PAGE_FOOTER_END").Row - 1
   
   HaeDocTiedot 'Hakee dokumentin tiedot DB2-sheetiltä
   VaihdaInfo   'Vaihtaa dokumentin tiedot info sheetille
   VaihdaInfo ("Revisions") 'Tiedot revisions sheetille
   'Haetaan ensin kerralla kopioitavien rivien määrä eli rivitysmerkinnät
   For i = DocStart To DocEnd
     For j = 1 To Sarakkeita
       Arvo = Cells(i, j).Value
       If Len(Arvo) > 2 Then 'Solussa on tietoa
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
     Sheets("ERRORS").Select
     Cells.Delete
     Range("A1").Select
     ActiveCell.Value = "Error on declaration!"
     ActiveCell.Font.Bold = True
     Range("A2").Value = "- You cannot have ££ and £1/2 links on same template."
     Range("A3").Value = "- Neither you can have £1/2 and £1/3 links on same template."
     Range("A4").Value = "- Please correct these errors and try again!"
     MsgBox "There where errors on template, see ERRORS sheet!", vbCritical, "Error!"
     Exit Sub
   End If
   'Rivitys merkinnätolivat oikein, etsitään otsikoita
   For i = DocStart To DocEnd
     For j = 1 To Sarakkeita
       Arvo = Cells(i, j).Value
       If Len(Arvo) > 2 Then 'Solussa on tietoa
         If Left(Arvo, 2) = "££" Then
           If EtsiOts(Mid(Arvo, 3), i, j, 1&) = False Then Virhe = True
         ElseIf Left(Arvo, 1) = "£" Then
           Apu = CLng(Mid(Arvo, 2, 1))
           If EtsiOts(Mid(Arvo, 5), i, j, Apu) = False Then Virhe = True
         End If
       End If
     Next j
   Next i
   If Virhe Then
     Sheets("ERRORS").Select
     Application.ScreenUpdating = True
     MsgBox "There were errors on the template! See ERRORS sheet.", vbCritical, "Error!"
   Else
     Sheets("Main").Select
     Application.ScreenUpdating = True
     CheckOK = True
     MsgBox "Check OK!", vbOKOnly, "OK!"
   End If
End Sub