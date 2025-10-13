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
' BeginFastMode: Poistaa väliaikaisesti Excelin UI-päivitykset, tapahtumat ja asettaa laskennan manuaaliseksi
' makrojen nopeuttamiseksi ja näytön välkkymisen estämiseksi.
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
' EndFastMode: Palauttaa Excelin UI- ja laskenta-asetukset edelliseen tilaan.
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
' HaeData: Hakee dataa Access-tietokannasta ODBC:n kautta, käyttäen faceplatessa määritettyjä SQL-kyselyitä.
' Täyttää DB1- ja DB2 -sheetit tuloksilla. Käyttää nopeaa tilaa suorituskyvyn parantamiseksi.
'''
Sub HaeData()
Dim Kanta As String
Dim sSQL(2) As String
Dim Valinta As Long
Dim i As Long
Dim TAULUKKO As QueryTable
Dim Yhteys As String

' Select which SQL query to use from Main sheet.
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
<<<<<<< HEAD
    ' Tyhjennä aiemmat datat ja aja kysely jokaiselle DB-sheetille
    Dim ws As Worksheet
    Set ws = Sheets("DB" & i)
=======
    ' Clear previous data and run query for each DB sheet
    Set ws = ThisWorkbook.Sheets("DB" & i)
>>>>>>> main
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
  .Delete ' Poista kysely päivityksen jälkeen, jotta yhteyksiä ei jää roikkumaan
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
' GenPrintout: Generoi uuden tulostetyökirjan käyttäen TEMPLATEa ja DB1:n dataa.
' Kopioi otsikot, alatunnisteet ja rungon, soveltaa muotoilut ja tallentaa tuloksen.
' Käyttää taulukkomuotoista siirtoa (array) rungon kirjoittamiseen nopeuden vuoksi. Kaikki muotoilut ja linkitykset säilytetään.
'''
Sub GenPrintout()
  If CheckOK = False Then
    MsgBox "Check data first!", vbCritical, "Error!"
  Else ' Data on tarkistettu ja valmis tulosteen luontiin.
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
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo 0
    
  Riveja = DocEnd - DocStart ' Runkodatan rivimäärä
  Sheets("DB1").Select
  ' Etsi DB1:stä viimeinen käytössä oleva rivi
  Recordeja = Cells.Find(What:="*", _
          After:=Range("A1"), _
          LookAt:=xlPart, _
          LookIn:=xlFormulas, _
          SearchOrder:=xlByRows, _
          SearchDirection:=xlPrevious, _
          MatchCase:=False).Row
    
  MacroWB = ActiveWorkbook.Name ' Tallenna makrotyökirjan nimi
  ' Kopioi Info-sheet uuteen työkirjaan (luo uuden työkirjan)
  Sheets("Info").Copy
  UusiWB = ActiveWorkbook.Name ' Tallenna uuden työkirjan nimi
  Cells.ClearComments ' Poista kommentit uudesta Info-sheetistä
  ' Kopioi TEMPLATE uuteen työkirjaan ja nimeä POSheetiksi
  Windows(MacroWB).Activate
  Sheets("TEMPLATE").Copy After:=Workbooks(UusiWB).Sheets(1)
  ActiveSheet.Name = POSheet


  'Copy Legend and Revisions sheets to new workbook
    Windows(MacroWB).Activate
  Sheets("Legend").Copy After:=Workbooks(UusiWB).Sheets(2) ' Kopioi Legend
  Windows(MacroWB).Activate
<<<<<<< HEAD
  Sheets("Revisions").Copy After:=Workbooks(UusiWB).Sheets(1) ' Kopioi Revisions
      Dim ViimRivi As Long
      Dim Riveja As Long
      Dim Recordeja, Recordeja2 As Long
      Dim Kerta As Long
      Dim Alku As Long
      Dim Apu As Long
      Dim Sivunvaihtoja As Long
      Dim Tiedosto As String
      Dim i As Long
      Dim Oletus As String
      Dim Sarjoja As Long
=======
  Sheets("Revisions").Copy After:=Workbooks(UusiWB).Sheets(1) ' Copy Revisions
  Windows(UusiWB).Activate
  PopulateRevisionsSimple 'Fast direct write of revision data without comment loops
>>>>>>> main
    Sheets(POSheet).Select
  ' Tyhjennä kaikki data tulostesheetistä (sarakkeiden leveydet säilyvät)
    Cells.Clear
  ' Poista kaikki muodot/kuvat, jos niitä on
    If ActiveSheet.Shapes.Count > 0 Then
      For i = 1 To ActiveSheet.Shapes.Count
        ActiveSheet.Shapes(1).Delete
      Next i
    End If
    
    
    'Copy header rows from TEMPLATE to printout
  ViimRivi = 1
  ' Kopioi otsikkorivit TEMPLATEsta tulostesheetille
  Windows(MacroWB).Activate
  Sheets("TEMPLATE").Rows(PHStart & ":" & PHEnd).Copy
  Windows(UusiWB).Activate
  Rows(ViimRivi & ":" & ViimRivi + PHEnd - PHStart).Select
  ActiveSheet.Paste
  ' Aseta tulostuksen otsikkorivit ja jäädytä ruudut otsikoiden alle
  ActiveSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  Cells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
'---------- [ FOOTER ] --------------------------
    ' Set up footers for first three sheets (Info, POSheet, Legend)
    For i = 1 To 3
      Sheets(i).Select
      Application.StatusBar = "Prosessing Footers: " & i & "/3"
  ' Vasen alatunniste: dokumentin tiedot
      ActiveSheet.PageSetup.LeftFooter = "&8Document: " & DIMetsoDocNo & Chr(10) _
                                       & "&8Revision: " & DIRevID & " - " & DIRevDate & Chr(10) _
                                       & "&8Status: " & DIStatus
  ' Keskialatunniste: asiakas ja tehdas
      ActiveSheet.PageSetup.CenterFooter = "&8 " & DICustomer & Chr(10) _
                                         & "&8 " & DIMill & Chr(10) _
                                         & "&8 " & DIDepartName & Chr(10) _
                                         & "&8 " & DIDocName2
  ' Oikea alatunniste: projekti ja tiedoston nimi
      ActiveSheet.PageSetup.RightFooter = "&8Project: " & DIProjNo & Chr(10) _
                                        & "&8File: &F" & Chr(10) _
                                        & "&8Page &P(&N)"
    Next i
    Sheets(POSheet).Select
'---------- [ FOOTER ] --------------------------
    Kerta = 0
    VaihdaLinkit 1, ViimRivi, Kerta

    'ActiveSheet.HPageBreaks(1).Location.Row
    Sivunvaihtoja = ActiveSheet.HPageBreaks.Count
    Sarjoja = 0
  ' --- Optimoitu: DB1-datan massasiirto POSheetiin arrayn avulla ---
  ' Tämä lohko lataa DB1-datan 2D-taulukkoon ja kirjoittaa sen kerralla nopeuden vuoksi.
    Dim dbData As Variant
    Dim dataRows As Long, dataCols As Long
    Dim destStartRow As Long, destEndRow As Long
    Dim destSheet As Worksheet
    Set destSheet = Sheets(POSheet)
<<<<<<< HEAD
  ' Etsi datan alue DB1:stä (rivistä 2 viimeiseen riviin, kaikki sarakkeet)
=======
  ' Find data range in DB1 (from row 2 to last row, all columns)
    Windows(MacroWB).Activate
>>>>>>> main
    With Sheets("DB1")
        dataRows = .Cells(.Rows.Count, 1).End(xlUp).Row
        ' Use Sarakkeita from TEMPLATE instead of DB1's last column to prevent extra columns
        dataCols = Sarakkeita
        If dataRows >= 2 And dataCols >= 1 Then
            dbData = .Range(.Cells(2, 1), .Cells(dataRows, dataCols)).Value
        Else
            dbData = Empty
        End If
    End With
<<<<<<< HEAD
  ' Kirjoita array-datan POSheetille oikeaan kohtaan (otsikkorivien jälkeen)
=======
    Windows(UusiWB).Activate
  ' Write data array to POSheet at the correct position (after header rows)
>>>>>>> main
    destStartRow = ViimRivi
    destEndRow = destStartRow + UBound(dbData, 1) - 1
    If Not IsEmpty(dbData) Then
        destSheet.Range(destSheet.Cells(destStartRow, 1), destSheet.Cells(destEndRow, dataCols)).Value = dbData
  ' Valinnainen: käytä vuoroväritystä joka toiselle blokille (kuten aiemmin)
  ' (säilyttää alkuperäisen ulkoasun)
        For i = 0 To UBound(dbData, 1) - 1 Step RMAX
            If (i \ RMAX) Mod 2 = 1 Then
                With destSheet.Range(destSheet.Cells(destStartRow + i, 1), destSheet.Cells(destStartRow + i + RMAX - 1, Sarakkeita)).Interior
                    .ColorIndex = 19
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                End With
            End If
        Next i
  ' Kutsu VaihdaLinkit jokaiselle blokille kuten aiemmin (hoitaa linkityksen)
        Kerta = 0
        For i = 0 To UBound(dbData, 1) - 1 Step RMAX
            VaihdaLinkit destStartRow + i, destStartRow + i + RMAX - 1, Kerta
            Kerta = Kerta + 1
        Next i
        ViimRivi = destEndRow + 1
    End If
  ' --- End optimized block ---
  
  ' Delete any extra columns beyond Sarakkeita to match TEMPLATE width
  Sheets(POSheet).Select
  If Columns.Count > Sarakkeita Then
    Dim lastCol As Long
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    If lastCol > Sarakkeita Then
      Columns(Sarakkeita + 1 & ":" & lastCol).Delete
    End If
  End If
  
  If AddFooter = True Then
  ' Kopioi alatunnisteen rivit TEMPLATEsta tulostesheetille
    Windows(MacroWB).Activate
    Sheets("TEMPLATE").Rows(PFStart & ":" & PFEnd).Copy
    Windows(UusiWB).Activate
    Rows(ViimRivi & ":" & ViimRivi + PFEnd - PFStart).Select
    ActiveSheet.Paste
  ' Aseta kaavat alatunnisteen summa-soluja varten
    Dim c As Range
    For Each c In Selection
      If Left(c.Value, 2) = "&&" Then
        c.Formula = "=sum(" & Cells(PHEnd, c.Column).Address & ":" & Cells(ViimRivi - 1, c.Column).Address & ")"
      End If
    Next c
  End If
  Cells.ClearComments ' Remove all comments from printout
  TeeLinkingKommentit ' Add comments to LINKING sheet for traceability
  
  ' Handle LINKING sheet visibility/deletion based on user preference
  On Error Resume Next
  If HideLINKING Then
    Sheets("LINKING").Visible = False
  Else
    Application.DisplayAlerts = False
    Sheets("LINKING").Delete ' Remove LINKING sheet if user requested
    Application.DisplayAlerts = True
  End If
  On Error GoTo 0
    Windows(UusiWB).Activate
    Worksheets(POSheet).Activate
    Application.StatusBar = False
    EndFastMode
  ' Kysy käyttäjältä tiedoston nimeä ja tallenna .xlsx-muodossa
    Oletus = DIPath & DIFile
    Tiedosto = InputBox("Give The File Name", "Save File", Oletus)
    If Tiedosto <> "" Then
      ActiveWorkbook.BuiltinDocumentProperties("Title").Value = POSheet
      ActiveWorkbook.SaveAs Tiedosto, xlOpenXMLWorkbook
    End If
  End If
End Sub
'''
' Checkout: Validates that all required headers and row markers exist in the TEMPLATE and DB1 sheets.
' Reports errors to the ERRORS sheet and sets CheckOK flag.
'''
Sub Checkout() 'Check if all headers exist in data
Dim i As Long
Dim j As Long
Dim Arvo As String
Dim Virhe As Boolean
Dim Apu As Long
CheckOK = False
RMAX = 0
Virhe = False
   Application.ScreenUpdating = False
   Sheets("ERRORS").Select
   Cells.Select
   Selection.Clear
   Range("A1").Select
   'Constants
   POSheet = Sheets("Main").Range("C16").Value
   On Error Resume Next
   HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
   On Error GoTo 0
   
   Sheets("TEMPLATE").Select
   Cells.ClearComments
   Cells(1, 1).Select
   'Find area markers
   PHStart = Cells.Find(What:="&&PAGE_HEADER_START").Row + 1
   PHEnd = Cells.Find(What:="&&PAGE_HEADER_END").Row - 1
   DocStart = Cells.Find(What:="&&DOC_DATA_START").Row + 1
   DocEnd = Cells.Find(What:="&&DOC_DATA_END").Row - 1
   Sarakkeita = Cells.Find(What:="&&END").Column
   PFStart = Cells.Find(What:="&&PAGE_FOOTER_START").Row + 1
   PFEnd = Cells.Find(What:="&&PAGE_FOOTER_END").Row - 1
   
   HaeDocTiedot 'Fetch document info from DB2 sheet
   VaihdaInfo   'Populate document info to Info sheet only (not Revisions during checkout)
   
   Sheets("TEMPLATE").Select
   'Search for row markers
   For i = DocStart To DocEnd
     For j = 1 To Sarakkeita
       Arvo = Cells(i, j).Value
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
   'Row markers were correct, now searching for headers
   For i = DocStart To DocEnd
     For j = 1 To Sarakkeita
       Arvo = Cells(i, j).Value
       If Len(Arvo) > 2 Then 'Cell has data
         If Left(Arvo, 2) = "££" Then
           If EtsiOts(Mid(Arvo, 3), i, j, 1) = False Then Virhe = True
         ElseIf Left(Arvo, 1) = "£" Then
           Apu = CInt(Mid(Arvo, 2, 1))
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