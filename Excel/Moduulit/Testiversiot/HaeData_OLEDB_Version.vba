'''
' HaeData - IMPROVED VERSION using OLE DB instead of ODBC
' This version should work better with 64-bit Office 365
' and doesn't require ODBC driver workarounds
'''
Sub HaeData_OLEDB()
Dim Kanta As String
Dim sSQL(2) As String
Dim Valinta As Long
Dim i As Long
Dim TAULUKKO As QueryTable
Dim Yhteys As String

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
  
  ' ===================================================================
  ' IMPROVED: Using OLE DB instead of ODBC
  ' ===================================================================
  ' ACE.OLEDB.16.0 is included with Office 365 64-bit
  ' Works with both .mdb and .accdb files
  ' Better compatibility with Access queries and 64-bit Office
  ' ===================================================================
  
  ' Determine connection string based on file extension
  Dim fileExt As String
  fileExt = LCase(Right(Kanta, 4))
  
  If fileExt = ".mdb" Or fileExt = ".accdb" Then
    ' Use ACE.OLEDB.16.0 (comes with Office 365 64-bit)
    Yhteys = "OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & Kanta
  Else
    ' Fallback to ODBC if file extension is unusual
    Yhteys = "ODBC;DBQ=" & Kanta & ";Driver={Microsoft Access Driver (*.mdb, *.accdb)}"
  End If
  
  Dim ws As Worksheet
  
  On Error GoTo ErrorHandler
  
  For i = 1 To 2
    ' Clear previous data and run query for each DB sheet
    Set ws = ThisWorkbook.Sheets("DB" & i)
    ws.Cells.Clear
    If sSQL(i) <> "" Then
      
      ' No need for bracket workarounds with OLE DB!
      Dim sqlQuery As String
      sqlQuery = sSQL(i)

      Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
      With TAULUKKO
        .Sql = sqlQuery
        .FieldNames = True
        .RefreshStyle = xlInsertDeleteCells
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .HasAutoFormat = True
        .SaveData = True
        .BackgroundQuery = False
        
        On Error Resume Next
        .Refresh
        If Err.Number <> 0 Then
          MsgBox "Error refreshing DB" & i & ":" & vbCrLf & vbCrLf & _
                 Err.Description & vbCrLf & vbCrLf & _
                 "Connection: " & Yhteys & vbCrLf & _
                 "Query: " & sqlQuery & vbCrLf & vbCrLf & _
                 "This sheet will be empty.", vbCritical, "Query Error"
          Err.Clear
        End If
        On Error GoTo ErrorHandler
        
        .Delete ' Remove query after refresh
      End With
      Set TAULUKKO = Nothing
      
      On Error GoTo ErrorHandler
      
      ' Check if query returned any data
      Dim rowCount As Long
      rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
      
      If rowCount <= 1 Then
        If i = 2 Then
          MsgBox "WARNING: DB2 query returned no data!" & vbCrLf & vbCrLf & _
                 "This means the Info sheet will be empty." & vbCrLf & vbCrLf & _
                 "Check the query in Main sheet.", vbExclamation, "Query Returned No Data"
        End If
      End If
    End If
  Next i
  EndFastMode
  MsgBox "Data brought successfully!", vbOKOnly, "Ready"
  Sheets("Main").Select
  Exit Sub
  
ErrorHandler:
  EndFastMode
  MsgBox "Database Error: " & Err.Description & vbCrLf & vbCrLf & _
         "Database: " & Kanta & vbCrLf & _
         "Connection: " & Yhteys & vbCrLf & _
         "SQL Query " & i & ": " & sSQL(i), vbCritical, "Database Connection Error"
  Err.Clear
  Sheets("Main").Select
End Sub
