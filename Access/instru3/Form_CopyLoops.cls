Option Compare Database
Option Explicit

'================================================================================
' Form: CopyLoops
' Purpose: Import loop records and device tables from external Access database
' Updated: 2025-11-11 - Added VBA7/64-bit support
'
' Description:
'   Loop import utility that:
'   - Links to external instru3 database
'   - Displays available loops from source database
'   - Imports selected loops to LOOPS table
'   - Imports associated device records from devTbl* tables
'   - Handles duplicate loop prevention
'   - Provides detailed status feedback
'
' Dependencies:
'   - LOOPS table (target for loop import)
'   - LOOPLINK temporary linked table
'   - devTable table (target for device import)
'   - devTbl* tables (source device tables in external database)
'   - comdlg32.dll (Windows Common Dialog for file selection)
'
' Controls:
'   - Kanta: TextBox for database path
'   - Loopit: Subform displaying available loops
'   - TTiedot: TextBox for status messages
'   - ValitseKanta: Button to select source database
'   - HaeLoopit: Button to import loops
'   - Ulos: Button to close form
'================================================================================

' Module-level variable for source database connection
Dim LahdeDB As DAO.Database  ' Source database connection (persists during session)

'--------------------------------------------------------------------------------
' Event: Form_Unload
' Purpose: Clean up temporary tables and database connections on form close
'
' Process:
'   1. Delete temporary LOOPLINK table if it exists
'   2. Close source database connection
'   3. Clear all object references
'
' Notes:
'   - Uses Resume Next to prevent errors if table doesn't exist
'   - Always runs when form closes
'   - Ensures no lingering temporary tables or connections
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  CurrentDb.TableDefs.Delete "LOOPLINK"  ' Remove temporary linked table
  Err.Clear
  On Error GoTo 0
  If Not LahdeDB Is Nothing Then
    LahdeDB.Close
    Set LahdeDB = Nothing
  End If
End Sub
'--------------------------------------------------------------------------------
' Event: HaeLoopit_Click
' Purpose: Import selected loops and associated device data from source database
'
' Process:
'   1. Validate source database is selected
'   2. Copy loop records from LOOPLINK to LOOPS table
'   3. Iterate through all devTbl* tables in source database
'   4. Import device records matching selected loops
'   5. Set TableName field for imported device records
'   6. Track duplicates and provide status feedback
'
' Parameters:
'   - Uses Loopit subform as source (linked to LOOPLINK table)
'   - Imports to LOOPS and devTable tables
'
' Notes:
'   - Prevents duplicate loops (checks LOOPS table first)
'   - Imports from devTbl* tables and devTable (case insensitive)
'   - Excludes devTblCommon table from import
'   - Provides detailed status messages in TTiedot textbox
'   - Uses SysCmd for progress indication
'--------------------------------------------------------------------------------
Private Sub HaeLoopit_Click()
On Error GoTo ErrorHandler
    Dim OliJo As String  ' Tracks loops that already existed (duplicates)
    Dim LahdeTaulu As DAO.Recordset  ' Source table recordset (from external DB)
    Dim KohdeTaulu As DAO.Recordset  ' Target table recordset (current DB)
    Dim Kopioitavat() As String  ' Array of loops to import (format: "AreaCode-LoopNo")
    Dim i As Integer, j As Integer, a As Integer  ' Loop counters
    Dim Taulukko As String  ' Current table name being processed
    
    ReDim Kopioitavat(0)  ' Initialize array
    
    ' Validate source database is selected
    If LahdeDB Is Nothing Then
        MsgBox "Select database first!", vbCritical, "Get Loops"
        Exit Sub
    End If
  
  'Kopioidaan ensin Loopit Loops taulukkoon - First copy Loops to LOOPS table
  TTiedot.Value = ""  ' Clear status messages
  
  ' Iterate through selected loops in subform
  With Loopit.Form.RecordsetClone
    .MoveFirst
    Do While Not .EOF
      ' Check if loop already exists in target database
      Set KohdeTaulu = CurrentDb.OpenRecordset("SELECT * FROM LOOPS WHERE Areacode='" & .Fields("AreaCode").Value & "' AND LoopNo='" & .Fields("LoopNo").Value & "'")
      If KohdeTaulu.RecordCount = 0 Then
        ' New loop - add to import list and copy record
        ReDim Preserve Kopioitavat(UBound(Kopioitavat) + 1)
        Kopioitavat(UBound(Kopioitavat)) = .Fields("AreaCode").Value & "-" & .Fields("LoopNo").Value
        KohdeTaulu.AddNew
        On Error Resume Next
        ' Copy all fields from source to target
        For i = 0 To .Fields.Count - 1
          KohdeTaulu.Fields(.Fields(i).Name).Value = .Fields(i).Value
        Next i
        KohdeTaulu.Update
        Err.Clear
        On Error GoTo 0
      Else
        ' Loop already exists - track for duplicate report
        OliJo = OliJo & .Fields("AreaCode") & "-" & .Fields("LoopNo") & vbCrLf
      End If
      .MoveNext
    Loop
  End With
  
  ' Import device records from devTbl* tables
  For i = 0 To LahdeDB.TableDefs.Count - 1
    Taulukko = LCase(LahdeDB.TableDefs(i).Name)
    Set KohdeTaulu = CurrentDb.OpenRecordset("devTable")
    
    ' Check if table is a device table (devTbl* or devTable, excluding devTblCommon)
    If (Left(Taulukko, 6) = "devtbl" And Left(Taulukko, 12) <> "devtblcommon") Or (LCase(Taulukko) = "devtable") Then
      SysCmd acSysCmdSetStatus, "Updating table: " & LahdeDB.TableDefs(i).Name  ' Show progress
      Set LahdeTaulu = LahdeDB.OpenRecordset(Taulukko)
      
      ' Iterate through all records in source device table
      Do While Not LahdeTaulu.EOF 'Käydään läpi kaikki taulukon recordit - Go through all table records
        ' Check if this record belongs to one of the imported loops
        For j = 1 To UBound(Kopioitavat) 'Tarkistetaan että ko. recordi kuuluu kopioitaviin - Check if record should be copied
          If Kopioitavat(j) = LahdeTaulu.Fields("AreaCode").Value & "-" & LahdeTaulu.Fields("LoopNo").Value Then
            ' Record belongs to imported loop - copy it
            KohdeTaulu.AddNew
            ' Copy all fields
            For a = 0 To LahdeTaulu.Fields.Count - 1
              KohdeTaulu.Fields(LahdeTaulu.Fields(a).Name).Value = LahdeTaulu.Fields(a).Value
              If Err <> 0 Then
                ' Log field copy errors
                TTiedot.Value = TTiedot.Value & "Problem with field: " & LahdeDB.TableDefs(i).Name & "!" & LahdeTaulu.Fields(a).Name
                Err.Clear
              End If
            Next a
            ' Set TableName field for device records (except devTable which already has it)
            If LCase(Taulukko) <> "devtable" Then
              KohdeTaulu.Fields("TABLEName").Value = LahdeDB.TableDefs(i).Name
            End If
            KohdeTaulu.Update
            Exit For  ' Found match, move to next record
          End If
        Next j
        LahdeTaulu.MoveNext
      Loop
    End If
  Next i
  
  ' Display duplicate loop report
  If OliJo <> "" Then
    TTiedot.Value = TTiedot.Value & "Following Loops were already in current database:" & vbCrLf & OliJo
  End If
  TTiedot.Value = TTiedot.Value & "Ready!"
  SysCmd acSysCmdClearStatus
  MsgBox "Ready!", vbInformation, "Get Loops"
  
  ' Cleanup recordsets
  If Not KohdeTaulu Is Nothing Then
    KohdeTaulu.Close
    Set KohdeTaulu = Nothing
  End If
  If Not LahdeTaulu Is Nothing Then
    LahdeTaulu.Close
    Set LahdeTaulu = Nothing
  End If
  Exit Sub

ErrorHandler:
  SysCmd acSysCmdClearStatus
  MsgBox "Error during loop import: " & Err.Description, vbCritical, "Get Loops"
  On Error Resume Next
  If Not KohdeTaulu Is Nothing Then KohdeTaulu.Close
  If Not LahdeTaulu Is Nothing Then LahdeTaulu.Close
  Set KohdeTaulu = Nothing
  Set LahdeTaulu = Nothing
  On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Event: Ulos_Click
' Purpose: Close the form
'
' Notes:
'   - Simple close button
'   - Form_Unload event will handle cleanup
'--------------------------------------------------------------------------------
Private Sub Ulos_Click()
  DoCmd.Close
End Sub
'--------------------------------------------------------------------------------
' Event: ValitseKanta_Click
' Purpose: Open file dialog to select source Access database
'
' Process:
'   1. Determine initial directory (current database folder or previously selected)
'   2. Display Windows Common Dialog for Access database selection
'   3. Store selected path in Kanta textbox
'   4. Open database connection (stored in LahdeDB module variable)
'   5. Delete existing LOOPLINK temporary table if exists
'   6. Create new linked table LOOPLINK pointing to Loops table in source DB
'   7. Set subform recordsource to display available loops
'
' Notes:
'   - Uses GetOpenFileName API from comdlg32.dll
'   - Filters for .accdb files only
'   - Creates temporary linked table for loop browsing
'   - Source database remains open until form closes
'   - LOOPLINK table is deleted in Form_Unload
'--------------------------------------------------------------------------------
Private Sub ValitseKanta_Click()
On Error GoTo ErrorHandler
    'Valitaan tietokanta - Select database
    Dim OpenFile As OPENFILENAME  ' Windows Common Dialog structure
    Dim lReturn As Long  ' Dialog return value
    Dim Filtteri As String  ' File type filter
    Dim AHakem As String  ' Initial directory path
    Dim Otsikko As String  ' Dialog title
    Dim i As Integer  ' Loop counter
    Dim Linkki As DAO.TableDef  ' Temporary linked table definition
    
    
    ' Determine initial directory from previously selected database or current database
    If InStr(Nz(Kanta.Value), "\") Then
      AHakem = Kanta.Value
    Else
      AHakem = Application.CurrentDb.Name
    End If
    
    ' Extract directory path (remove filename)
    For i = 1 To Len(AHakem)
      If Mid(AHakem, Len(AHakem) - i, 1) = "\" Then
        AHakem = Left(AHakem, Len(AHakem) - i)
        Exit For
      End If
    Next i
    
    ' Configure file dialog
    Filtteri = "Microsoft Access Databases (*.accdb)" & Chr(0) & "*.ACCDB" & Chr(0)
    Otsikko = "Choose Instru Database"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = Me.Hwnd  ' Parent window handle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)  ' Buffer for filename
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    
    ' Display file dialog
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then
      ' Extract selected filename (remove null terminator)
      Kanta.Value = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
      
      ' Open source database connection
      Set LahdeDB = OpenDatabase(Kanta.Value)
      
      ' Clean up any existing LOOPLINK table
      On Error Resume Next
      CurrentDb.TableDefs.Delete "LOOPLINK"
      Err.Clear
      On Error GoTo 0
      
      ' Create linked table to Loops table in source database
      Set Linkki = CurrentDb.CreateTableDef("LOOPLINK")
      Linkki.Connect = ";DATABASE=" & Kanta.Value
      Linkki.SourceTableName = "Loops"
      CurrentDb.TableDefs.Append Linkki
      Set Linkki = Nothing
      
      ' Display available loops in subform
      Loopit.Form.RecordSource = "SELECT * FROM LOOPLINK"
    End If
    Me.Refresh
    Exit Sub

ErrorHandler:
    MsgBox "Error selecting database: " & Err.Description, vbCritical, "Choose Database"
    On Error Resume Next
    Set Linkki = Nothing
    On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Event: Form_Resize
' Purpose: Adjust subform size when form window is resized
'
' Process:
'   - Resize Loopit subform width to match form window
'   - Adjust height if sufficient vertical space available
'
' Notes:
'   - Maintains 400 twips margin on right side
'   - Only adjusts height if window taller than 3300 twips
'   - Provides responsive UI for better loop viewing
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
  Loopit.Width = Me.WindowWidth - 400  ' Maintain right margin
  If Me.WindowHeight > 3300 Then
    Loopit.Height = Me.WindowHeight - 3300  ' Adjust height for better viewing
  End If
End Sub

