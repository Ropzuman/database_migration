Option Compare Database
Option Explicit

'================================================================================
' Form: zFunc
' Purpose: Utility functions for zDetails table management and batch operations
' Updated: 2025-11-12 - Added VBA7/64-bit support
'
' Description:
'   Collection of utility functions for:
'   - zDetails table maintenance (adding, updating, clearing records)
'   - Batch query execution for instrument loops
'   - Empty string to Null conversion across all tables
'   - Block placement numbering
'   - Finding next record functionality
'
' Dependencies:
'   - zDetails table (instrument details)
'   - Function_zdetails* queries
'   - InstrumentIndex table
'   - Multiple Tag* queries
'   - KaikkiTaulukot combo box (populated with table names)
'
' Controls:
'   - KaikkiTaulukot: ComboBox for table selection
'   - Multiple command buttons for different operations
'================================================================================

Private Sub Command1_Click()
On Error GoTo Err_Command1_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command1_Click:
    Exit Sub

Err_Command1_Click:
    MsgBox Err.Description
    Resume Exit_Command1_Click
    
End Sub


Private Sub Command10_Click()
DoCmd.OpenQuery "Function_zdetails15_loopsprintselect", acNormal, acEdit
End Sub

'--------------------------------------------------------------------------------
' Event: Command11_Click
' Purpose: Convert empty strings to Null across selected table
'
' Process:
'   1. Open selected table from KaikkiTaulukot combo
'   2. Iterate through all records
'   3. For each field, if value is empty string, set to Null
'   4. Update record with transaction support
'
' Notes:
'   - Improves database normalization
'   - Prevents empty strings in database
'   - Can be time-consuming for large tables
'   - Optimized: Uses Edit once per record, transaction support
'--------------------------------------------------------------------------------
Private Sub Command11_Click()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Current database
    Dim WS As DAO.Workspace  ' Workspace for transaction support
    Dim Taulu As DAO.Recordset  ' Selected table
    Dim i As Integer  ' Field counter
    Dim RecordsProcessed As Long  ' Counter for processed records
    Dim RecordsUpdated As Long  ' Counter for updated records
    Dim NeedsUpdate As Boolean  ' Flag if current record needs update
    
    Set DB = CurrentDb
    Set WS = DBEngine.Workspaces(0)  ' Get default workspace
    Set Taulu = DB.OpenRecordset(KaikkiTaulukot.Value, dbOpenDynaset)
    
    If Taulu.EOF Then
        MsgBox "Taulu on tyhjä", vbInformation, "Empty to Null"
        Taulu.Close
        Set Taulu = Nothing
        DB.Close
        Set DB = Nothing
        Exit Sub
    End If
    
    ' Begin transaction for better performance
    WS.BeginTrans
    
    Do While Not Taulu.EOF
      NeedsUpdate = False
      RecordsProcessed = RecordsProcessed + 1
      
      ' Check all fields first before editing
      For i = 0 To Taulu.Fields.Count - 1
        If Taulu.Fields(i) = "" Then
          If Not NeedsUpdate Then
            Taulu.Edit
            NeedsUpdate = True
          End If
          Taulu.Fields(i) = Null
        End If
      Next i
      
      ' Only update if changes were made
      If NeedsUpdate Then
        Taulu.Update
        RecordsUpdated = RecordsUpdated + 1
      End If
      
      Taulu.MoveNext
    Loop
    
    ' Commit transaction
    WS.CommitTrans
    
    MsgBox "Valmis: " & RecordsProcessed & " riviä käyty, " & RecordsUpdated & " päivitetty", vbInformation, "Empty to Null"  ' "Ready"
    
    ' Cleanup
    Taulu.Close
    Set Taulu = Nothing
    Set WS = Nothing
    DB.Close
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    ' Rollback transaction on error
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set WS = Nothing
    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Error converting empty strings: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' Event: Command3_Click
' Purpose: Assign block placement numbers to zDetails records
'
' Process:
'   1. Open sorted recordset (by Loop, then AllDevTag)
'   2. For each loop group, assign sequential block placement numbers (14, 15, 16...)
'   3. Reset counter for each new loop
'
' Notes:
'   - Starts block placement at 14
'   - Groups by Loop identifier ([AreaCode] & [Loopno] & [pageno])
'   - Used for AutoCAD block positioning
'--------------------------------------------------------------------------------
Private Sub Command3_Click()
On Error GoTo Err_Command3_Click
    Dim tbl As DAO.Recordset  ' zDetails table recordset
    Dim q As String  ' Current loop identifier
    Dim q1 As String  ' Previous loop (unused)
    Dim qq As String  ' Next loop identifier
    Dim qq1 As String  ' Next loop comparison (unused)
    Dim i As Integer  ' Block placement counter
    
    ' Open sorted recordset
    Set tbl = CurrentDb.OpenRecordset("SELECT [AreaCode] & [Loopno] & [pageno] AS Loop, zDetails.Blockplace, zDetails.AllDevTag FROM zDetails ORDER BY [AreaCode] & [Loopno] & [pageno], zDetails.AllDevTag;")
  
    tbl.MoveFirst
    Do
      i = 14  ' Start block placement at 14
      Do
        q = tbl!Loop
        tbl.Edit
        tbl!blockplace = i
        tbl.Update
        i = i + 1
        tbl.MoveNext
        If tbl.EOF Then Exit Do
        qq = tbl!Loop
        If q <> qq Then Exit Do  ' New loop group, reset counter
      Loop
    Loop Until tbl.EOF
  
    MsgBox "Homma hoidettu", vbInformation, "Block Placement"  ' "Job done"
    
    ' Cleanup
    tbl.Close
    Set tbl = Nothing
    Exit Sub

Err_Command3_Click:
    MsgBox "Error assigning block placements: " & Err.Description, vbCritical, "Error"
    On Error Resume Next
    If Not tbl Is Nothing Then tbl.Close
    Set tbl = Nothing
    On Error GoTo 0
End Sub

Private Sub Command4_Click()
DoCmd.OpenQuery "Function_zDetails03_lisättävät", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails04_FileID", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails09_update", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails14_Fileid2", acNormal, acEdit
End Sub


'--------------------------------------------------------------------------------
' Event: Command5_Click
' Purpose: Delete zDetails records that no longer exist in InstrumentIndex
'
' Process:
'   1. Find zDetails records with no matching record in Function_zdetails01
'   2. For each orphaned record, delete it from zDetails
'   3. Use hourglass cursor to indicate progress
'   4. Transaction support for data integrity
'
' Notes:
'   - Cleans up obsolete instrument detail records
'   - Uses LEFT JOIN to find unmatched records
'   - L = Chr(34) for quote character in SQL
'   - Optimized: Hourglass only shown during processing
'--------------------------------------------------------------------------------
Private Sub Command5_Click()
On Error GoTo ErrorHandler
    ' Clean up orphaned zDetails records
    Dim DB As DAO.Database  ' Current database
    Dim WS As DAO.Workspace  ' Workspace for transaction support
    Dim tbl As DAO.Recordset  ' Query results (orphaned records)
    Dim tbl2 As DAO.Recordset  ' zDetails deletion target
    Dim qry As DAO.Recordset  ' Query recordset (unused)
    Dim sSQL As String  ' Main query SQL
    Dim ssSQL As String  ' Deletion SQL
    Dim acode As String  ' Area code
    Dim LoopNo As String  ' Loop number
    Dim LoopSymb As String  ' Loop symbol
    Dim Suffix As String  ' Tag suffix
    Dim L As String  ' Quote character
    Dim DeleteCount As Long  ' Count of deleted records

    L = Chr$(34)  ' Double quote for SQL strings
    DeleteCount = 0

    Set DB = CurrentDb
    Set WS = DBEngine.Workspaces(0)  ' Get default workspace

    ' Build query to find orphaned zDetails records
    sSQL = ""
    sSQL = sSQL & "SELECT zDetails.AreaCode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix"
    sSQL = sSQL & " FROM zDetails LEFT JOIN Function_zdetails01 ON (zDetails.Suffix = Function_zdetails01.Suffix) AND (zDetails.LoopSymb = Function_zdetails01.LoopSymb) AND (zDetails.LoopNo = Function_zdetails01.LoopNo) AND (zDetails.AreaCode = Function_zdetails01.AreaCode)"
    sSQL = sSQL & " WHERE (((Function_zdetails01.AreaCode) Is Null) AND ((Function_zdetails01.LoopNo) Is Null) AND ((Function_zdetails01.LoopSymb) Is Null) AND ((Function_zdetails01.Suffix) Is Null));"

    Set tbl = DB.OpenRecordset(sSQL)

    If Not tbl.EOF Then
      DoCmd.Hourglass True
      WS.BeginTrans  ' Start transaction
      
      tbl.MoveFirst
      Do Until tbl.EOF
          acode = tbl!AreaCode
          LoopNo = tbl!LoopNo
          Suffix = tbl!Suffix
          LoopSymb = tbl!LoopSymb

          ' Build SQL to delete matching record
          ssSQL = ""
          ssSQL = ssSQL & "SELECT zDetails.AreaCode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix, zDetails.CirDWG, zDetails.Pageno, zDetails.AllDevTag, zDetails.Name, zDetails.Manufacturer, zDetails.Type, zDetails.ProcConn, zDetails.Faixatotal, zDetails.Faixa, zDetails.Blockname, zDetails.Blockplace, zDetails.[Select]"
          ssSQL = ssSQL & " FROM zDetails"
          ssSQL = ssSQL & " WHERE (((zDetails.AreaCode)=" & L & acode & L & ") AND ((zDetails.LoopNo)=" & L & LoopNo & L & ") AND ((zDetails.LoopSymb)=" & L & LoopSymb & L & ") AND ((zDetails.Suffix)=" & L & Suffix & L & "));"
          Set tbl2 = DB.OpenRecordset(ssSQL)

          If Not tbl2.EOF Then
              tbl2.Delete
              DeleteCount = DeleteCount + 1
          End If
          
          tbl2.Close
          Set tbl2 = Nothing
          tbl.MoveNext
      Loop
      
      WS.CommitTrans  ' Commit transaction
      DoCmd.Hourglass False
      
      tbl.Close
      Set tbl = Nothing
      MsgBox "Poistettu " & DeleteCount & " vanhentunutta tietuetta", vbInformation, "Cleanup Complete"  ' "Deleted X orphaned records"
    Else
      MsgBox "Ei vanhentuneita tietueita", vbInformation, "Cleanup Complete"  ' "No orphaned records found"
    End If
    
    ' Cleanup
    Set tbl2 = Nothing
    Set WS = Nothing
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not tbl Is Nothing Then tbl.Close
    Set tbl = Nothing
    Set tbl2 = Nothing
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Error deleting orphaned records: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Command6_Click()
DoCmd.OpenQuery "Function_zDetails06_clearprint", acNormal, acEdit
End Sub

Private Sub Command7_Click()
DoCmd.OpenQuery "Function_zDetails05_printselect", acNormal, acEdit
End Sub

Private Sub Command8_Click()
DoCmd.OpenQuery "Function_Zdetails08_Clearzdetailskuviin", acNormal, acEdit
End Sub

Private Sub Command9_Click()
DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
DoCmd.OpenQuery "Tag1_CR0-1_FC0-1_F0-2_PWC0_CCM0-1_JB0120", acNormal, acEdit
DoCmd.OpenQuery "Tag1_DCS", acNormal, acEdit
DoCmd.OpenQuery "Tag1_F10-11", acNormal, acEdit
DoCmd.OpenQuery "Tag2_CCM0-1_PWC0", acNormal, acEdit
DoCmd.OpenQuery "Tag1_JB10-12", acNormal, acEdit
DoCmd.OpenQuery "TAG1_FC10-11", acNormal, acEdit
'DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
End Sub

'--------------------------------------------------------------------------------
' Event: Form_Load
' Purpose: Populate KaikkiTaulukot combo box with user table names
'
' Process:
'   1. Iterate through all TableDefs in database
'   2. Exclude system tables (USys*, MSys*, Dev*) and old tables
'   3. Build semicolon-delimited list for combo box
'   4. Set default selection to first table
'
' Notes:
'   - Filters out: USys*, MSys*, Dev*, tables containing "old"
'   - Populates KaikkiTaulukot combo box RowSource
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Current database
    Dim i As Long  ' Loop counter
    Dim Taulukko As String  ' Current table name
    Dim TJoukko As String  ' Semicolon-delimited table list
    
    Set DB = CurrentDb
    For i = 0 To DB.TableDefs.Count - 1
      Taulukko = DB.TableDefs(i).Name
      ' Filter system tables and old tables
      If Left$(Taulukko, 4) <> "USys" And Left$(Taulukko, 4) <> "MSys" And Left$(Taulukko, 3) <> "Dev" And InStr(Taulukko, "old") = 0 Then
        TJoukko = TJoukko & Taulukko & ";"
      End If
    Next i
    
    ' Set combo box row source and default value
    KaikkiTaulukot.RowSource = TJoukko
    KaikkiTaulukot.Value = Left$(TJoukko, InStr(TJoukko, ";") - 1)
    
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbCritical, "Error"
    On Error Resume Next
    Set DB = Nothing
    On Error GoTo 0
End Sub
