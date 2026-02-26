Option Compare Database
Option Explicit

'================================================================================
' Form: Linkkien vaihto (Link Updater)
' Purpose: Update linked table paths to current database directory
' Updated: 2025-11-11 - Added VBA7/64-bit support
'
' Description:
'   Utility form to relink external tables to the current database directory:
'   - Scans MSysObjects for linked tables
'   - Extracts database filename from current link path
'   - Relinks table to same filename in current database folder
'   - Prevents links to wrong directory after database moves
'   - Excludes MotorData table from relinking
'
' Dependencies:
'   - MSysObjects system table (Access internal)
'   - Linked Access database tables
'
' Use Case:
'   When database is moved to new folder, linked tables still point to old
'   location. This form updates all links to current directory.
'================================================================================

'--------------------------------------------------------------------------------
' Event: Command0_Click
' Purpose: Relink all external tables to current database directory
'
' Process:
'   1. Get current database path (directory only)
'   2. Query MSysObjects for all linked tables (excluding MotorData)
'   3. For each linked table:
'      - Extract filename from current link path
'      - Compare directory paths
'      - If paths differ, drop and recreate link to current directory
'   4. Preserve table name and structure
'
' Notes:
'   - Only relinks if source directory differs from current directory
'   - Uses case-insensitive path comparison
'   - Prevents duplicate relinking of same table
'   - MotorData table excluded from relinking (special handling needed)
'   - Uses DoCmd.TransferDatabase to recreate links
'--------------------------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer  ' Loop counters
    Dim Polku As String  ' Current database directory path
    Dim Polku2 As String  ' Linked table directory path
    Dim Nimi As String  ' Database filename
    Dim Taulut() As String  ' Array for table tracking (unused)
    Dim s As Integer  ' Counter (unused)
    Dim Taul As DAO.Recordset  ' MSysObjects query results
    Dim Taulu As String  ' Current table name being processed

    ' Get current database directory path
    Polku = Left$(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name)))
    
    ' Query for all linked tables (excluding MotorData)
    Set Taul = CurrentDb.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like 'MotorData'))")
    
    ' Process each linked table
    Do While Not Taul.EOF
        ' Prevent processing same table multiple times
        If Taul.Fields("Name") <> Taulu Then
            ' Extract directory and filename from current link path
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
               Polku2 = LCase$(Left$(Taul.Fields("Database"), i))  ' Directory path
                    Nimi = Mid$(Taul.Fields("Database"), i + 1)  ' Filename
                    Exit For
                End If
            Next i
            
            ' Check if link path differs from current database path
            If Polku2 <> LCase$(Polku) Then
                Taulu = Taul.Fields("Name")
                
                ' Drop existing linked table
                CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
                
                ' Recreate link to current directory
                On Error Resume Next
                DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
                Err.Clear
                On Error GoTo ErrorHandler
            End If
        End If
        Taul.MoveNext
    Loop
    
    ' Cleanup
    Taul.Close
    Set Taul = Nothing
    MsgBox "Valmis", vbInformation, "Link Update"  ' "Ready"
    Exit Sub

ErrorHandler:
    MsgBox "Error updating links: " & Err.Description, vbCritical, "Link Update"
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    On Error GoTo 0
End Sub
