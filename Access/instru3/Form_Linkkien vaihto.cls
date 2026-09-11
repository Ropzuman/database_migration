Option Compare Database
Option Explicit

Private Const POISSULJETTU_TAULU As String = "MotorData"
Private Const LINKITETTY_TAULU_TYYPPI As Long = 6

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
'      - If paths differ, refresh the link to the current directory
'   4. Preserve table name and structure
'
' Notes:
'   - Only relinks if source directory differs from current directory
'   - Uses case-insensitive path comparison
'   - Prevents duplicate relinking of same table
'   - MotorData table excluded from relinking (special handling needed)
'   - Uses TableDef.RefreshLink to update links without dropping tables
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
    Dim tdf As DAO.TableDef  ' Linkitetyn taulun määritys päivitystä varten
    Dim db As DAO.Database
    Dim Kohdetiedosto As String
    Dim virhenumero As Long
    Dim virhekuvaus As String

    ' Get current database directory path
    Set db = CurrentDb
    Polku = Left$(db.Name, Len(db.Name) - Len(Dir$(db.Name)))
    
    ' Query for all linked tables (excluding MotorData)
    Set Taul = db.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like '" & POISSULJETTU_TAULU & "') AND ([Type] = " & LINKITETTY_TAULU_TYYPPI & "))", dbOpenSnapshot, dbForwardOnly)
    
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
                
                Kohdetiedosto = Polku & Nimi
                If Len(Dir$(Kohdetiedosto)) = 0 Then
                    MsgBox "Tietokantatiedostoa ei löydy:" & vbCrLf & Kohdetiedosto, vbExclamation, "Linkitysvirhe"
                Else
                    ' Päivitetään linkki rikkomatta taulua tai sen relaatioita.
                    Set tdf = db.TableDefs(Taulu)
                    tdf.Connect = ";DATABASE=" & Kohdetiedosto
                    On Error Resume Next
                    tdf.RefreshLink
                    If Err.Number <> 0 Then
                        MsgBox "Virhe päivitettäessä taulua " & tdf.Name & ": " & Err.Description, vbExclamation, "Linkin päivitysvirhe"
                    End If
                    Err.Clear
                    On Error GoTo ErrorHandler
                    Set tdf = Nothing
                End If
            End If
        End If
        Taul.MoveNext
    Loop
    
    ' Cleanup
    Taul.Close
    Set Taul = Nothing
    Set db = Nothing
    MsgBox "Valmis", vbInformation, "Link Update"  ' "Ready"
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    Set tdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    MsgBox "Virhe tauluyhteyksien päivityksessä (" & virhenumero & "): " & virhekuvaus, vbCritical
End Sub
