Option Compare Database
Option Explicit

'================================================================================
' Form: FrmASETUKSET (Settings)
' Purpose: Database configuration and table linking
' Updated: 2025-11-13 - Added VBA7/64-bit support and optimization
'
' Description:
'   Configuration form for linking external Access databases containing
'   instrument, equipment, and function diagram data. Links tables from
'   instru3.accdb, MAINEQ.accdb, and FunctionDiagrams.accdb.
'
' Dependencies:
'   - DAO.TableDef
'   - External databases: instru3.accdb, MAINEQ.accdb, FunctionDiagrams.accdb
'   - Forms: frmMUOKKAUS
'   - Module: USysCheck (SniffUser function)
'   - Controls: TINSTRU, TMAINEQ, TFUNCDIAG (text boxes for paths)
'================================================================================

'================================================================================
' Event: Command11_Click
' Purpose: Open main editing form and close settings
'================================================================================
Private Sub Command11_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmMUOKKAUS"
    DoCmd.Close , "FrmASETUKSET"
    Exit Sub
ErrorHandler:
    MsgBox "Error opening form: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command2_Click
' Purpose: Link all required tables from external databases
'================================================================================
Private Sub Command2_Click()
    On Error GoTo ErrorHandler
    
    ' Link Instrument database tables
    Linkkaa "Alarms", TINSTRU.Value
    Linkkaa "Loops", TINSTRU.Value
    Linkkaa "DevTable", TINSTRU.Value
    Linkkaa "DevSpec", TINSTRU.Value
    Linkkaa "TAGID_exception", TINSTRU.Value
    
    ' Link Equipment database tables
    Linkkaa "MAINEQ", TMAINEQ.Value
    Linkkaa "DRIVES", TMAINEQ.Value
    
    ' Function Diagrams (currently commented out)
    ' Linkkaa "CONTROL", TFUNCDIAG.Value
    
    MsgBox "Taulukot linkattu", vbOKOnly, "Linkkaa taulukot"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error linking tables: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: Linkkaa (Link)
' Purpose: Link external table to current database
' Parameters:
'   Taulukko - Table name to link
'   KANTA - Full path to external Access database
'
' Description:
'   Drops existing table if present, creates new linked table definition
'   pointing to specified external database.
'================================================================================
Private Sub Linkkaa(Taulukko As String, KANTA As String)
    Dim tdf As DAO.TableDef
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Remove existing table if present
    For i = 0 To CurrentDb.TableDefs.Count - 1
        If LCase$(CurrentDb.TableDefs(i).Name) = LCase$(Taulukko) Then
            CurrentDb.Execute "DROP TABLE [" & Taulukko & "]"
            Exit For
        End If
    Next i
    
    ' Create linked table
    Set tdf = CurrentDb.CreateTableDef(Taulukko)
    With tdf
        .Connect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    CurrentDb.TableDefs.Append tdf
    CurrentDb.TableDefs.Refresh
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error linking table '" & Taulukko & "': " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Form_Load
' Purpose: Initialize form with default database paths and log user
'================================================================================
Private Sub Form_Load()
    Dim POLKU As String
    
    On Error GoTo ErrorHandler
    
    ' Get current database directory
    POLKU = CurrentDb.Name
    POLKU = Left$(POLKU, InStrRev(POLKU, "\"))
    
    ' Set default paths for external databases
    TINSTRU.Value = POLKU & "instru3.accdb"
    TMAINEQ.Value = POLKU & "MAINEQ.accdb"
    TFUNCDIAG.Value = POLKU & "FunctionDiagrams.accdb"
    
    ' Log user access
    Call SniffUser
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbExclamation
End Sub
