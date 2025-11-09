Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub BExit_Click()
  DoCmd.Close
End Sub
'''
' Lista Double Click - Load selected recipient data into form fields
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim SQL As String
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    TName = Nz(taulu.Fields("name"), "")
    TUnit = Nz(taulu.Fields("unit"), "")
    TDescription = Nz(taulu.Fields("description"), "")
    TCompany = Nz(taulu.Fields("company"), "")
    TCopies = taulu.Fields("copies")
    TAddress = Nz(taulu.Fields("address"), "")
    Temail = Nz(taulu.Fields("email"), "")
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.Close
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error loading recipient: " & Err.Description, vbExclamation, "Load Error"
  Resume Cleanup
End Sub
'''
' Save Recipient - Validate and save new recipient to database
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support
'''
Private Sub Valmis_Click()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  ' Validation
  If IsNull(TName) Then
    MsgBox "You have to give name for the recipient!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(TAddress) And IsNull(Temail) Then
    MsgBox "You have to give some address to the recipient!", vbCritical, "Error!"
    Exit Sub
  End If
  
  Set DB = CurrentDb
  DB.BeginTrans
  
  Set taulu = DB.OpenRecordset("USysRecipients", dbOpenTable)
  With taulu
    .AddNew
    If IsNull(TCopies) = False Then
      .Fields("To Distribution") = True
    End If
    .Fields("Name") = TName
    .Fields("Unit") = Nz(TUnit, Null)
    .Fields("Description") = Nz(TDescription, Null)
    .Fields("Company") = Nz(TCompany, Null)
    .Fields("Copies") = TCopies
    .Fields("Address") = Nz(TAddress, Null)
    .Fields("email") = Nz(Temail, Null)
    .Update
  End With
  
  DB.CommitTrans
  
Cleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.Close
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  DoCmd.Close
  Exit Sub
  
ErrorHandler:
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error saving recipient: " & Err.Description, vbCritical, "Save Error"
  Resume Cleanup
End Sub