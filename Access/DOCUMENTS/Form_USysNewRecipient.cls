Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysNewRecipient
' SOVELLUS : DOCUMENTS — Uuden vastaanottajan lisäys
' KUVAUS   : Lisää tai hakee vastaanottajatietoja USysRecipients-tauluun.
'            Lista_DblClick lataa valitun vastaanottajan tiedot kenttiin.
' PÄIVITETTY: 2026-03-03
'==========================================================================
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
SQL = "SELECT * FROM USysRecipients WHERE ID=" & Me.Lista.ItemData(Me.Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)

  If Not taulu.EOF Then
    Me.TName = Nz(taulu.Fields("name"), "")
    Me.TUnit = Nz(taulu.Fields("unit"), "")
    Me.TDescription = Nz(taulu.Fields("description"), "")
    Me.TCompany = Nz(taulu.Fields("company"), "")
    Me.TCopies = taulu.Fields("copies")
    Me.TAddress = Nz(taulu.Fields("address"), "")
    Me.Temail = Nz(taulu.Fields("email"), "")
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
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
  If IsNull(Me.TName) Then
    MsgBox "You have to give name for the recipient!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.TAddress) And IsNull(Me.Temail) Then
    MsgBox "You have to give some address to the recipient!", vbCritical, "Error!"
    Exit Sub
  End If
  
  Set DB = CurrentDb
  DBEngine.BeginTrans
  
  Set taulu = DB.OpenRecordset("USysRecipients", dbOpenTable)
  With taulu
    .AddNew
    If IsNull(Me.TCopies) = False Then
      .Fields("To Distribution") = True
    End If
    .Fields("Name") = Me.TName
    .Fields("Unit") = Nz(Me.TUnit, Null)
    .Fields("Description") = Nz(Me.TDescription, Null)
    .Fields("Company") = Nz(Me.TCompany, Null)
    .Fields("Copies") = Me.TCopies
    .Fields("Address") = Nz(Me.TAddress, Null)
    .Fields("email") = Nz(Me.Temail, Null)
    .Update
  End With
  
  DBEngine.CommitTrans
  
Cleanup:
  If Not taulu Is Nothing Then
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
    DBEngine.Rollback
  End If
  MsgBox "Error saving recipient: " & Err.Description, vbCritical, "Save Error"
  Resume Cleanup
End Sub