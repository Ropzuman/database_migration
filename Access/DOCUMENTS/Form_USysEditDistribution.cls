Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub BExit_Click()
  DoCmd.Close
End Sub

'''
' Form Load Event - Load distribution details from database
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT * FROM USySDISTRIBUTION WHERE ID=" & Common, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    DistrNo = Nz(taulu.Fields("No"), "")
    TDate = taulu.Fields("Date")
    TReason = Nz(taulu.Fields("Reason"), "")
    Desc = Nz(taulu.Fields("Description"), "")
    TSender = Nz(taulu.Fields("Sender"), "")
    Note = Nz(taulu.Fields("Note"), "")
  End If
  
  Vastaanottajat.RowSource = "SELECT Counter, * FROM USysRecipByDistr WHERE DistrID=" & Common
  
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
  MsgBox "Error loading distribution: " & Err.Description, vbExclamation, "Form Load Error"
  Resume Cleanup
End Sub
Private Sub Lisaa_Click()
'Lisää uusi vastaanottaja
   DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Lista.Requery
End Sub
'''
' Lista Double Click - Add selected recipient to distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support
'''
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim Taulu2 As DAO.Recordset
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  
  Oletus = IIf(IsNull(taulu.Fields("Copies")), "", taulu.Fields("Copies"))
  Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are usually marked as ""4+e"".", "Copies", Oletus)
  
  If Kopioita <> "" Then
    DB.BeginTrans
    Set Taulu2 = DB.OpenRecordset("USysRecipByDistr", dbOpenTable)
    With Taulu2
      .AddNew
      .Fields("DistrID") = Common
      .Fields("Copies") = Kopioita
      .Fields("Unit") = Nz(taulu.Fields("Unit"), "")
      .Fields("Company") = Nz(taulu.Fields("Company"), "")
      .Fields("Name") = Nz(taulu.Fields("Name"), "")
      .Fields("Description") = Nz(taulu.Fields("Description"), "")
      .Fields("Address") = Nz(taulu.Fields("Address"), "")
      .Fields("email") = Nz(taulu.Fields("email"), "")
      .Update
    End With
    DB.CommitTrans
  End If
  
  Vastaanottajat.Requery
  
Cleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.Close
    Set taulu = Nothing
  End If
  If Not Taulu2 Is Nothing Then
    If Taulu2.State = 1 Then Taulu2.Close
    Set Taulu2 = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error adding recipient: " & Err.Description, vbExclamation, "Add Error"
  Resume Cleanup
End Sub

'''
' Vastaanottajat Double Click - Remove recipient from distribution
' OPTIMIZED 2025-11-09: Added explicit DAO typing, error handling
'''
Private Sub Vastaanottajat_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  DB.Execute "DELETE * FROM USysRecipByDistr WHERE (((USysRecipByDistr.Counter)=" & Vastaanottajat.ItemData(Vastaanottajat.ItemsSelected(0)) & "));"
  Vastaanottajat.Requery
  
  DB.Close
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  MsgBox "Error removing recipient: " & Err.Description, vbExclamation, "Delete Error"
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
End Sub

'''
' Valmis Click - Update distribution details
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support
'''
Private Sub Valmis_Click()
  Dim SQL As String
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  ' Validation
  If IsNull(TDate) Then
    MsgBox "You have to select date!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Desc) Then
    MsgBox "You have to put description!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(TReason) Then
    MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
    Exit Sub
  End If
  
  Set DB = CurrentDb
  DB.BeginTrans
  
  SQL = "SELECT * FROM UsysDISTRIBUTION WHERE ID=" & Common
  Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset)
  
  With taulu
    .Edit
    .Fields("Description") = Desc
    .Fields("Date") = TDate
    .Fields("Sender") = TSender
    .Fields("Reason") = TReason
    .Fields("Note") = Nz(Note, Null)
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
  MsgBox "Error updating distribution: " & Err.Description, vbCritical, "Update Error"
  Resume Cleanup
End Sub