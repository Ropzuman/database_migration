Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub BExit_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
Dim DB As Database
Dim taulu As Recordset
Set DB = CurrentDb
Set taulu = DB.OpenRecordset("SELECT * FROM USySDISTRIBUTION WHERE ID=" & Common)
With taulu
  DistrNo = .Fields("No")
'  MUnit = .Fields("MetsoUnit")
  TDate = .Fields("Date")
  TReason = .Fields("Reason")
  Desc = .Fields("Description")
  TSender = .Fields("Sender")
  Note = .Fields("Note")
End With
Vastaanottajat.RowSource = "SELECT Counter, * FROM USysRecipByDistr WHERE DistrID=" & Common
End Sub
Private Sub Lisaa_Click()
'Lis�� uusi vastaanottaja
   DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Lista.Requery
End Sub
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As Database
  Dim taulu As Recordset
  Dim Taulu2 As Recordset
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String
  
  Set DB = CurrentDb
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL)
  Set Taulu2 = DB.OpenRecordset("USysRecipByDistr")
  Oletus = IIf(IsNull(taulu.Fields("Copies")), "", taulu.Fields("Copies"))
  Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are usually marked as ""4+e"".", "Copies", Oletus)
  If Kopioita <> "" Then
    With Taulu2
      .AddNew
      .Fields("DistrID") = Common
      .Fields("Copies") = Kopioita
      .Fields("Unit") = taulu.Fields("Unit")
      .Fields("Company") = taulu.Fields("Company")
      .Fields("Name") = taulu.Fields("Name")
      .Fields("Description") = taulu.Fields("Description")
      .Fields("Address") = taulu.Fields("Address")
      .Fields("email") = taulu.Fields("email")
      .Update
      .Close
    End With
  End If
  Vastaanottajat.Requery
  Set DB = Nothing
  Set taulu = Nothing
  Set Taulu2 = Nothing
End Sub
Private Sub Vastaanottajat_DblClick(Cancel As Integer)
  CurrentDb.Execute "DELETE * FROM USysRecipByDistr WHERE (((USysRecipByDistr.Counter)=" & Vastaanottajat.ItemData(Vastaanottajat.ItemsSelected(0)) & "));"
  Vastaanottajat.Requery
End Sub
Private Sub Valmis_Click()
'T�m� lis�� uuden distribuution USysDISTRIBUTION taulukkoon ja
'Vastaanottajat USysRecipByDistr taulukkoon
Dim SQL As String
Dim DB As Database
Dim taulu As Recordset
If IsNull(TDate) Then
 MsgBox "You have to select date!", vbCritical, "Error!"
ElseIf IsNull(Desc) Then
 MsgBox "You have to put description!", vbCritical, "Error!"
ElseIf IsNull(TReason) Then
 MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
Else
  Set DB = CurrentDb
  SQL = "SELECT * FROM UsysDISTRIBUTION WHERE ID=" & Common
  Set taulu = DB.OpenRecordset(SQL)
  With taulu
    .Edit
    .Fields("Description") = Desc
    .Fields("Date") = TDate
    .Fields("Sender") = TSender
    .Fields("Reason") = TReason
    .Fields("Note") = Note
    .Update
    .Close
  End With
  Set taulu = Nothing
  Set DB = Nothing
  DoCmd.Close
End If
End Sub