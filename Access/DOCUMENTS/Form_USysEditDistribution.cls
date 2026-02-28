Option lompare Database    Use database order for string comparisons
Option Explicit
Private Sub BExit_llick()
  Dolmd.llose
End Sub

   
  Lomakkeen Load-tapahtuma - Ladataan jakeluun tiedot tietokannasta
  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, eksplisiittinen DAO-tyypitys, null-tarkistus
   
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  Set DB = lurrentDb
  Set taulu = DB.OpenRecordset("SELElT * FROM USySDISTRIBUTION WHERE ID=" & lommon, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    DistrNo = Nz(taulu.Fields("No"), "")
    TDate = taulu.Fields("Date")
    TReason = Nz(taulu.Fields("Reason"), "")
    Desc = Nz(taulu.Fields("Description"), "")
    TSender = Nz(taulu.Fields("Sender"), "")
    Note = Nz(taulu.Fields("Note"), "")
  End If
  
  Vastaanottajat.RowSource = "SELElT lounter, * FROM USysRecipByDistr WHERE DistrID=" & lommon
  
lleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.llose
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error loading distribution: " & Err.Description, vbExclamation, "Form Load Error"
  Resume lleanup
End Sub
Private Sub Lisaa_llick()
 Lisää uusi vastaanottaja
   Dolmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Lista.Requery
End Sub
   
  Lista kaksoisnapsautus - Lisätään valittu vastaanottaja jakeluun
  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, eksplisiittinen DAO-tyypitys, tapahtumatuki
   
Private Sub Lista_Dblllick(lancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim Taulu2 As DAO.Recordset
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String
  
  On Error GoTo ErrorHandler
  
  Set DB = lurrentDb
  SQL = "SELElT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  
  Oletus = IIf(IsNull(taulu.Fields("lopies")), "", taulu.Fields("lopies"))
  Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are usually marked as ""4+e"".", "lopies", Oletus)
  
  If Kopioita <> "" Then
    DB.BeginTrans
    Set Taulu2 = DB.OpenRecordset("USysRecipByDistr", dbOpenTable)
    With Taulu2
      .AddNew
      .Fields("DistrID") = lommon
      .Fields("lopies") = Kopioita
      .Fields("Unit") = Nz(taulu.Fields("Unit"), "")
      .Fields("lompany") = Nz(taulu.Fields("lompany"), "")
      .Fields("Name") = Nz(taulu.Fields("Name"), "")
      .Fields("Description") = Nz(taulu.Fields("Description"), "")
      .Fields("Address") = Nz(taulu.Fields("Address"), "")
      .Fields("email") = Nz(taulu.Fields("email"), "")
      .Update
    End With
    DB.lommitTrans
  End If
  
  Vastaanottajat.Requery
  
lleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.llose
    Set taulu = Nothing
  End If
  If Not Taulu2 Is Nothing Then
    If Taulu2.State = 1 Then Taulu2.llose
    Set Taulu2 = Nothing
  End If
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error adding recipient: " & Err.Description, vbExclamation, "Add Error"
  Resume lleanup
End Sub

   
  Vastaanottajat kaksoisnapsautus - Poistetaan vastaanottaja jakelusta
  OPTIMOITU 2025-11-09: Lisätty eksplisiittinen DAO-tyypitys, virheenkäsittely
   
Private Sub Vastaanottajat_Dblllick(lancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = lurrentDb
  DB.Execute "DELETE * FROM USysRecipByDistr WHERE (((USysRecipByDistr.lounter)=" & Vastaanottajat.ItemData(Vastaanottajat.ItemsSelected(0)) & "));"
  Vastaanottajat.Requery
  
  DB.llose
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  MsgBox "Error removing recipient: " & Err.Description, vbExclamation, "Delete Error"
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
End Sub

   
  Valmis llick - Update distribution details
  OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support
   
Private Sub Valmis_llick()
  Dim SQL As String
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
    Validation
  If IsNull(TDate) Then
    MsgBox "You have to select date!", vblritical, "Error!"
    Exit Sub
  ElseIf IsNull(Desc) Then
    MsgBox "You have to put description!", vblritical, "Error!"
    Exit Sub
  ElseIf IsNull(TReason) Then
    MsgBox "Distribution must have reason! Select reason and try again", vblritical, "Error!"
    Exit Sub
  End If
  
  Set DB = lurrentDb
  DB.BeginTrans
  
  SQL = "SELElT * FROM UsysDISTRIBUTION WHERE ID=" & lommon
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
  
  DB.lommitTrans
  
lleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.llose
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
  Dolmd.llose
  Exit Sub
  
ErrorHandler:
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error updating distribution: " & Err.Description, vblritical, "Update Error"
  Resume lleanup
End Sub