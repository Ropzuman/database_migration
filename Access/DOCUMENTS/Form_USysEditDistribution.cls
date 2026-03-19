Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysEditDistribution
' SOVELLUS : DOCUMENTS — Distribuution muokkaus
' KUVAUS   : Lataa olemassa olevan distribuution tiedot muokkauslomakkeelle
'            ja mahdollistaa vastaanottajien lisäyksen tai poiston.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private mDistrID As Long

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

  ' Distribuution tunniste luetaan avaavan lomakkeen OpenArgs-parametrista.
  If IsNull(Me.OpenArgs) Or Len(Trim$(CStr(Me.OpenArgs))) = 0 Then
    MsgBox "Distribuution tunniste puuttuu.", vbCritical, "Muokkaa distribuutiota"
    DoCmd.Close acForm, Me.Name
    Exit Sub
  End If
  If Not IsNumeric(Me.OpenArgs) Then
    MsgBox "Virheellinen distribuution tunniste.", vbCritical, "Muokkaa distribuutiota"
    DoCmd.Close acForm, Me.Name
    Exit Sub
  End If
  mDistrID = CLng(Me.OpenArgs)
  
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT * FROM USySDISTRIBUTION WHERE ID=" & mDistrID, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    Me.DistrNo = Nz(taulu.Fields("No"), "")
    Me.TDate = taulu.Fields("Date")
    Me.TReason = Nz(taulu.Fields("Reason"), "")
    Me.Desc = Nz(taulu.Fields("Description"), "")
    Me.TSender = Nz(taulu.Fields("Sender"), "")
    Me.Note = Nz(taulu.Fields("Note"), "")
  End If

  Me.Vastaanottajat.RowSource = "SELECT Counter, * FROM USysRecipByDistr WHERE DistrID=" & mDistrID
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
  MsgBox "Error loading distribution: " & Err.Description, vbExclamation, "Form Load Error"
  Resume Cleanup
End Sub
Private Sub Lisaa_Click()
'Lisää uusi vastaanottaja
   DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Me.Lista.Requery
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
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Me.Lista.ItemData(Me.Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  
  Oletus = IIf(IsNull(taulu.Fields("Copies")), "", taulu.Fields("Copies"))
  Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are usually marked as ""4+e"".", "Copies", Oletus)
  
  If Kopioita <> "" Then
    DBEngine.BeginTrans
    Set Taulu2 = DB.OpenRecordset("USysRecipByDistr", dbOpenTable)
    With Taulu2
      .AddNew
      .Fields("DistrID") = mDistrID
      .Fields("Copies") = Kopioita
      .Fields("Unit") = Nz(taulu.Fields("Unit"), "")
      .Fields("Company") = Nz(taulu.Fields("Company"), "")
      .Fields("Name") = Nz(taulu.Fields("Name"), "")
      .Fields("Description") = Nz(taulu.Fields("Description"), "")
      .Fields("Address") = Nz(taulu.Fields("Address"), "")
      .Fields("email") = Nz(taulu.Fields("email"), "")
      .Update
    End With
    DBEngine.CommitTrans
  End If
  
  Me.Vastaanottajat.Requery
  
Cleanup:
  If Not taulu Is Nothing Then
    Set taulu = Nothing
  End If
  If Not Taulu2 Is Nothing Then
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
    DBEngine.Rollback
  End If
  MsgBox "Error adding recipient: " & Err.Description, vbExclamation, "Add Error"
End Sub

' IsNumeric() — using VBA built-in; private override removed to avoid compile ambiguity.

'''
' Vastaanottajat Double Click - Remove recipient from distribution
' OPTIMIZED 2025-11-09: Added explicit DAO typing, error handling
'''
Private Sub Vastaanottajat_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  DB.Execute "DELETE * FROM USysRecipByDistr WHERE (((USysRecipByDistr.Counter)=" & Me.Vastaanottajat.ItemData(Me.Vastaanottajat.ItemsSelected(0)) & "));"
  Me.Vastaanottajat.Requery
  
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
  If IsNull(Me.TDate) Then
    MsgBox "You have to select date!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.Desc) Then
    MsgBox "You have to put description!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.TReason) Then
    MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
    Exit Sub
  End If
  
  Set DB = CurrentDb
  DBEngine.BeginTrans
  
  SQL = "SELECT * FROM UsysDISTRIBUTION WHERE ID=" & mDistrID
  Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset)
  
  With taulu
    .Edit
    .Fields("Description") = Me.Desc
    .Fields("Date") = Me.TDate
    .Fields("Sender") = Me.TSender
    .Fields("Reason") = Me.TReason
    .Fields("Note") = Nz(Me.Note, Null)
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
  MsgBox "Error updating distribution: " & Err.Description, vbCritical, "Update Error"
  Resume Cleanup
End Sub