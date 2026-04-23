Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysNewDistribution
' SOVELLUS : DOCUMENTS — Uuden distribuution luonti
' KUVAUS   : Käyttäjä valitsee vastaanottajat ja syöttää distribuution
'            perustiedot. Tallennuksessa luodaan merkinnät tauluihin
'            UsysDISTRIBUTION ja USysRecipByDistr.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub BExit_Click()
  DoCmd.Close
End Sub

'==========================================================================
' Tapahtuma: Form_Load
' KUVAUS  : Alustaa uuden distribuution lomakkeen tilan.
'           Nollaa vain nykyisen käyttäjän aiemmat valinnat.
'==========================================================================
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim Kayttaja As String

  On Error GoTo ErrorHandler

  ' Haetaan nykyinen Windows-käyttäjätunnus omien valintojen tunnistamiseksi
  Kayttaja = GetNetworkUserName()

  Set DB = CurrentDb
  ' Nollataan VAIN tämän käyttäjän aiemmat unohtuneet valinnat —
  ' muiden käyttäjien samanaikaiset valinnat säilyvät koskemattomina
  DB.Execute "UPDATE USysRecipients SET [To Distribution] = No, SelectedBy = Null " & _
             "WHERE SelectedBy = '" & Replace(Kayttaja, "'", "") & "';", dbFailOnError
  Me.TDate.Value = Date

  ' EI DB.Close: CurrentDb-viitettä ei tule sulkea koodista
  Set DB = Nothing
  Exit Sub

ErrorHandler:
  MsgBox "Error initializing form: " & Err.Description, vbExclamation, "Form Load Error"
  If Not DB Is Nothing Then
    Set DB = Nothing
  End If
End Sub

Private Sub Lisaa_Click()
' Avataan uuden vastaanottajan lisäyslomake
   DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Me.Lista.Requery
End Sub

'==========================================================================
' Tapahtuma: Lista_DblClick
' KUVAUS  : Lisää tai poistaa vastaanottajan käyttäjän nykyisestä
'           distribuutiovalinnasta.
'==========================================================================
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
SQL = "SELECT * FROM USysRecipients WHERE ID=" & Me.Lista.ItemData(Me.Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset)

  If Not taulu.EOF Then
    taulu.Edit
    If taulu.Fields("To Distribution") Then
      ' Poistetaan valinta — tyhjennetään myös käyttäjätunnus
      taulu.Fields("To Distribution") = False
      taulu.Fields("SelectedBy") = Null
    Else
      Oletus = IIf(IsNull(taulu.Fields("Copies")), "", taulu.Fields("Copies"))
      Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are marked as ""4+e"".", "Copies", Oletus)
      If Kopioita <> "" Then
        ' Merkitään valituksi ja kirjataan valitsijan käyttäjätunnus monikäyttäjätuen takaamiseksi
        taulu.Fields("To Distribution") = True
        taulu.Fields("Copies") = Kopioita
        taulu.Fields("SelectedBy") = GetNetworkUserName()
      End If
    End If
    taulu.Update
  End If

  Me.Lista.Requery

Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  ' EI DB.Close: CurrentDb-viitettä ei tule sulkea koodista
  Set DB = Nothing
  Exit Sub

ErrorHandler:
  MsgBox "Error updating recipient: " & Err.Description, vbExclamation, "Update Error"
  Resume Cleanup
End Sub

Private Sub Valmis_Click()
' Lisätään uusi distribuutio UsysDISTRIBUTION-tauluun ja
' vastaanottajat USysRecipByDistr-tauluun transaktion sisällä atomaarisuuden takaamiseksi
  Dim SQL As String
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim Taulu2 As DAO.Recordset
  Dim DistrID As Long
  Dim inTransaction As Boolean
  Dim Kayttaja As String

  On Error GoTo ErrorHandler
  inTransaction = False

  Set DB = CurrentDb
  Kayttaja = GetNetworkUserName()

  ' Tarkistetaan valinnat vain tämän käyttäjän rivien perusteella
  Set taulu = DB.OpenRecordset("SELECT * FROM USysRecipients WHERE [To Distribution] = Yes AND SelectedBy = '" & Replace(Kayttaja, "'", "") & "';")
  If taulu.EOF Then
    MsgBox "You have to select recipients for Distribution!", vbCritical, "Error!"
    GoTo Cleanup
  ElseIf IsNull(Me.DistrNo) Then
    MsgBox "You have to give number for distribution!", vbCritical, "Error!"
    GoTo Cleanup
  ElseIf IsNull(Me.TDate) Then
    MsgBox "You have to select date!", vbCritical, "Error!"
    GoTo Cleanup
  ElseIf IsNull(Me.Desc) Then
    MsgBox "You have to put description!", vbCritical, "Error!"
    GoTo Cleanup
  ElseIf IsNull(Me.TReason) Then
    MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
    GoTo Cleanup
  End If

  taulu.Close
  Set taulu = Nothing

  SQL = "SELECT * FROM UsysDISTRIBUTION WHERE No='" & Replace(Me.DistrNo, "'", "") & "'"
  Set taulu = DB.OpenRecordset(SQL)
  If Not taulu.EOF Then
    MsgBox "Distribuution numero """ & Me.DistrNo & """ on jo käytössä! Valitse toinen numero.", vbCritical, "Error!"
    GoTo Cleanup
  End If
  taulu.Close
  Set taulu = Nothing

  ' Aloitetaan transaktio: molemmat taulut päivitetään atomaarisesti
  DBEngine.BeginTrans
  inTransaction = True

  Set taulu = DB.OpenRecordset("UsysDISTRIBUTION", dbOpenDynaset)
  With taulu
    .AddNew
    .Fields("No") = Me.DistrNo
    .Fields("Description") = Me.Desc
    .Fields("Date") = Me.TDate
    .Fields("Sender") = Me.TSender
    .Fields("Reason") = Me.TReason
    .Fields("Note") = Me.Note
    DistrID = .Fields("ID")
    .Update
    .Close
  End With
  Set taulu = Nothing

  ' Tehokkaampi INSERT INTO -kysely — kopioidaan vain tämän käyttäjän valitsemat vastaanottajat
  SQL = "INSERT INTO USysRecipByDistr (DistrID, Copies, Unit, Company, Name, Description, Address, email) " & _
        "SELECT " & DistrID & ", Copies, Unit, Company, Name, Description, Address, email " & _
        "FROM USysRecipients WHERE [To Distribution] = Yes AND SelectedBy = '" & Replace(Kayttaja, "'", "") & "';"
  DB.Execute SQL, dbFailOnError

  ' Siivotaan tämän käyttäjän valinnat lomakkeen sulkemisen jälkeen
  DB.Execute "UPDATE USysRecipients SET [To Distribution] = No, SelectedBy = Null " & _
             "WHERE SelectedBy = '" & Replace(Kayttaja, "'", "") & "';", dbFailOnError

  DBEngine.CommitTrans
  inTransaction = False
  DoCmd.Close

Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  If Not Taulu2 Is Nothing Then
    Taulu2.Close
    Set Taulu2 = Nothing
  End If
  ' EI DB.Close: CurrentDb-viitettä ei tule sulkea koodista
  Set DB = Nothing
  Exit Sub

ErrorHandler:
  If inTransaction Then
    DBEngine.Rollback  ' Perutaan kaikki muutokset virhetilanteessa
    inTransaction = False
  End If
  MsgBox "Error saving distribution: " & Err.Description, vbCritical, "Database Error"
  Resume Cleanup
End Sub

Private Function GetNetworkUserName() As String
  ' Luetaan käyttäjänimi suoraan ympäristömuuttujasta, jotta lomake ei riipu ulkoisesta moduulista.
  On Error GoTo FallbackHandler
  GetNetworkUserName = Trim$(CStr(Environ$("USERNAME")))
  If GetNetworkUserName = "" Then GoTo FallbackHandler
  Exit Function

FallbackHandler:
  GetNetworkUserName = "Unknown"
End Function

' IsNumeric() käyttää VBA:n sisäänrakennettua toteutusta.
' Yksityinen ylikirjoitus poistettiin käännösepäselvyyden estämiseksi.