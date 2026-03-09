Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysReserve
' SOVELLUS : DOCUMENTS — Asiakkaan piirustunumeron varaus
' KUVAUS   : Varaa asiakkaan piirustunumeron UsysClientDocNos-tauluun.
'            Tallennuksessa kirjataan varaajan nimi ja aikaleima.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Public SelDept As String
Public Varaaja As String
Private Sub Form_Load()
  Varaaja = NetworkUserName
End Sub
Private Sub Peruuta_Click()
  DoCmd.Close
End Sub
Private Sub TalletaNappi_Click()
    ' Formin tietojen talletus ja sulkeminen asiakkaan piirustunumeron varauksen jälkeen
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    Dim EtsittavaNumero As String

    On Error GoTo ErrorHandler

    If IsNull(Me.TPaperSize.Value) Or IsNull(Me.TDiscipline.Value) Or IsNull(Me.TNumber.Value) Then
      MsgBox "You didn't select client number!" & vbCrLf _
           & "Please fill all fields and try again.", vbCritical, "Error!"
      Exit Sub
    End If

    Set DB = CurrentDb
    ' Puhdistetaan heittomerkit SQL-injektioriskin poistamiseksi
    EtsittavaNumero = Replace(Me.TNumber.Value, "'", "")

    SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & EtsittavaNumero & "'"
    Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset, dbDenyWrite)

    If IsNull(taulu.Fields("Reserved")) Then
      ' Numero vielä vapaa — kirjataan varaus ja aikaleima
      taulu.Edit
      taulu.Fields("Number") = Me.TClientNo.Caption
      taulu.Fields("Reserved") = Varaaja & "@" & Now
      taulu.Update
      Me.ClientNo.Value = Me.TClientNo.Caption
      DoCmd.Close
    Else
      ' Kilpailutilanne: joku kerkesi varaamaan numeron samanaikaisesti
      MsgBox "Someone reserved number " & EtsittavaNumero & " at the same time!" & vbCrLf & _
             "Please choose another number.", vbCritical, "Error!"
    End If

Cleanup:
    If Not taulu Is Nothing Then
      taulu.Close
      Set taulu = Nothing
    End If
    ' EI DB.Close: CurrentDb-viitettä ei tule sulkea koodista
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Database Error"
    Resume Cleanup
End Sub

Private Sub TDiscipline_Change()
  MuodostaNumero
End Sub
Private Sub TNumber_Change()
  MuodostaNumero
End Sub
Private Sub TPaperSize_Change()
  MuodostaNumero
End Sub
Private Sub MuodostaNumero()
  ' Muodostetaan asiakkaan piirustunumero valituista kenttien arvoista
  Me.TClientNo.Caption = Me.TPaperSize.Value & "-2090-210-" & Me.TDiscipline.Value & "-" & Me.TNumber.Value
End Sub