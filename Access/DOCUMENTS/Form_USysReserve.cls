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
    'Formin tietojen talletus ja sulkeminen
    If IsNull(Me.TPaperSize.Value) Or IsNull(Me.TDiscipline.Value) Or IsNull(Me.TNumber.Value) Then
      MsgBox "You didn't select client number!" & vbCrLf _
           & "Please fill all fields and try again.", vbCritical, "Error!"
      Exit Sub
    End If
    Dim taulu As DAO.Recordset
    Dim SQL As String
    SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & Me.TNumber.Value & "'"
    Set taulu = CurrentDb.OpenRecordset(SQL, , dbDenyWrite)
    If IsNull(taulu.Fields("Reserved")) Then 'Tarkistetaan että kyseinen numero on vielä vapaana
      taulu.Edit
      taulu.Fields("Number") = Me.TClientNo.Caption
      taulu.Fields("Reserved") = Varaaja & "@" & Now
      taulu.Update
      taulu.Close
    Else 'Virhe joku kerkesi jo varaamaan numeron välissä
      MsgBox "Someone reserved number " & Me.TNumber.Value & " at the same time!" & vbCrLf & "Please choose another number.", vbCritical, "Error!"
      taulu.Close
      Set taulu = Nothing
      Exit Sub
    End If
    Set taulu = Nothing
    'Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    Me.ClientNo.Value = Me.TClientNo.Caption    'Asiakkaan piirustunumeron osa1: Size
    DoCmd.Close
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