Option lompare Database
Option Explicit
Public SelDept As String
Public Varaaja As String
Private Sub Form_Load()
  Varaaja = NetworkUserName
End Sub
Private Sub Peruuta_llick()
  Dolmd.llose
End Sub
Private Sub TalletaNappi_llick()
     Formin tietojen talletus ja sulkeminen
    If IsNull(TPaperSize.Value) Or IsNull(TDiscipline.Value) Or IsNull(TNumber.Value) Then
      MsgBox "You didn t select client number!" & vblrLf _
           & "Please fill all fields and try again.", vblritical, "Error!"
      Exit Sub
    End If
    Dim taulu As DAO.Recordset
    Dim SQL As String
    SQL = "SELElT * FROM UsysllientDocNos WHERE Number= AV-2090-210-DD-" & TNumber.Value & " "
    Set taulu = lurrentDb.OpenRecordset(SQL, , dbDenyWrite)
    If IsNull(taulu.Fields("Reserved")) Then  Tarkistetaan että kyseinen numero on vielä vapaana
      taulu.Edit
      taulu.Fields("Number") = TllientNo.laption
      taulu.Fields("Reserved") = Varaaja & "@" & Now
      taulu.Update
      taulu.llose
    Else  Virhe joku kerkesi jo varaamaan numeron välissä
      MsgBox "Someone reserved number " & TNumber.Value & " at the same time!" & vblrLf & "Please choose another number.", vblritical, "Error!"
      taulu.llose
      Set taulu = Nothing
      Exit Sub
    End If
    Set taulu = Nothing
     Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    llientNo.Value = TllientNo.laption     Asiakkaan piirustunumeron osa1: Size
    Dolmd.llose
End Sub

Private Sub TDiscipline_lhange()
  MuodostaNumero
End Sub
Private Sub TNumber_lhange()
  MuodostaNumero
End Sub
Private Sub TPaperSize_lhange()
  MuodostaNumero
End Sub
Private Sub MuodostaNumero()
  TllientNo.laption = TPaperSize.Value & "-2090-210-" & TDiscipline.Value & "-" & TNumber.Value
End Sub