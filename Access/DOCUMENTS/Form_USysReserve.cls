VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
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
    If IsNull(TPaperSize.Value) Or IsNull(TDiscipline.Value) Or IsNull(TNumber.Value) Then
      MsgBox "You didn't select client number!" & vbCrLf _
           & "Please fill all fields and try again.", vbCritical, "Error!"
      Exit Sub
    End If
    Dim taulu As Recordset
    Dim SQL As String
    SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & TNumber.Value & "'"
    Set taulu = CurrentDb.OpenRecordset(SQL, , dbDenyWrite)
    If IsNull(taulu.Fields("Reserved")) Then 'Tarkistetaan että kyseinen numero on yhä vapaana
      taulu.Edit
      taulu.Fields("Number") = TClientNo.Caption
      taulu.Fields("Reserved") = Varaaja & "@" & Now
      taulu.Update
      taulu.Close
    Else 'Virhe joku kerkesi jo varaamaan numeron välissä
      MsgBox "Someone reserved number " & TNumber.Value & " at the same time!" & vbCrLf & "Please choose another number.", vbCritical, "Error!"
      taulu.Close
      Set taulu = Nothing
      Exit Sub
    End If
    Set taulu = Nothing
    'Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    ClientNo.Value = TClientNo.Caption    'Asiakkaan piirustunumeron osa1: Size
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
  TClientNo.Caption = TPaperSize.Value & "-2090-210-" & TDiscipline.Value & "-" & TNumber.Value
End Sub
