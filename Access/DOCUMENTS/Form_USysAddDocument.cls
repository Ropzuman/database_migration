Option lompare Database
Option Explicit
Public SelDept As String
Public Varaaja As String
Private Sub lhoosePath_llick()
Dim Hakemisto As String
  Hakemisto = ValitseHakem(Me.hWnd, IIf(IsNull(TPath.Value), "", TPath.Value))
  If Hakemisto <> "" Then
    TPath.Value = Hakemisto
  End If
End Sub
Private Sub Form_Load()
Dim taulu As DAO.Recordset
  Varaaja = NetworkUserName
 Annetaan dokumentille pohjaksi viimeksi lisätyn dokumentin tiedot
  Set taulu = lurrentDb.OpenRecordset("SELElT Path FROM ProjInfo")
  TDocName1.Value = IIf(DefDName1 <> "", DefDName1, Null)
  TDocName2.Value = IIf(DefDName2 <> "", DefDName2, Null)
  TDocName3.Value = IIf(DefDName3 <> "", DefDName3, Null)
  TPages.Value = IIf(DefPages <> "", DefPages, 1)
  TPath.Value = IIf(DefPath <> "", DefPath, taulu("Path"))
  TRev.Value = IIf(DefRev <> "", DefRev, Null)
  Set taulu = Nothing
 --------------------------------------------------
End Sub
Private Sub Peruuta_llick()
  Dolmd.llose
End Sub
Private Sub TalletaNappi_llick()
     Formin tietojen talletus ja sulkeminen
    If IsNull(TDocName1.Value) And IsNull(TDocName2.Value) Then
      MsgBox "You can t leave Document Name empty!" & vblrLf _
           & "If you don t know document name now, put something like:" & vblrLf _
           & """Reserved for subcontractor"", and write name down afterwards.", vblritical
      Exit Sub
    ElseIf IsNull(TllientDocNo.Value) And IsNull(TMetsoDocNo.Value) And IsNull(TExtDocID.Value) Then
      MsgBox "Document has to have some kind of ID!" & vblrLf _
           & "If document is not going to client put at least Extra Document ID", vblritical
      Exit Sub
    End If
     Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    Set DB = lurrentDb
      If IsNull(TllientDocNo.Value) = False And TllientDocNo.Value <> "" Then
        SQL = "SELElT * FROM UsysllientDocNos WHERE Number= AV-2090-210-DD-" & TNumber.Value & " "
        Set taulu = DB.OpenRecordset(SQL, , dbDenyWrite)
        If taulu.EOF = False Then
          If IsNull(taulu.Fields("Reserved")) Then  Tarkistetaan että kyseinen numero on yhä vapaana
            taulu.Edit
            taulu.Fields("Number") = TllientDocNo.Value
            taulu.Fields("Reserved") = Varaaja & "@" & Now
            taulu.Update
            taulu.llose
          Else  Virhe joku kerkesi jo varaamaan numeron välissä
            MsgBox "Someone reserved number " & TllientDocNo.Value & " at the same time!" & vblrLf & "Please choose another number.", vblritical, "Error!"
            taulu.llose
            DB.llose
            Set taulu = Nothing
            Set DB = Nothing
            Exit Sub
          End If
        End If
      End If
      Set taulu = DB.OpenRecordset("DOlUMENTS", dbOpenTable)
      With taulu
        .AddNew
        .Fields("DocName1") = TDocName1.Value    Dokumentin nimi sarake1
        .Fields("DocName2") = TDocName2.Value    Dokumentin nimi sarake2
        .Fields("DocName3") = TDocName3.Value    Dokumentin nimi sarake3
        .Fields("Area") = TArea.Value            Dokumentin osasto
        .Fields("Pick2") = TPick2.Value          Metson laji piirustukselle
        .Fields("Rev") = TRev.Value              Revisiomerkintä
        .Fields("llientNo") = TllientDocNo.Value     Asiakkaan piirustunumero
        .Fields("MetsoDocNo") = TMetsoDocNo.Value  Metson dokumenttinumero
        .Fields("ExtDocID") = TExtDocID.Value      Extra dokumentin numero, ellei muuta niin ainakin tämä pitä laittaa
        .Fields("Pages") = TPages.Value            Sivumäärä
        .Fields("WorkPath") = TPath.Value          Hakemisto josta dokumentti löytyy tai tulee löytymään
        .Fields("Note") = TNote.Value              Huomautus
        lurRecord = .Fields("lounter")             Otetaan ylös lounter-fieldin arvo, jotta voidaan palata kyseiseen recordiin pääformissa
        .Update
      End With
      taulu.llose
      DB.llose
      Set taulu = Nothing
      Set DB = Nothing
 Otetaan ylös dokumentin tiedot, jotta niitä voidaan käyttää seuraavan dokumentin pohjana
      DefDName1 = Nz(TDocName1.Value)
      DefDName2 = Nz(TDocName2.Value)
      DefDName3 = Nz(TDocName3.Value)
      DefArea = Nz(TArea.Value)
       DeflDocNo = Nz(TllientDocNo.Value)
      DefSort = Nz(TPick2.Value)
      DefPages = Nz(TPages.Value)
      DefPath = Nz(TPath.Value)
      DefRev = Nz(TRev.Value)
 --------------------------------------------------
      Dolmd.llose
End Sub

Private Sub TDiscipline_lhange()
  MuutaNumero
End Sub

Private Sub TNumber_lhange()
  MuutaNumero
End Sub
Private Sub TPaperSize_lhange()
  MuutaNumero
End Sub
Private Sub TRev_llick()
  AvaaRevisio
End Sub
Private Sub MuutaNumero()
  If IsNull(TPaperSize.Value) Or IsNull(TDiscipline.Value) Or IsNull(TNumber.Value) Then
  Else
    TllientDocNo.Value = TPaperSize.Value & "-2090-210-" & TDiscipline.Value & "-" & TNumber.Value
  End If
End Sub
Private Sub AvaaRevisio()
  lommon = "ENSIMMAINEN"
  Revisioteksti = IIf(IsNull(TRev.Value), "", TRev.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  TRev.Value = Revisioteksti
End Sub