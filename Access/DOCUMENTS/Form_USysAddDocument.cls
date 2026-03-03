Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysAddDocument
' SOVELLUS : DOCUMENTS — Uuden dokumentin lisäys
' KUVAUS   : Lisää uuden dokumentin rekisteriin. Käyttää ed. dokumentin
'            tietoja oletusarvoina nopeuttaakseen sarjalisausta.
'            Tarkistaa asiakkaan numeron varauksen ennen tallennusta.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Public SelDept As String
Public Varaaja As String
Private Sub ChoosePath_Click()
Dim Hakemisto As String
  Hakemisto = ValitseHakem(Me.hWnd, IIf(IsNull(Me.TPath.Value), "", Me.TPath.Value))
  If Hakemisto <> "" Then
    Me.TPath.Value = Hakemisto
  End If
End Sub
Private Sub Form_Load()
Dim taulu As DAO.Recordset
  Varaaja = NetworkUserName
' Käytetään viimeksi lisätyn dokumentin tietoja uuden pohjaksi
  Set taulu = CurrentDb.OpenRecordset("SELECT Path FROM ProjInfo")
  Me.TDocName1.Value = IIf(DefDName1 <> "", DefDName1, Null)
  Me.TDocName2.Value = IIf(DefDName2 <> "", DefDName2, Null)
  Me.TDocName3.Value = IIf(DefDName3 <> "", DefDName3, Null)
  Me.TPages.Value = IIf(DefPages <> "", DefPages, 1)
  Me.TPath.Value = IIf(DefPath <> "", DefPath, taulu("Path"))
  Me.TRev.Value = IIf(DefRev <> "", DefRev, Null)
  Set taulu = Nothing
'--------------------------------------------------
End Sub
Private Sub Peruuta_Click()
  DoCmd.Close
End Sub
Private Sub TalletaNappi_Click()
    'Formin tietojen talletus ja sulkeminen
    If IsNull(Me.TDocName1.Value) And IsNull(Me.TDocName2.Value) Then
      MsgBox "You can't leave Document Name empty!" & vbCrLf _
           & "If you don't know document name now, put something like:" & vbCrLf _
           & """Reserved for subcontractor"", and write name down afterwards.", vbCritical
      Exit Sub
    ElseIf IsNull(Me.TClientDocNo.Value) And IsNull(Me.TMetsoDocNo.Value) And IsNull(Me.TExtDocID.Value) Then
      MsgBox "Document has to have some kind of ID!" & vbCrLf _
           & "If document is not going to client put at least Extra Document ID", vbCritical
      Exit Sub
    End If
    'Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    Set DB = CurrentDb
      If IsNull(Me.TClientDocNo.Value) = False And Me.TClientDocNo.Value <> "" Then
        SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & Me.TNumber.Value & "'"
        Set taulu = DB.OpenRecordset(SQL, , dbDenyWrite)
        If taulu.EOF = False Then
          If IsNull(taulu.Fields("Reserved")) Then
            taulu.Edit
            taulu.Fields("Number") = Me.TClientDocNo.Value
            taulu.Fields("Reserved") = Varaaja & "@" & Now
            taulu.Update
            taulu.Close
          Else
            MsgBox "Someone reserved number " & Me.TClientDocNo.Value & " at the same time!" & vbCrLf & "Please choose another number.", vbCritical, "Error!"
            taulu.Close
            DB.Close
            Set taulu = Nothing
            Set DB = Nothing
            Exit Sub
          End If
        End If
      End If
      Set taulu = DB.OpenRecordset("DOCUMENTS", dbOpenTable)
      With taulu
        .AddNew
        .Fields("DocName1") = Me.TDocName1.Value
        .Fields("DocName2") = Me.TDocName2.Value
        .Fields("DocName3") = Me.TDocName3.Value
        .Fields("Area") = Me.TArea.Value
        .Fields("Pick2") = Me.TPick2.Value
        .Fields("Rev") = Me.TRev.Value
        .Fields("ClientNo") = Me.TClientDocNo.Value
        .Fields("MetsoDocNo") = Me.TMetsoDocNo.Value
        .Fields("ExtDocID") = Me.TExtDocID.Value
        .Fields("Pages") = Me.TPages.Value
        .Fields("WorkPath") = Me.TPath.Value
        .Fields("Note") = Me.TNote.Value
        CurRecord = .Fields("Counter")            'Otetaan ylös Counter-fieldin arvo, jotta voidaan palata kyseiseen recordiin pääformissa
        .Update
      End With
      taulu.Close
      DB.Close
      Set taulu = Nothing
      Set DB = Nothing
' Tallennetaan viimeksi lisätyn dokumentin tiedot seuraavan pohjaksi
      DefDName1 = Nz(Me.TDocName1.Value)
      DefDName2 = Nz(Me.TDocName2.Value)
      DefDName3 = Nz(Me.TDocName3.Value)
      DefArea = Nz(Me.TArea.Value)
      DefSort = Nz(Me.TPick2.Value)
      DefPages = Nz(Me.TPages.Value)
      DefPath = Nz(Me.TPath.Value)
      DefRev = Nz(Me.TRev.Value)
'--------------------------------------------------
      DoCmd.Close
End Sub

Private Sub TDiscipline_Change()
  MuutaNumero
End Sub

Private Sub TNumber_Change()
  MuutaNumero
End Sub
Private Sub TPaperSize_Change()
  MuutaNumero
End Sub
Private Sub TRev_Click()
  AvaaRevisio
End Sub
Private Sub MuutaNumero()
  If IsNull(Me.TPaperSize.Value) Or IsNull(Me.TDiscipline.Value) Or IsNull(Me.TNumber.Value) Then
  Else
    Me.TClientDocNo.Value = Me.TPaperSize.Value & "-2090-210-" & Me.TDiscipline.Value & "-" & Me.TNumber.Value
  End If
End Sub
Private Sub AvaaRevisio()
  Common = "ENSIMMAINEN"
  Revisioteksti = IIf(IsNull(Me.TRev.Value), "", Me.TRev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  Me.TRev.Value = Revisioteksti
End Sub