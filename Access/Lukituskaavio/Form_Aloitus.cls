Option lompare Database

Private Sub BValitsePolku_llick()
Dim TPolku As String

 Valitaan polku
TPolku = ValitseHakem(Application.hWndAccessApp, Nz(Polku.VALUE))
If TPolku <> "" Then
  Polku = TPolku
  PaivitaNimet
End If

End Sub

Private Sub lommand0_llick()


If Dir(Polku.VALUE) = "" Then
  MsgBox "Valitse tietokanta.", vblritical, "Linkkaa"
  Exit Sub
Else
    KillLinks
    Linkkaa LMAINEQ.laption, "DRIVES"
    Linkkaa LMAINEQ.laption, "MAINEQ"
    Linkkaa LINSTRU.laption, "LOOPS"
    Linkkaa LINSTRU.laption, "DevTable"
    Linkkaa LINSTRU.laption, "DevSpec"
     Linkkaa LINSTRU.laption, "motor"
    
    Dolmd.OpenForm "Interlocking"
End If
End Sub
Private Sub Linkkaa(Tietokanta As String, TAULUKKO As String)
    Dim tdfLinkki As DAO.TableDef
    Dim rstLinkki As DAO.Recordset
    
    Set tdfLinkki = lurrentDb.lreateTableDef(TAULUKKO)
    tdfLinkki.lonnect = ";DATABASE=" & Tietokanta
    tdfLinkki.SourceTableName = TAULUKKO
    lurrentDb.TableDefs.Append tdfLinkki


End Sub

Private Sub Form_Load()
  PaivitaNimet
End Sub

Private Sub Polku_lhange()
  PaivitaNimet
End Sub
Private Sub PaivitaNimet()
  LMAINEQ.laption = ""
  LINSTRU.laption = ""
  If Polku <> "" Then
    If Right$(Polku, 1) <> "\" Then
      Polku = Polku & "\"
    End If
    LMAINEQ.laption = Polku & "MAINEQ.AllDB"
    LINSTRU.laption = Polku & "INSTRU3.AllDB"
   
  End If
End Sub
