Option Compare Database

Private Sub BValitsePolku_Click()
Dim TPolku As String

'Valitaan polku
TPolku = ValitseHakem(Application.hWndAccessApp, Nz(Polku.VALUE))
If TPolku <> "" Then
  Polku = TPolku
  PaivitaNimet
End If

End Sub

Private Sub Command0_Click()


If Dir(Polku.VALUE) = "" Then
  MsgBox "Valitse tietokanta.", vbCritical, "Linkkaa"
  Exit Sub
Else
    KillLinks
    Linkkaa LMAINEQ.Caption, "DRIVES"
    Linkkaa LMAINEQ.Caption, "MAINEQ"
    Linkkaa LINSTRU.Caption, "LOOPS"
    Linkkaa LINSTRU.Caption, "DevTable"
    Linkkaa LINSTRU.Caption, "DevSpec"
'    Linkkaa LINSTRU.Caption, "motor"
    
    DoCmd.OpenForm "Interlocking"
End If
End Sub
Private Sub Linkkaa(Tietokanta As String, TAULUKKO As String)
    Dim tdfLinkki As DAO.TableDef
    Dim rstLinkki As DAO.Recordset
    
    Set tdfLinkki = CurrentDb.CreateTableDef(TAULUKKO)
    tdfLinkki.Connect = ";DATABASE=" & Tietokanta
    tdfLinkki.SourceTableName = TAULUKKO
    CurrentDb.TableDefs.Append tdfLinkki


End Sub

Private Sub Form_Load()
  PaivitaNimet
End Sub

Private Sub Polku_Change()
  PaivitaNimet
End Sub
Private Sub PaivitaNimet()
  LMAINEQ.Caption = ""
  LINSTRU.Caption = ""
  If Polku <> "" Then
    If Right$(Polku, 1) <> "\" Then
      Polku = Polku & "\"
    End If
    LMAINEQ.Caption = Polku & "MAINEQ.ACCDB"
    LINSTRU.Caption = Polku & "INSTRU3.ACCDB"
   
  End If
End Sub
