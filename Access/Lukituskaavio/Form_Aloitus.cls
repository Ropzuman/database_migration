Option Compare Database
Option Explicit

Private Sub BValitsePolku_Click()
Dim TPolku As String

' Avataan kansioselain polun valitsemiseksi
TPolku = ValitseHakem(Application.hWndAccessApp, Nz(Me.Polku.Value))
If TPolku <> "" Then
  Me.Polku = TPolku
  PaivitaNimet
End If

End Sub

Private Sub Command0_Click()


If Dir(Me.Polku.Value) = "" Then
  MsgBox "Valitse tietokanta.", vbCritical, "Linkkaa"
  Exit Sub
Else
    KillLinks
    Linkkaa Me.LMAINEQ.Caption, "DRIVES"
    Linkkaa Me.LMAINEQ.Caption, "MAINEQ"
    Linkkaa Me.LINSTRU.Caption, "LOOPS"
    Linkkaa Me.LINSTRU.Caption, "DevTable"
    Linkkaa Me.LINSTRU.Caption, "DevSpec"
'    Linkkaa LINSTRU.Caption, "motor"
    
    DoCmd.OpenForm "Interlocking"
End If
End Sub
Private Sub Linkkaa(Tietokanta As String, TAULUKKO As String)
    Dim tdfLinkki As DAO.TableDef
    
    Set tdfLinkki = CurrentDb.CreateTableDef(TAULUKKO)
    tdfLinkki.Connect = ";DATABASE=" & Tietokanta
    tdfLinkki.SourceTableName = TAULUKKO
    CurrentDb.TableDefs.Append tdfLinkki
    Set tdfLinkki = Nothing

End Sub

Private Sub Form_Load()
  PaivitaNimet
End Sub

Private Sub Polku_Change()
  PaivitaNimet
End Sub
Private Sub PaivitaNimet()
  Me.LMAINEQ.Caption = ""
  Me.LINSTRU.Caption = ""
  If Me.Polku <> "" Then
    If Right$(Me.Polku, 1) <> "\" Then
      Me.Polku = Me.Polku & "\"
    End If
    Me.LMAINEQ.Caption = Me.Polku & "MAINEQ.ACCDB"
    Me.LINSTRU.Caption = Me.Polku & "INSTRU3.ACCDB"
   
  End If
End Sub
