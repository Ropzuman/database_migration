Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_DBUsers
' SOVELLUS : FunctionDiagrams - Käyttäjäseuranta
' KUVAUS   : Näyttää kaikki aktiiviset käyttäjät lukemalla .LACCDB-
'            tiedoston WhosOn()-funktiolla. Tiedosto on 64-tavuinen
'            binaaritietue: koneen nimi tavuissa 1–32, käyttäjänimi 33–64.
'            Lomake mahdollistaa myös viestin lähettämisen yksittäiselle
'            käyttäjälle Shell-komennon kautta.
' HUOM    : InputBox-kirjoitusasun korjaus; dbCurrent.Close korvattu Set = Nothing.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' Tietuemääritys LACCDB-tiedoston 64-tavuisen tietueen lukemiseen
Private Type UserRec
   bMach(1 To 32) As Byte  ' Koneen nimi (tavut 1–32)
   bUser(1 To 32) As Byte  ' Käyttäjänimi (tavut 33–64)
End Type
Private Sub Command27_Click()
Dim Viesti As String
  ' Lähetetään viesti valitulle verkkokäyttäjälle net send -komennolla
  Viesti = InputBox("Send message to user " & NetworkName.Value & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    ' Lainausmerkit korvataan, jotta viesti ei riko komentoriviä
    Viesti = Replace(Viesti, """", "'")
    Call Shell("net send " & NetworkName.Value & " """ & Viesti & """")
  End If
End Sub

Private Sub Form_Open(Cancel As Integer)
   
   Me.LoggedOn.RowSource = WhosOn()

End Sub

Private Sub OKBtn_Click()
   
   DoCmd.Close

End Sub

Private Sub UpdateBtn_Click()
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
End Sub

'--------------------------------------------------------------------------
' WhosOn() — Lukee aktiivisen .LACCDB-tiedoston ja palauttaa
' puolipisteellä eroteltujen "KoneNimi -- KäyttäjäNimi" -parien listan.
' LACCDB-tietue on 64 tavua: koneen nimi tavuissa 1–32, käyttäjänimi 33–64.
' Luetaan nykyisen tietokannan lukitustiedoston kone- ja käyttäjätiedot.
Private Function WhosOn() As String
  Dim Tiedosto As Integer
  Dim TiedostoAuki As Boolean
  Dim Tietueita As Long
  Dim TietueNro As Long
  Dim i As Long
  Dim PaatePos As Long
  Dim Polku As String
  Dim Paate As String
  Dim Kone As String
  Dim Kayttaja As String
  Dim Rivi As String
  Dim Kirjautumiset As String
  Dim Virhekuvaus As String
  Dim Tietue As UserRec
  Dim DB As DAO.Database

  On Error GoTo Virhe

  Set DB = CurrentDb
  Polku = DB.Name
  PaatePos = InStrRev(Polku, ".")
  If PaatePos <= InStrRev(Polku, "\") Then
    Err.Raise vbObjectError + 1000, "WhosOn", "Tietokannan tiedostopääte puuttuu."
  End If

  Paate = LCase$(Mid$(Polku, PaatePos + 1))
  Select Case Paate
    Case "mdb", "mde"
      Polku = Left$(Polku, PaatePos) & "ldb"
    Case "accdb", "accde", "accdr"
      Polku = Left$(Polku, PaatePos) & "laccdb"
    Case Else
      Err.Raise vbObjectError + 1001, "WhosOn", "Tuntematon tietokannan tiedostopääte: " & Paate
  End Select

  Tiedosto = FreeFile
  Open Polku For Binary Access Read Shared As #Tiedosto
  TiedostoAuki = True
  If LOF(Tiedosto) Mod 64 <> 0 Then
    Err.Raise vbObjectError + 1002, "WhosOn", "Lukitustiedoston koko ei vastaa 64 tavun tietueita."
  End If
  Tietueita = LOF(Tiedosto) \ 64

  For TietueNro = 1 To Tietueita
    Get #Tiedosto, , Tietue
    Kone = ""
    Kayttaja = ""
    For i = 1 To 32
      If Tietue.bMach(i) = 0 Then Exit For
      Kone = Kone & Chr$(Tietue.bMach(i))
    Next i
    For i = 1 To 32
      If Tietue.bUser(i) = 0 Then Exit For
      Kayttaja = Kayttaja & Chr$(Tietue.bUser(i))
    Next i
    Rivi = Kone & " -- " & Kayttaja
    If InStr(1, ";" & Kirjautumiset, ";" & Rivi & ";", vbTextCompare) = 0 Then
      Kirjautumiset = Kirjautumiset & Rivi & ";"
    End If
  Next TietueNro

  Close #Tiedosto
  TiedostoAuki = False
  Set DB = Nothing
  WhosOn = Kirjautumiset
  Exit Function

Virhe:
  Virhekuvaus = Err.Description
  If TiedostoAuki Then
    On Error Resume Next
    Close #Tiedosto
    On Error GoTo 0
  End If
  Set DB = Nothing
  MsgBox "Käyttäjälistan lukeminen epäonnistui: " & Virhekuvaus, vbExclamation, "Käyttäjälista"
End Function
