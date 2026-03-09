Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_DBUsers
' SOVELLUS : FunctionDiagrams — Käyttäjäseuranta
' KUVAUS   : Näyttää kaikki aktiiviset käyttäjät lukemalla .LACCDB-
'            tiedoston WhosOn()-funktiolla. Tiedosto on 64-tavuinen
'            binaaritietue: koneen nimi tavuissa 1–32, käyttäjänimi 33–64.
'            Lomake mahdollistaa myös viestin lähettämisen yksittäiselle
'            käyttäjälle Shell-komennon kautta.
' DIPLE    : InputBox-kirjoitusasun korjaus; dbCurrent.Close → Set = Nothing.
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
'--------------------------------------------------------------------------
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer
   Dim i As Integer
   Dim SPath As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' Defined in General
   Dim dbCurrent As DAO.Database

   ' Nykyisen tietokannan polku LACCDB-tiedoston etsintään
   Set dbCurrent = CurrentDb
   SPath = dbCurrent.Name
   Set dbCurrent = Nothing

   ' Muodostetaan LACCDB-tiedostopolku (.accdb → .LACCDB)
   SPath = Left$(SPath, InStr(1, SPath, ".")) + "LACCDB"

   iLDBFile = FreeFile

   Open SPath For Binary Access Read Shared As iLDBFile
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      With rUser
         i = 1
         sMach = ""
         While .bMach(i) <> 0
            sMach = sMach & Chr(.bMach(i))
            i = i + 1
         Wend
         i = 1
         sUser = ""
         While .bUser(i) <> 0
            sUser = sUser & Chr(.bUser(i))
            i = i + 1
         Wend
      End With
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
   Loop
   Close iLDBFile
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   If Err = 68 Then
      MsgBox "LACCDB-tiedostoa ei löytynyt", 48, "Kirjautuneet käyttäjät"
   Else
      MsgBox "Virhe " & Err.Number & vbCrLf & Err.Description
      Close iLDBFile
   End If
   Resume Exit_WhosOn

End Function

