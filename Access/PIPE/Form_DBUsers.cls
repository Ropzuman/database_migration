Option Compare Database
Option Explicit

'================================================================================
' Lomake: DBUsers
' Tarkoitus: Näyttää tietokantaan tällä hetkellä kirjautuneet käyttäjät
'
' Kuvaus:
'   Lukee .LACCDB-lukitustiedostosta aktiiviset käyttäjät:
'   - Näyttää tietokoneen nimen ja käyttäjänimen jokaiselle yhteydelle
'   - Päivityspainike listan uudelleenlataamista varten
'   - Viestin lähetys (net send, ei toimi uusissa Windowseissa)
'
' Riippuvuudet:
'   - .LACCDB-tiedosto (Access 2013+ lukitustiedosto)
'   - WhosOn()-funktio lukitustiedoston lukemiseen
'================================================================================


' Tietueen esittely kirjautumistietojen purkamista varten
Private Type UserRec
   bMach(1 To 32) As Byte  ' 1st 32 bytes hold machine name
   bUser(1 To 32) As Byte  ' 2nd 32 bytes hold user name
End Type
'--------------------------------------------------------------------------------
' Tapahtuma: Command27_Click
' Tarkoitus: Lähetä verkkoviesti valitulle käyttäjälle (vanhentunut toiminto)
' Huom: Käyttää vanhaa "net send" -komentoa — ei toimi uusissa Windowseissa
'--------------------------------------------------------------------------------
Private Sub Command27_Click()
On Error GoTo ErrorHandler
Dim Viesti As String
  Viesti = InputBox("Lähetä viesti käyttäjälle " & Me.NetworkName.Value & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Anna viesti:", "Lähetä viesti")
  ' HUOM: net send poistettu Windows Vistasta alkaen — käytetään msg.exe-komentoa
  If Viesti <> "" Then
    Shell "msg " & Me.NetworkName.Value & " """ & Viesti & """"
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Virhe viestin lähetyksessä: " & Err.Description, vbCritical, "Lähetä viesti"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Open
' Tarkoitus: Täyttää kirjautuneiden käyttäjien lista lomakkeen avautuessa
'--------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrorHandler
   Me.LoggedOn.RowSource = WhosOn()
   Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen avaamisessa: " & Err.Description, vbCritical, "DBUsers"
    Cancel = True
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: OKBtn_Click
' Tarkoitus: Sulkee lomakkeen
'--------------------------------------------------------------------------------
Private Sub OKBtn_Click()
   DoCmd.Close
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: UpdateBtn_Click
' Tarkoitus: Päivittää kirjautuneiden käyttäjien listan
'--------------------------------------------------------------------------------
Private Sub UpdateBtn_Click()
On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub

ErrorHandler:
    MsgBox "Virhe listan päivittämisessä: " & Err.Description, vbCritical, "Päivitä"
End Sub

'--------------------------------------------------------------------------------
' Funktio: WhosOn
' Tarkoitus: Lukee .LACCDB-lukitustiedostosta nykyiset kirjautuneet käyttäjät
' Palautusarvo: Puolipisteillä eroteltu merkkijono "KoneNimi -- Käyttäjä"-pareista
' Huom: LACCDB-tietuerakenne: 64 tavua / tietue
'   - Tavut 1–32: tietokoneen nimi (nollapäätteinen)
'   - Tavut 33–64: käyttäjänimi (nollapäätteinen)
'--------------------------------------------------------------------------------
Private Function WhosOn() As String
On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim i As Integer
   Dim sPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' User record structure
' Haetaan nykyisen tietokannan polku
   sPath = CurrentDb.Name

' Muodostetaan kirjautumistiedostopolku (.laccdb)
   sPath = Left$(sPath, InStr(1, sPath, ".")) & "LACCDB"

' Test for valid file
   X = Dir$(sPath)
   If X = "" Then
      WhosOn = ""
      GoTo Exit_WhosOn
   End If
   
   iStart = 1
   iLDBFile = FreeFile

   Open sPath For Binary Access Read Shared As iLDBFile

   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      With rUser
         ' Poimitaan tietokoneen nimi (nollapäätteinen)
         ' KORJATTU: Do While + Exit Do, koska VBA ei oikosulje And-operaattoria
         ' (While .bMach(i)<>0 And i<=32 lukisi indeksin 33 ennen rajakieltoa -> Error 9)
         i = 1
         sMach = ""
         Do While i <= 32
            If .bMach(i) = 0 Then Exit Do
            sMach = sMach & Chr$(.bMach(i))
            i = i + 1
         Loop

         ' Poimitaan käyttäjänimi (nollapäätteinen)
         i = 1
         sUser = ""
         Do While i <= 32
            If .bUser(i) = 0 Then Exit Do
            sUser = sUser & Chr$(.bUser(i))
            i = i + 1
         Loop
      End With
      
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
      iStart = iStart + 64 ' Siirrytään seuraavaan tietueeseen (64 tavua / tietue)
   Loop
   
   Close iLDBFile
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   ' Tallennetaan alkuperäinen virhetieto ennen Close-kutsu (muuten ylikirjoittuu)
   Dim lErrNum As Long
   Dim sErrDesc As String
   lErrNum = Err.Number
   sErrDesc = Err.Description
   On Error Resume Next
   If iLDBFile > 0 Then Close #iLDBFile
   On Error GoTo 0
   
   If lErrNum = 68 Then
      MsgBox "Tiedostoa ei löytynyt.", 48, "Ei LACCDB-tiedostoa"
   Else
      MsgBox "Virhe: " & lErrNum & vbCrLf & sErrDesc, vbCritical, "WhosOn"
   End If
   WhosOn = ""
   Resume Exit_WhosOn

End Function

