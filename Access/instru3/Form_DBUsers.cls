Option Compare Database
Option Explicit

' Käyttäjätietuetyyppi kirjautumisnimien purkamiseen LACCDB-tiedostosta

Private Type UserRec
   bMach(1 To 32) As Byte  ' Ensimmäiset 32 tavua: koneen nimi
   bUser(1 To 32) As Byte  ' Seuraavat 32 tavua: käyttäjänimi
End Type

Private Sub Command27_Click()
    Dim Viesti As String
    Dim PuhdistettuViesti As String
    Dim PuhdistettuKohde  As String

    Viesti = InputBox("Syötä viesti käyttäjälle " & Me.NetworkName.Value & ":" & vbCrLf & vbCrLf, "Lähetä viesti")

    If Trim(Viesti) <> "" Then
        ' Estetään Command Injection — poistetaan lainausmerkit ja Shell-erikoismerkit
        PuhdistettuViesti = Replace(Viesti, """", "'")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "&", "ja")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "|", "")
        PuhdistettuViesti = PuhdistaMerkit(PuhdistettuViesti, "", "^<>%!" & vbCr & vbLf)
        ' Kohteeksi sallitaan vain koneen/käyttäjän nimessä käytettävät merkit
        PuhdistettuKohde  = PuhdistaMerkit(CStr(Nz(Me.NetworkName.Value, "")), "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-.\", "")
        If PuhdistettuKohde = "" Then Exit Sub

        ' Käytetään modernia MSG.EXE-komentoa — NET SEND poistettu Windows Vistasta lähtien
        Call Shell("msg " & PuhdistettuKohde & " """ & PuhdistettuViesti & """", vbHide)
    End If
End Sub

' Poistaa merkkijonosta kielletyt merkit; jos Sallitut on annettu, vain ne säilytetään
Private Function PuhdistaMerkit(ByVal Teksti As String, ByVal Sallitut As String, ByVal Kielletyt As String) As String
    Dim i As Long  ' Merkin indeksi
    Dim Merkki As String  ' Käsiteltävä merkki
    Dim Tulos As String  ' Puhdistettu teksti

    For i = 1 To Len(Teksti)
        Merkki = Mid$(Teksti, i, 1)
        If Sallitut <> "" Then
            If InStr(1, Sallitut, Merkki, vbBinaryCompare) > 0 Then Tulos = Tulos & Merkki
        ElseIf InStr(1, Kielletyt, Merkki, vbBinaryCompare) = 0 Then
            Tulos = Tulos & Merkki
        End If
    Next i
    PuhdistaMerkit = Tulos
End Function

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

'-------------------------------------------------------------------------------------
'   Funktio  : WhosOn()
'   Tarkoitus: Lukee *.LACCDB-tiedoston ja palauttaa kirjautuneiden käyttäjien
'              ja heidän tietokoneensa tiedot.
'
'              LACCDB-tietueessa on 64 tavua.
'
'              Koneen nimi alkaa tavusta 1 ja on null-päätteinen.
'
'              Kirjautumisnimi alkaa tavusta 33 ja on myös null-päätteinen.
'
'-------------------------------------------------------------------------------------
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer
   Dim i As Integer
   Dim SPath As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' Käyttäjätietue, määritelty moduulin alussa
   Dim dbCurrent As DAO.Database

' Haetaan nykyisen tietokannan polku.

   Set dbCurrent = CurrentDb
   SPath = dbCurrent.Name
   ' CurrentDb-viittausta EI suljeta .Close-kutsulla — vain Set Nothing
   Set dbCurrent = Nothing

' Rakennetaan LACCDB-tiedoston polku.

   SPath = Left$(SPath, InStrRev(SPath, ".")) & "LACCDB"

' Tarkistetaan, että tiedosto on olemassa.

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
      MsgBox "Käyttäjälistaa ei voitu muodostaa: LACCDB-tiedostoa ei löydy.", 48, "LACCDB puuttuu"
   Else
      MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description
      Close iLDBFile
   End If
   Resume Exit_WhosOn

End Function

