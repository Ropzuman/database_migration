Option Compare Database
Option Explicit

' K‰ytt‰j‰tietuetyyppi kirjautumisnimien purkamiseen LACCDB-tiedostosta

Private Type UserRec
   bMach(1 To 32) As Byte  ' Ensimm‰iset 32 tavua: koneen nimi
   bUser(1 To 32) As Byte  ' Seuraavat 32 tavua: k‰ytt‰j‰nimi
End Type

Private Sub Command27_Click()
    Dim Viesti As String
    Dim PuhdistettuViesti As String
    Dim PuhdistettuKohde  As String

    Viesti = InputBox("Syˆt‰ viesti k‰ytt‰j‰lle " & Me.NetworkName.Value & ":" & vbCrLf & vbCrLf, "L‰het‰ viesti")

    If Trim(Viesti) <> "" Then
        ' Estet‰‰n Command Injection ó poistetaan lainausmerkit ja Shell-erikoismerkit
        PuhdistettuViesti = Replace(Viesti, """", "'")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "&", "ja")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "|", "")
        PuhdistettuKohde  = Replace(CStr(Nz(Me.NetworkName.Value, "")), " ", "")

        ' K‰ytet‰‰n modernia MSG.EXE-komentoa ó NET SEND poistettu Windows Vistasta l‰htien
        Call Shell("msg " & PuhdistettuKohde & " """ & PuhdistettuViesti & """", vbHide)
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

'-------------------------------------------------------------------------------------
'   Funktio  : WhosOn()
'   Tarkoitus: Lukee *.LACCDB-tiedoston ja palauttaa kirjautuneiden k‰ytt‰jien
'              ja heid‰n tietokoneensa tiedot.
'
'              LACCDB-tietueessa on 64 tavua.
'
'              Koneen nimi alkaa tavusta 1 ja on null-p‰‰tteinen.
'
'              Kirjautumisnimi alkaa tavusta 33 ja on myˆs null-p‰‰tteinen.
'
'-------------------------------------------------------------------------------------
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer
   Dim i As Integer
   Dim SPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' K‰ytt‰j‰tietue, m‰‰ritelty moduulin alussa
   Dim dbCurrent As DAO.Database

' Haetaan nykyisen tietokannan polku.

   Set dbCurrent = CurrentDb
   SPath = dbCurrent.Name
   ' CurrentDb-viittausta EI suljeta .Close-kutsulla ó vain Set Nothing
   Set dbCurrent = Nothing

' Rakennetaan LACCDB-tiedoston polku.

   SPath = Left(SPath, InStr(1, SPath, ".")) + "LACCDB"

' Tarkistetaan, ett‰ tiedosto on olemassa.

   X = Dir(SPath)  'Tarkistetaan, ett‰ tiedosto on olemassa
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
      MsgBox "Couldn't populate the list", 48, "No LDB File"
   Else
      MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
      Close iLDBFile
   End If
   Resume Exit_WhosOn

End Function

