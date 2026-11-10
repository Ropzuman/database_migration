Option Compare Database
Option Explicit

'================================================================================
' Lomake: DBUsers
' Tarkoitus: N‰ytt‰‰ tietokantaan kirjautuneet k‰ytt‰j‰t ja l‰hett‰‰ heille viestin
' P‰ivitetty: 2026-11-10 - net send korvattu msg.exe-komennolla, syˆtteiden
'             puhdistus, LACCDB-polun muodostus ja kommentit suomeksi
'
' Riippuvuudet:
'   - Tietokannan lukitustiedosto (.laccdb)
'   - msg.exe (Windows)
'================================================================================

' Lukitustiedoston tietue: ensimm‰iset 32 tavua koneen nimi, seuraavat 32 k‰ytt‰j‰nimi
Private Type UserRec
   bMach(1 To 32) As Byte  ' Koneen nimi (nollap‰‰tteinen)
   bUser(1 To 32) As Byte  ' K‰ytt‰j‰nimi (nollap‰‰tteinen)
End Type

'--------------------------------------------------------------------------------
' Tapahtuma: Command27_Click
' Tarkoitus: L‰hett‰‰ viestin valitulle k‰ytt‰j‰n koneelle msg.exe-komennolla
'
' Huomiot:
'   - net send on poistettu Windowsista Vistan j‰lkeen, joten k‰ytet‰‰n msg.exe:t‰
'   - Kohteeksi sallitaan vain koneen/k‰ytt‰j‰n nimess‰ k‰ytett‰v‰t merkit ja
'     viestist‰ poistetaan Shell-erikoismerkit (komentoinjektion esto)
'--------------------------------------------------------------------------------
Private Sub Command27_Click()
    Dim Viesti As String  ' K‰ytt‰j‰n kirjoittama viesti
    Dim PuhdistettuViesti As String  ' Shell-kutsuun kelpaava viesti
    Dim PuhdistettuKohde As String  ' Shell-kutsuun kelpaava vastaanottaja

    Viesti = InputBox("Kirjoita viesti k‰ytt‰j‰lle " & Me.NetworkName.Value & ":", "L‰het‰ viesti")
    If Trim$(Viesti) = "" Then Exit Sub

    PuhdistettuViesti = Replace(Viesti, """", "'")
    PuhdistettuViesti = Replace(PuhdistettuViesti, "&", "ja")
    PuhdistettuViesti = PuhdistaMerkit(PuhdistettuViesti, "", "|^<>%!" & vbCr & vbLf)
    PuhdistettuKohde = PuhdistaMerkit(CStr(Nz(Me.NetworkName.Value, "")), "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-.\\", "")
    If PuhdistettuKohde = "" Then
        MsgBox "Vastaanottajan nimi puuttuu tai sis‰lt‰‰ kiellettyj‰ merkkej‰.", vbExclamation, "L‰het‰ viesti"
        Exit Sub
    End If

    Call Shell("msg " & PuhdistettuKohde & " """ & PuhdistettuViesti & """", vbHide)
End Sub

' Poistaa merkkijonosta kielletyt merkit; jos Sallitut on annettu, vain ne s‰ilytet‰‰n
Private Function PuhdistaMerkit(ByVal Teksti As String, ByVal Sallitut As String, ByVal Kielletyt As String) As String
    Dim i As Long  ' Merkin indeksi
    Dim Merkki As String  ' K‰sitelt‰v‰ merkki
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

'--------------------------------------------------------------------------------
' Funktio: WhosOn
' Tarkoitus: Lukee .laccdb-tiedoston ja palauttaa kirjautuneet k‰ytt‰j‰t ja koneet
'            puolipisteell‰ eroteltuna ("kone -- k‰ytt‰j‰;")
'
' Huomiot:
'   - Tietue on 64 tavua: koneen nimi tavusta 1, k‰ytt‰j‰nimi tavusta 33,
'     molemmat nollap‰‰tteisi‰
'   - Nimen luku rajataan 32 tavuun, jottei nollap‰‰tteetˆn tietue aiheuta virhett‰
'--------------------------------------------------------------------------------
Private Function WhosOn() As String
On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer  ' Lukitustiedoston kahva
   Dim TiedostoAuki As Boolean  ' Onko tiedosto avattu (virheenk‰sittely‰ varten)
   Dim i As Integer  ' Tavun indeksi
   Dim SPath As String  ' Lukitustiedoston polku
   Dim sLogStr As String  ' Yksitt‰isen kirjautumisen teksti
   Dim sLogins As String  ' Kaikki kirjautumiset
   Dim sMach As String  ' Koneen nimi
   Dim sUser As String  ' K‰ytt‰j‰nimi
   Dim rUser As UserRec  ' Luettava tietue

   ' Lukitustiedosto on tietokannan kanssa samassa kansiossa, p‰‰te .laccdb/.ldb
   ' (viimeinen piste erottaa p‰‰tteen, vaikka kansion nimess‰ olisi pisteit‰)
   SPath = CurrentDb.Name
   SPath = Left$(SPath, InStrRev(SPath, ".")) & "LACCDB"

   If Len(Dir$(SPath)) = 0 Then Err.Raise 68  ' Tiedostoa ei lˆydy

   iLDBFile = FreeFile
   Open SPath For Binary Access Read Shared As iLDBFile
   TiedostoAuki = True
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      With rUser
         i = 1
         sMach = ""
         Do While i <= 32
            If .bMach(i) = 0 Then Exit Do
            sMach = sMach & Chr$(.bMach(i))
            i = i + 1
         Loop
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
   Loop
   Close iLDBFile
   TiedostoAuki = False
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   If Err.Number = 68 Then
      MsgBox "Lukitustiedostoa (.laccdb) ei lˆytynyt.", vbExclamation, "Ei laccdb-tiedostoa"
   Else
      MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbExclamation
   End If
   If TiedostoAuki Then Close iLDBFile
   Resume Exit_WhosOn

End Function
