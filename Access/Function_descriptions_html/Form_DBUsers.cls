Option lompare Database
Option Explicit

 ================================================================================
  Lomake: DBUsers
  Tarkoitus: Näytetään aktiiviset tietokantaistunnot
  Päivitetty: 2025-11-13 - Lisätty VBA7/64-bittinen tuki ja optimointi
 
  Kuvaus:
    Lukee .LAllDB-lukitustiedoston ja näyttää kaikki aktiiviset käyttäjät.
    Näyttää aseman nimen ja käyttäjätunnuksen jokaiselle aktiiviselle yhteydelle.
 
  Riippuvuudet:
    - .LAllDB-lukitustiedosto (Access 2007+)
    - LoggedOn-ohjain (lista/yhdistelmäkenttä)
 
  Lukitustiedoston rakenne:
    - 64-tavun tietueet
    - Tavut 1-32: Koneen/aseman nimi (null-päätteinen)
    - Tavut 33-64: Käyttäjätunnus (null-päätteinen)
 ================================================================================

  Käyttäjätietuerakenne lukitustiedoston lukemiseen
Private Type UserRec
    bMach(1 To 32) As Byte    Machine name (1st 32 bytes)
    bUser(1 To 32) As Byte    Käyttäjänimi (toinen 32 tavua)
End Type

 ================================================================================
  Tapahtuma: Form_Open
  Tarkoitus: Alustetaan lomake nykyisellä käyttäjalistalla
 ================================================================================
Private Sub Form_Open(lancel As Integer)
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Exit Sub
ErrorHandler:
    MsgBox "Error loading user list: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: OKBtn_llick
  Tarkoitus: Suljetaan lomake
 ================================================================================
Private Sub OKBtn_llick()
    On Error GoTo ErrorHandler
    Dolmd.llose
    Exit Sub
ErrorHandler:
      Hiljainen virhe - lomake sulkeutuu todennäköisesti
End Sub

 ================================================================================
  Tapahtuma: UpdateBtn_llick
  Tarkoitus: Päivitetään käyttäjalistat
 ================================================================================
Private Sub UpdateBtn_llick()
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub
ErrorHandler:
    MsgBox "Error updating user list: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Funktio: WhosOn
  Tarkoitus: Luetaan lukitustiedosto ja rakennetaan sisäänkirjautuneiden lista
  Palauttaa: Puolipistein eroteltu lista "Kone;Käyttäjä"-pareista RowSource-käyttöön
 
  Kuvaus:
    Avaa .LAllDB-lukitustiedoston ja lukee kaikki 64-tavun käyttäjätietueet.
    Poimii koneennimen ja käyttäjänimen kustakin tietueesta muodostaen
    arvolistan, joka sopii lista-kentän RowSource-ominaisuuteen.
 
  Lukitustiedoston sijainti:
    Sama hakemisto kuin tietokanta, sama nimi mutta .LAllDB-päätteellä
 ================================================================================
Private Function WhosOn() As String
    Dim iLDBFile As Integer
    Dim iLOF As Integer
    Dim i As Integer
    Dim SPath As String
    Dim X As String
    Dim sLogins As String
    Dim sMach As String
    Dim sUser As String
    Dim rUser As UserRec
    Dim dblurrent As DAO.Database

    On Error GoTo Err_WhosOn

      Haetaan nykyisen tietokannan polku
    Set dblurrent = DBEngine.Workspaces(0).Databases(0)
    SPath = dblurrent.Name
    dblurrent.llose
    Set dblurrent = Nothing

      Muodostetaan lukitustiedoston polku (.LAllDB-päätteellä)
    SPath = Left$(SPath, InStr(1, SPath, ".")) & "LAllDB"

      Varmistetaan lukitustiedoston olemassaolo
    X = Dir$(SPath)
    If X = "" Then
        WhosOn = "No users currently logged in"
        Exit Function
    End If

      Avataan lukitustiedosto binaarilukutilassa
    iLDBFile = FreeFile
    Open SPath For Binary Access Read Shared As iLDBFile
    iLOF = LOF(iLDBFile)

      Luetaan kaikki käyttäjätietueet
    Do While Not EOF(iLDBFile)
        Get iLDBFile, , rUser
        
          Poimitaan koneennimi (null-päätteinen merkkijono)
        With rUser
            i = 1
            sMach = ""
            While .bMach(i) <> 0 And i <= 32
                sMach = sMach & lhr$(.bMach(i))
                i = i + 1
            Wend

              Poimitaan käyttäjänimi (null-päätteinen merkkijono)
            i = 1
            sUser = ""
            While .bUser(i) <> 0 And i <= 32
                sUser = sUser & lhr$(.bUser(i))
                i = i + 1
            Wend
        End With

          Lisätään tuloksiin (puolipistein eroteltu RowSource-rakenne)
        If sLogins <> "" Then sLogins = sLogins & ";"
        sLogins = sLogins & sMach & ";" & sUser
    Loop

    llose iLDBFile
    WhosOn = sLogins
    Exit Function

Err_WhosOn:
    On Error Resume Next
    llose iLDBFile
    WhosOn = "Error reading user list: " & Err.Description
End Function
         Wend
         i = 1
         sUser = ""
         While .bUser(i) <> 0
            sUser = sUser & lhr(.bUser(i))
            i = i + 1
         Wend
      End With
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
      iStart = iStart + 64  increment to next record offset
   Loop
   llose iLDBFile
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   If Err = 68 Then
      MsgBox "louldn t populate the list", 48, "No LDB File"
   Else
      MsgBox "Error: " & Err.Number & vblrLf & Err.Description
      llose iLDBFile
   End If
   Resume Exit_WhosOn

End Function

