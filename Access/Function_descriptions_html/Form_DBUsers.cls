Option Compare Database
Option Explicit

'================================================================================
' Lomake: DBUsers
' Tarkoitus: Näyttää tietokantaan tällä hetkellä kirjautuneet käyttäjät
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — Korjattu dbCurrent.Close-virhe; poistettu orvokoodi
'
' Kuvaus:
'   Lukee .LACCDB-lukitustiedoston, jotta nähdään ketkä käyttävät tietokantaa.
'   Näyttää aseman nimen ja kirjautuneen käyttäjän jokaisesta aktiivisesta
'   yhteydestä.
'
' Riippuvuudet:
'   - .LACCDB-lukitustiedosto (Access 2007+)
'   - LoggedOn-ohjain (lista-/yhdistelmäruutu)
'
' Lukitustiedoston rakenne:
'   - 64 bitin tietueet
'   - Tavut 1–32: Tietokoneen nimi (null-päätetty merkkijono)
'   - Tavut 33–64: Käyttäjän kirjautumisnimi (null-päätetty merkkijono)
'================================================================================

' Käyttäjätietue lukitustiedoston lukemiseen
Private Type UserRec
    bMach(1 To 32) As Byte  ' Tietokoneen nimi (ensimmäiset 32 tavua)
    bUser(1 To 32) As Byte  ' Käyttäjänimi (jälkimmäiset 32 tavua)
End Type

'================================================================================
' Tapahtuma: Form_Open
' Tarkoitus: Alustaa lomake nykyisellä käyttäjälistalla lomakkeen avautuessa
'================================================================================
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Exit Sub
ErrorHandler:
    MsgBox "Virhe käyttäjälistan latauksessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: OKBtn_Click
' Tarkoitus: Sulkee lomakkeen
'================================================================================
Private Sub OKBtn_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close
    Exit Sub
ErrorHandler:
    ' Hiljainen virhe — lomake suljetaan jo todennäköisesti muualta
End Sub

'================================================================================
' Tapahtuma: UpdateBtn_Click
' Tarkoitus: Päivittää käyttäjälistan
'================================================================================
Private Sub UpdateBtn_Click()
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub
ErrorHandler:
    MsgBox "Virhe käyttäjälistan päivityksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: WhosOn
' Tarkoitus: Lukee lukitustiedoston ja koostaa kirjautuneiden käyttäjien listan
' Palauttaa: Puolipisteellä eroteltu "Kone;Käyttäjä"-lista (RowSource-arvo)
'
' Kuvaus:
'   Avaa .LACCDB-lukitustiedoston ja lukee kaikki 64 tavun käyttäjätietueet.
'   Poimii tietokoneen nimen ja käyttäjänimen kustakin tietueesta sekä
'   rakentaa arvoluettelon, joka sopii listalaatikon RowSource-ominaisuudeksi.
'
' Tiedoston sijainti:
'   Sama hakemisto kuin tietokanta, sama nimi mutta .LACCDB-pääte
'================================================================================
Private Function WhosOn() As String
    Dim iLDBFile As Integer
    Dim i As Integer
    Dim SPath As String
    Dim X As String
    Dim sLogins As String
    Dim sMach As String
    Dim sUser As String
    Dim rUser As UserRec

    On Error GoTo Err_WhosOn

    ' Haetaan tietokannan polku — DBEngine-viittausta EI suljeta (.Close kaataisi kannan)
    SPath = DBEngine.Workspaces(0).Databases(0).Name

    ' Muodostetaan lukitustiedoston polku (.LACCDB-pääte)
    SPath = Left$(SPath, InStr(1, SPath, ".")) & "LACCDB"

    ' Tarkistetaan, onko lukitustiedosto olemassa
    X = Dir$(SPath)
    If X = "" Then
        WhosOn = "Ei aktiivisia käyttäjiä"
        Exit Function
    End If

    ' Avataan lukitustiedosto binäärilukua varten
    iLDBFile = FreeFile
    Open SPath For Binary Access Read Shared As iLDBFile

    ' Luetaan kaikki käyttäjätietueet
    Do While Not EOF(iLDBFile)
        Get iLDBFile, , rUser

        ' Poimitaan tietokoneen nimi (null-päätetty merkkijono)
        With rUser
            i = 1
            sMach = ""
            While .bMach(i) <> 0 And i <= 32
                sMach = sMach & Chr$(.bMach(i))
                i = i + 1
            Wend

            ' Poimitaan käyttäjänimi (null-päätetty merkkijono)
            i = 1
            sUser = ""
            While .bUser(i) <> 0 And i <= 32
                sUser = sUser & Chr$(.bUser(i))
                i = i + 1
            Wend
        End With

        ' Lisätään tuloksiin (puolipisteellä eroteltu RowSource-arvo)
        If sLogins <> "" Then sLogins = sLogins & ";"
        sLogins = sLogins & sMach & ";" & sUser
    Loop

    Close iLDBFile
    WhosOn = sLogins
    Exit Function

Err_WhosOn:
    On Error Resume Next
    Close iLDBFile
    WhosOn = "Virhe käyttäjälistan luvussa: " & Err.Description
End Function

