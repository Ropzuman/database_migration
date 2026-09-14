Option Compare Database
Option Explicit

'================================================================================
' Lomake: DBUsers
' Tarkoitus: N‰ytt‰‰ tietokantaan t‰ll‰ hetkell‰ kirjautuneet k‰ytt‰j‰t
' P‰ivitetty: 2025-11-13 ó VBA7/64-bit tuki lis‰tty
'             2026-03-03 ó Korjattu dbCurrent.Close-virhe; poistettu orvokoodi
'
' Kuvaus:
'   Lukee .LACCDB-lukitustiedoston, jotta n‰hd‰‰n ketk‰ k‰ytt‰v‰t tietokantaa.
'   N‰ytt‰‰ aseman nimen ja kirjautuneen k‰ytt‰j‰n jokaisesta aktiivisesta
'   yhteydest‰.
'
' Riippuvuudet:
'   - .LACCDB-lukitustiedosto (Access 2007+)
'   - LoggedOn-ohjain (lista-/yhdistelm‰ruutu)
'
' Lukitustiedoston rakenne:
'   - 64 bitin tietueet
'   - Tavut 1ñ32: Tietokoneen nimi (null-p‰‰tetty merkkijono)
'   - Tavut 33ñ64: K‰ytt‰j‰n kirjautumisnimi (null-p‰‰tetty merkkijono)
'================================================================================

' K‰ytt‰j‰tietue lukitustiedoston lukemiseen
Private Type UserRec
    bMach(1 To 32) As Byte  ' Tietokoneen nimi (ensimm‰iset 32 tavua)
    bUser(1 To 32) As Byte  ' K‰ytt‰j‰nimi (j‰lkimm‰iset 32 tavua)
End Type

'================================================================================
' Tapahtuma: Form_Open
' Tarkoitus: Alustaa lomake nykyisell‰ k‰ytt‰j‰listalla lomakkeen avautuessa
'================================================================================
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Exit Sub
ErrorHandler:
    MsgBox "Virhe k‰ytt‰j‰listan latauksessa: " & Err.Description, vbExclamation
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
    ' Hiljainen virhe ó lomake suljetaan jo todenn‰kˆisesti muualta
End Sub

'================================================================================
' Tapahtuma: UpdateBtn_Click
' Tarkoitus: P‰ivitt‰‰ k‰ytt‰j‰listan
'================================================================================
Private Sub UpdateBtn_Click()
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub
ErrorHandler:
    MsgBox "Virhe k‰ytt‰j‰listan p‰ivityksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: WhosOn
' Tarkoitus: Lukee lukitustiedoston ja koostaa kirjautuneiden k‰ytt‰jien listan
' Palauttaa: Puolipisteell‰ eroteltu "Kone;K‰ytt‰j‰"-lista (RowSource-arvo)
'
' Kuvaus:
'   Avaa .LACCDB-lukitustiedoston ja lukee kaikki 64 tavun k‰ytt‰j‰tietueet.
'   Poimii tietokoneen nimen ja k‰ytt‰j‰nimen kustakin tietueesta sek‰
'   rakentaa arvoluettelon, joka sopii listalaatikon RowSource-ominaisuudeksi.
'
' Tiedoston sijainti:
'   Sama hakemisto kuin tietokanta, sama nimi mutta .LACCDB-p‰‰te
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

    ' Haetaan tietokannan polku ó DBEngine-viittausta EI suljeta (.Close kaataisi kannan)
    SPath = DBEngine.Workspaces(0).Databases(0).Name

    ' Muodostetaan lukitustiedoston polku (.LACCDB-p‰‰te)
    SPath = Left$(SPath, InStr(1, SPath, ".")) & "LACCDB"

    ' Tarkistetaan, onko lukitustiedosto olemassa
    X = Dir$(SPath)
    If X = "" Then
        WhosOn = "Ei aktiivisia k‰ytt‰ji‰"
        Exit Function
    End If

    ' Avataan lukitustiedosto bin‰‰rilukua varten
    iLDBFile = FreeFile
    Open SPath For Binary Access Read Shared As iLDBFile

    ' Luetaan kaikki k‰ytt‰j‰tietueet
    Do While Not EOF(iLDBFile)
        Get iLDBFile, , rUser

        ' Poimitaan tietokoneen nimi (null-p‰‰tetty merkkijono)
        With rUser
            i = 1
            sMach = ""
            While .bMach(i) <> 0 And i <= 32
                sMach = sMach & Chr$(.bMach(i))
                i = i + 1
            Wend

            ' Poimitaan k‰ytt‰j‰nimi (null-p‰‰tetty merkkijono)
            i = 1
            sUser = ""
            While .bUser(i) <> 0 And i <= 32
                sUser = sUser & Chr$(.bUser(i))
                i = i + 1
            Wend
        End With

        ' Lis‰t‰‰n tuloksiin (puolipisteell‰ eroteltu RowSource-arvo)
        If sLogins <> "" Then sLogins = sLogins & ";"
        sLogins = sLogins & sMach & ";" & sUser
    Loop

    Close iLDBFile
    WhosOn = sLogins
    Exit Function

Err_WhosOn:
    On Error Resume Next
    Close iLDBFile
    WhosOn = "Virhe k‰ytt‰j‰listan luvussa: " & Err.Description
End Function

