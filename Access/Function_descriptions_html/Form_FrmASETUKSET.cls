Option Compare Database
Option Explicit

'================================================================================
' Lomake: FrmASETUKSET (Asetukset)
' Tarkoitus: Tietokannan konfigurointi ja taulujen linkitys
' P‰ivitetty: 2025-11-13 ó VBA7/64-bit tuki lis‰tty
'             2026-03-03 ó Me.-etuliite lis‰tty kaikille lomakeohjaimille
'
' Kuvaus:
'   Asetukset-lomake ulkoisten Access-tietokantojen taulujen linkitt‰miseen.
'   Linkitt‰‰ taulut tietokannoista instru3.accdb, MAINEQ.accdb ja
'   FunctionDiagrams.accdb.
'
' Riippuvuudet:
'   - DAO.TableDef
'   - Ulkoiset tietokannat: instru3.accdb, MAINEQ.accdb, FunctionDiagrams.accdb
'   - Lomakkeet: frmMUOKKAUS
'   - Moduuli: USysCheck (SniffUser-funktio)
'   - Ohjaimet: TINSTRU, TMAINEQ, TFUNCDIAG (polkujen tekstikent‰t)
'================================================================================

'================================================================================
' Tapahtuma: Command11_Click
' Tarkoitus: Avaa p‰‰muokkauslomakkeen ja sulkee asetukset
'================================================================================
Private Sub Command11_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmMUOKKAUS"
    DoCmd.Close , "FrmASETUKSET"
    Exit Sub
ErrorHandler:
    MsgBox "Virhe lomakkeen avauksessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command2_Click
' Tarkoitus: Linkitt‰‰ kaikki tarvittavat taulut ulkoisista tietokannoista
'================================================================================
Private Sub Command2_Click()
    On Error GoTo ErrorHandler
    
    ' Instrumenttitietokannan taulut
    Linkkaa "Alarms", Me.TINSTRU.Value
    Linkkaa "Loops", Me.TINSTRU.Value
    Linkkaa "DevTable", Me.TINSTRU.Value
    Linkkaa "DevSpec", Me.TINSTRU.Value
    Linkkaa "TAGID_exception", Me.TINSTRU.Value
    
    ' Laitteistotietokannan taulut
    Linkkaa "MAINEQ", Me.TMAINEQ.Value
    Linkkaa "DRIVES", Me.TMAINEQ.Value
    
    ' Toimintakaaviot (toistaiseksi pois k‰ytˆst‰)
    ' Linkkaa "CONTROL", Me.TFUNCDIAG.Value
    
    MsgBox "Taulukot linkattu", vbOKOnly, "Linkkaa taulukot"
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe taulujen linkityksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: Linkkaa
' Tarkoitus: Linkitt‰‰ ulkoisen taulun nykyiseen tietokantaan
' Parametrit:
'   Taulukko ó Linkitett‰v‰n taulun nimi
'   KANTA    ó Ulkoisen Access-tietokannan t‰ydellinen polku
'
' Kuvaus:
'   Poistaa olemassa olevan taulun ensin (jos lˆytyy), luo sitten uuden
'   linkitetyn taulum‰‰rityksen, joka osoittaa ulkoiseen tietokantaan.
'================================================================================
Private Sub Linkkaa(Taulukko As String, KANTA As String)
    Dim tdf As DAO.TableDef
    Dim db As DAO.Database
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' V‰limuistitetaan CurrentDb ó v‰ltet‰‰n toistuva uuden instanssin luonti per kutsu
    Set db = CurrentDb
    
    ' Poistetaan olemassa oleva taulu ennen uudelleenlinkityst‰
    For i = 0 To db.TableDefs.Count - 1
        If LCase$(db.TableDefs(i).Name) = LCase$(Taulukko) Then
            db.Execute "DROP TABLE [" & Taulukko & "]"
            Exit For
        End If
    Next i
    
    ' Luodaan linkitetty taulum‰‰ritys
    Set tdf = db.CreateTableDef(Taulukko)
    With tdf
        .Connect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
    
    ' Vapautetaan DAO-resurssit
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Set tdf = Nothing
    Set db = Nothing
    MsgBox "Virhe taulun '" & Taulukko & "' linkityksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Form_Load
' Tarkoitus: Alustaa lomake oletuspolkiilla ja kirjaa k‰ytt‰j‰n
'================================================================================
Private Sub Form_Load()
    Dim POLKU As String
    
    On Error GoTo ErrorHandler
    
    ' Haetaan nykyisen tietokannan hakemisto
    POLKU = CurrentDb.Name
    POLKU = Left$(POLKU, InStrRev(POLKU, "\"))
    
    ' Asetetaan ulkoisten tietokantojen oletuspolut
    Me.TINSTRU.Value = POLKU & "instru3.accdb"
    Me.TMAINEQ.Value = POLKU & "MAINEQ.accdb"
    Me.TFUNCDIAG.Value = POLKU & "FunctionDiagrams.accdb"
    
    ' Kirjataan k‰ytt‰j‰ lokiin
    Call SniffUser
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe lomakkeen latauksessa: " & Err.Description, vbExclamation
End Sub
