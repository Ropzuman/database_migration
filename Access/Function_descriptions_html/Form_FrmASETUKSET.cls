Option Compare Database
Option Explicit

'================================================================================
' Lomake: FrmASETUKSET (Asetukset)
' Tarkoitus: Tietokannan konfigurointi ja taulujen linkitys
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — Me.-etuliite lisätty kaikille lomakeohjaimille
'
' Kuvaus:
'   Asetukset-lomake ulkoisten Access-tietokantojen taulujen linkittämiseen.
'   Linkittää taulut tietokannoista instru3.accdb, MAINEQ.accdb ja
'   FunctionDiagrams.accdb.
'
' Riippuvuudet:
'   - DAO.TableDef
'   - Ulkoiset tietokannat: instru3.accdb, MAINEQ.accdb, FunctionDiagrams.accdb
'   - Lomakkeet: frmMUOKKAUS
'   - Moduuli: USysCheck (SniffUser-funktio)
'   - Ohjaimet: TINSTRU, TMAINEQ, TFUNCDIAG (polkujen tekstikentät)
'================================================================================

'================================================================================
' Tapahtuma: Command11_Click
' Tarkoitus: Avaa päämuokkauslomakkeen ja sulkee asetukset
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
' Tarkoitus: Linkittää kaikki tarvittavat taulut ulkoisista tietokannoista
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
    
    ' Toimintakaaviot (toistaiseksi pois käytöstä)
    ' Linkkaa "CONTROL", Me.TFUNCDIAG.Value
    
    MsgBox "Taulukot linkattu", vbOKOnly, "Linkkaa taulukot"
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe taulujen linkityksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: Linkkaa
' Tarkoitus: Linkittää ulkoisen taulun nykyiseen tietokantaan
' Parametrit:
'   Taulukko — Linkitettävän taulun nimi
'   KANTA    — Ulkoisen Access-tietokannan täydellinen polku
'
' Kuvaus:
'   Poistaa olemassa olevan taulun ensin (jos löytyy), luo sitten uuden
'   linkitetyn taulumäärityksen, joka osoittaa ulkoiseen tietokantaan.
'================================================================================
Private Sub Linkkaa(Taulukko As String, KANTA As String)
    Dim tdf As DAO.TableDef
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Poistetaan olemassa oleva taulu ennen uudelleenlinkitystä
    For i = 0 To CurrentDb.TableDefs.Count - 1
        If LCase$(CurrentDb.TableDefs(i).Name) = LCase$(Taulukko) Then
            CurrentDb.Execute "DROP TABLE [" & Taulukko & "]"
            Exit For
        End If
    Next i
    
    ' Luodaan linkitetty taulumääritys
    Set tdf = CurrentDb.CreateTableDef(Taulukko)
    With tdf
        .Connect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    CurrentDb.TableDefs.Append tdf
    CurrentDb.TableDefs.Refresh
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe taulun '" & Taulukko & "' linkityksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Form_Load
' Tarkoitus: Alustaa lomake oletuspolkiilla ja kirjaa käyttäjän
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
    
    ' Kirjataan käyttäjä lokiin
    Call SniffUser
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe lomakkeen latauksessa: " & Err.Description, vbExclamation
End Sub
