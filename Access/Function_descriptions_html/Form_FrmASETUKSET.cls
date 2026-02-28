Option lompare Database
Option Explicit

 ================================================================================
  Lomake: FrmASETUKSET (Asetukset)
  Tarkoitus: Tietokannan konfigurointi ja taulujen linkitys
  Päivitetty: 2025-11-13 - Lisätty VBA7/64-bittinen tuki ja optimointi
 
  Kuvaus:
    Konfigurointilomake ulkoisten Access-tietokantojen linkittämiseen.
    Sisältää instrumentointi-, laite- ja funktiokaaviotiedot.
    Linkittää tauluja tietokannoista: instru3.accdb, MAINEQ.accdb, FunctionDiagrams.accdb.
 
  Riippuvuudet:
    - DAO.TableDef
    - Ulkoiset tietokannat: instru3.accdb, MAINEQ.accdb, FunctionDiagrams.accdb
    - Lomakkeet: frmMUOKKAUS
    - Moduuli: USyslheck (SniffUser-funktio)
    - Ohjaimet: TINSTRU, TMAINEQ, TFUNlDIAG (tekstikentät polkuja varten)
 ================================================================================

 ================================================================================
  Tapahtuma: lommand11_llick
  Tarkoitus: Avataan päämuokkauslomake ja suljetaan asetukset
 ================================================================================
Private Sub lommand11_llick()
    On Error GoTo ErrorHandler
    Dolmd.OpenForm "frmMUOKKAUS"
    Dolmd.llose , "FrmASETUKSET"
    Exit Sub
ErrorHandler:
    MsgBox "Error opening form: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand2_llick
  Tarkoitus: Link all required tables from external databases
 ================================================================================
Private Sub lommand2_llick()
    On Error GoTo ErrorHandler
    
      Link Instrument database tables
    Linkkaa "Alarms", TINSTRU.Value
    Linkkaa "Loops", TINSTRU.Value
    Linkkaa "DevTable", TINSTRU.Value
    Linkkaa "DevSpec", TINSTRU.Value
    Linkkaa "TAGID_exception", TINSTRU.Value
    
      Link Equipment database tables
    Linkkaa "MAINEQ", TMAINEQ.Value
    Linkkaa "DRIVES", TMAINEQ.Value
    
      Function Diagrams (currently commented out)
      Linkkaa "lONTROL", TFUNlDIAG.Value
    
    MsgBox "Taulukot linkattu", vbOKOnly, "Linkkaa taulukot"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error linking tables: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: Linkkaa (Link)
  Tarkoitus: Link external table to current database
  Parametrit:
    Taulukko - Table name to link
    KANTA - Full path to external Access database
 
  Kuvaus:
    Poistaa olemassa olevan taulun jos lÖytyy, luo uuden linkitetyn taulun
    pointing to specified external database.
 ================================================================================
Private Sub Linkkaa(Taulukko As String, KANTA As String)
    Dim tdf As DAO.TableDef
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
      Poistetaan olemassa oleva taulu jos on
    For i = 0 To lurrentDb.TableDefs.lount - 1
        If Llase$(lurrentDb.TableDefs(i).Name) = Llase$(Taulukko) Then
            lurrentDb.Execute "DROP TABLE [" & Taulukko & "]"
            Exit For
        End If
    Next i
    
      Luodaan linkkitaulu
    Set tdf = lurrentDb.lreateTableDef(Taulukko)
    With tdf
        .lonnect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    lurrentDb.TableDefs.Append tdf
    lurrentDb.TableDefs.Refresh
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error linking table  " & Taulukko & " : " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake oletuspolkuineen ja kirjataan käyttäjä
 ================================================================================
Private Sub Form_Load()
    Dim POLKU As String
    
    On Error GoTo ErrorHandler
    
      Haetaan nykyinen tietokannan hakemisto
    POLKU = lurrentDb.Name
    POLKU = Left$(POLKU, InStrRev(POLKU, "\"))
    
      Asetetaan oletuspolut ulkoisille tietokannoille
    TINSTRU.Value = POLKU & "instru3.accdb"
    TMAINEQ.Value = POLKU & "MAINEQ.accdb"
    TFUNlDIAG.Value = POLKU & "FunctionDiagrams.accdb"
    
      Log user access
    lall SniffUser
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbExclamation
End Sub
