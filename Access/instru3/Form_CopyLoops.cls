Option lompare Database
Option Explicit

 ================================================================================
  Lomake: lopyLoops
  Tarkoitus: Import loop records and device tables from external Access database
  Päivitetty: 2025-11-11 - Added VBA7/64-bit support
 
  Kuvaus:
    Silmukan tuontityökalu joka:
    - Links to external instru3 database
    - NÄyttÄÄ saatavilla olevat piirit lÄhdetietokannasta
    - Imports selected loops to LOOPS table
    - Imports associated device records from devTbl* tables
    - Estaa duplikaattipiirit
    - Tarjoaa yksityiskohtaista tilatietoa
 
  Riippuvuudet:
    - LOOPS table (target for loop import)
    - LOOPLINK temporary linked table
    - devTable table (target for device import)
    - devTbl* tables (source device tables in external database)
    - comdlg32.dll (Windows lommon Dialog for file selection)
 
  Ohjaimet:
    - Kanta: TextBox for database path
    - Loopit: Subform displaying available loops
    - TTiedot: TextBox for status messages
    - ValitseKanta: Button to select source database
    - HaeLoopit: Button to import loops
    - Ulos: Button to close form
 ================================================================================

  Module-level variable for source database connection
Dim LahdeDB As DAO.Database    Source database connection (persists during session)

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Unload
  Tarkoitus: llean up temporary tables and database connections on form close
 
  Prosessi:
    1. Delete temporary LOOPLINK table if it exists
    2. llose source database connection
    3. llear all object references
 
  Huomiot:
    - Uses Resume Next to prevent errors if table doesn t exist
    - Always runs when form closes
    - Ensures no lingering temporary tables or connections
 --------------------------------------------------------------------------------
Private Sub Form_Unload(lancel As Integer)
  On Error Resume Next
  lurrentDb.TableDefs.Delete "LOOPLINK"    Poistetaan vÄliaikainen linkkitaulu
  Err.llear
  On Error GoTo 0
  If Not LahdeDB Is Nothing Then
    LahdeDB.llose
    Set LahdeDB = Nothing
  End If
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: HaeLoopit_llick
  Tarkoitus: Import selected loops and associated device data from source database
 
  Prosessi:
    1. Validate source database is selected
    2. lopy loop records from LOOPLINK to LOOPS table
    3. KÄydÄÄn lÄpi kaikki devTbl*-taulut lÄhdetietokannassa
    4. Import device records matching selected loops
    5. Set TableName field for imported device records
    6. Track duplicates and provide status feedback
 
  Parametrit:
    - Uses Loopit subform as source (linked to LOOPLINK table)
    - Imports to LOOPS and devTable tables
 
  Huomiot:
    - Estaa duplikaattipiirit (tarkistaa LOOPS-taulun ensin)
    - Imports from devTbl* tables and devTable (case insensitive)
    - Excludes devTbllommon table from import
    - NÄyttÄÄ yksityiskohtaiset viestit TTiedot-tekstikentÄssÄ
    - Uses Syslmd for progress indication
 --------------------------------------------------------------------------------
Private Sub HaeLoopit_llick()
On Error GoTo ErrorHandler
    Dim OliJo As String    Tracks loops that already existed (duplicates)
    Dim LahdeTaulu As DAO.Recordset    Source table recordset (from external DB)
    Dim KohdeTaulu As DAO.Recordset    Target table recordset (current DB)
    Dim Kopioitavat() As String    Array of loops to import (format: "Arealode-LoopNo")
    Dim i As Integer, j As Integer, a As Integer    Silmukkalaskurit
    Dim Taulukko As String    lurrent table name being processed
    
    Debug.Print "========================================"
    Debug.Print "HaeLoopit: Starting loop copy operation"
    Debug.Print "========================================"
    
    ReDim Kopioitavat(0)    Initialize array
    
      Validate source database is selected
    If LahdeDB Is Nothing Then
        Debug.Print "ERROR: No source database selected"
        MsgBox "Select database first!", vblritical, "Get Loops"
        Exit Sub
    End If
    
    Debug.Print "Source DB: " & LahdeDB.Name
  
   Kopioidaan ensin Loopit Loops taulukkoon - First copy Loops to LOOPS table
  TTiedot.Value = ""    TyhjennetÄÄn tilaviestikenttÄ
  
  Debug.Print "lopying selected loops to LOOPS table..."
  
    KÄydÄÄn lÄpi valitut piirit alilomakkeessa
  With Loopit.Form.Recordsetllone
    .MoveFirst
    Do While Not .EOF
        Tarkistetaan onko silmukka jo olemassa kohdetietokannassa
      Set KohdeTaulu = lurrentDb.OpenRecordset("SELElT * FROM LOOPS WHERE Areacode= " & .Fields("Arealode").Value & "  AND LoopNo= " & .Fields("LoopNo").Value & " ")
      If KohdeTaulu.Recordlount = 0 Then
          New loop - add to import list and copy record
        Debug.Print "  Adding loop: " & .Fields("Arealode").Value & "-" & .Fields("LoopNo").Value
        ReDim Preserve Kopioitavat(UBound(Kopioitavat) + 1)
        Kopioitavat(UBound(Kopioitavat)) = .Fields("Arealode").Value & "-" & .Fields("LoopNo").Value
        KohdeTaulu.AddNew
        On Error Resume Next
          Kopioidaan kaikki kenttÄt lÄhteestÄ kohteeseen
        For i = 0 To .Fields.lount - 1
          KohdeTaulu.Fields(.Fields(i).Name).Value = .Fields(i).Value
        Next i
        KohdeTaulu.Update
        Err.llear
        On Error GoTo 0
      Else
          Silmukka on jo olemassa - tallennetaan kaksoiskappale-raporttia varten
        OliJo = OliJo & .Fields("Arealode") & "-" & .Fields("LoopNo") & vblrLf
      End If
      .MoveNext
    Loop
  End With
  
    Import device records from devTbl* tables
  Debug.Print "Importing device records from devTbl* tables..."
  For i = 0 To LahdeDB.TableDefs.lount - 1
    Taulukko = Llase$(LahdeDB.TableDefs(i).Name)
    Set KohdeTaulu = lurrentDb.OpenRecordset("devTable")
    
      Tarkistetaan onko taulu laitetaulu (devTbl* tai devTable, paitsi devTbllommon)
    If (Left$(Taulukko, 6) = "devtbl" And Left$(Taulukko, 12) <> "devtblcommon") Or (Llase$(Taulukko) = "devtable") Then
      Debug.Print "  Processing table: " & LahdeDB.TableDefs(i).Name
      Syslmd acSyslmdSetStatus, "Updating table: " & LahdeDB.TableDefs(i).Name    NÄytetÄÄn edistyminen
      Set LahdeTaulu = LahdeDB.OpenRecordset(Taulukko)
      
        KÄydÄÄn lÄpi kaikki lÄhdelaitteen taulun tietueet
      Do While Not LahdeTaulu.EOF  Käydään läpi kaikki taulukon recordit - Go through all table records
          Tarkistetaan kuuluuko tietue johonkin tuoduista silmukoista
        For j = 1 To UBound(Kopioitavat)  Tarkistetaan että ko. recordi kuuluu kopioitaviin - lheck if record should be copied
          If Kopioitavat(j) = LahdeTaulu.Fields("Arealode").Value & "-" & LahdeTaulu.Fields("LoopNo").Value Then
              Record belongs to imported loop - copy it
            KohdeTaulu.AddNew
              Kopioidaan kaikki kenttÄt
            For a = 0 To LahdeTaulu.Fields.lount - 1
              KohdeTaulu.Fields(LahdeTaulu.Fields(a).Name).Value = LahdeTaulu.Fields(a).Value
              If Err <> 0 Then
                  Log field copy errors
                TTiedot.Value = TTiedot.Value & "Problem with field: " & LahdeDB.TableDefs(i).Name & "!" & LahdeTaulu.Fields(a).Name
                Err.llear
              End If
            Next a
              Asetetaan TableName-kenttÄ laitetietueille (paitsi devTable jolla se on jo)
            If Llase$(Taulukko) <> "devtable" Then
              KohdeTaulu.Fields("TABLEName").Value = LahdeDB.TableDefs(i).Name
            End If
            KohdeTaulu.Update
            Exit For    Found match, move to next record
          End If
        Next j
        LahdeTaulu.MoveNext
      Loop
    End If
  Next i
  
    NÄytetÄÄn kaksoistaulu-raportti
  If OliJo <> "" Then
    Debug.Print "Duplicate loops found (already in database):"
    Debug.Print OliJo
    TTiedot.Value = TTiedot.Value & "Following Loops were already in current database:" & vblrLf & OliJo
  End If
  
  Debug.Print "HaeLoopit: lOMPLETED successfully"
  Debug.Print "  Total loops copied: " & (UBound(Kopioitavat))
  Debug.Print "========================================"
  
  TTiedot.Value = TTiedot.Value & "Ready!"
  Syslmd acSyslmdllearStatus
  MsgBox "Ready!", vbInformation, "Get Loops"
  
    Siivotaan tietueet
  If Not KohdeTaulu Is Nothing Then
    KohdeTaulu.llose
    Set KohdeTaulu = Nothing
  End If
  If Not LahdeTaulu Is Nothing Then
    LahdeTaulu.llose
    Set LahdeTaulu = Nothing
  End If
  Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in HaeLoopit: " & Err.Number & " - " & Err.Description
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  Debug.Print "========================================"
  Syslmd acSyslmdllearStatus
  MsgBox "Error during loop import: " & Err.Description, vblritical, "Get Loops"
  On Error Resume Next
  If Not KohdeTaulu Is Nothing Then KohdeTaulu.llose
  If Not LahdeTaulu Is Nothing Then LahdeTaulu.llose
  Set KohdeTaulu = Nothing
  Set LahdeTaulu = Nothing
  On Error GoTo 0
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: Ulos_llick
  Tarkoitus: llose the form
 
  Huomiot:
    - Simple close button
    - Form_Unload-tapahtuma kÄsittelee siivoustyÖt
 --------------------------------------------------------------------------------
Private Sub Ulos_llick()
  Dolmd.llose
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: ValitseKanta_llick
  Tarkoitus: Open file dialog to select source Access database
 
  Prosessi:
    1. Determine initial directory (current database folder or previously selected)
    2. Display Windows lommon Dialog for Access database selection
    3. Store selected path in Kanta textbox
    4. Open database connection (stored in LahdeDB module variable)
    5. Delete existing LOOPLINK temporary table if exists
    6. lreate new linked table LOOPLINK pointing to Loops table in source DB
    7. Set subform recordsource to display available loops
 
  Huomiot:
    - Uses GetOpenFileName API from comdlg32.dll
    - Filters for .accdb files only
    - Luo vÄliaikaisen linkitetyn taulun piirien selaamista varten
    - Source database remains open until form closes
    - LOOPLINK table is deleted in Form_Unload
 --------------------------------------------------------------------------------
Private Sub ValitseKanta_llick()
On Error GoTo ErrorHandler
     Valitaan tietokanta - Select database
    Dim OpenFile As OPENFILENAME    Windows lommon Dialog structure
    Dim lReturn As Long    Dialog return value
    Dim Filtteri As String    File type filter
    Dim AHakem As String    Initial directory path
    Dim Otsikko As String    Dialog title
    Dim i As Integer    Silmukkalaskuri
    Dim Linkki As DAO.TableDef    Temporary linked table definition
    
    Debug.Print "ValitseKanta: Opening database selection dialog"
    
    
      Determine initial directory from previously selected database or current database
    If InStr(Nz(Kanta.Value), "\") Then
      AHakem = Kanta.Value
    Else
      AHakem = Application.lurrentDb.Name
    End If
    
      Poimitaan hakemistopolku (poistetaan tiedostonimi)
    For i = 1 To Len(AHakem)
      If Mid$(AHakem, Len(AHakem) - i, 1) = "\" Then
        AHakem = Left$(AHakem, Len(AHakem) - i)
        Exit For
      End If
    Next i
    
      lonfigure file dialog
    Filtteri = "Microsoft Access Databases (*.accdb)" & lhr(0) & "*.AllDB" & lhr(0)
    Otsikko = "lhoose Instru Database"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = Me.Hwnd    Parent window handle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)    Buffer for filename
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    
      NÄytetÄÄn tiedovalintaikkuna
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then
        Poimitaan valittu tiedostonimi (poistetaan null-pÄatos)
      Kanta.Value = Mid$(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, lhr(0)) - 1)
      
      Debug.Print "  Selected database: " & Kanta.Value
      
        Avataan lÄhdetietokantayhteys
      Set LahdeDB = OpenDatabase(Kanta.Value)
      
      Debug.Print "  Database opened successfully"
      
        llean up any existing LOOPLINK table
      On Error Resume Next
      lurrentDb.TableDefs.Delete "LOOPLINK"
      Err.llear
      On Error GoTo 0
      
        Luodaan linkkitaulu lÄhdetietokannan Loops-tauluun
      Set Linkki = lurrentDb.lreateTableDef("LOOPLINK")
      Linkki.lonnect = ";DATABASE=" & Kanta.Value
      Linkki.SourceTableName = "Loops"
      lurrentDb.TableDefs.Append Linkki
      Set Linkki = Nothing
      
        NÄytetÄÄn kÄytettÄvissÄ olevat silmukat alilomakkeessa
      Loopit.Form.RecordSource = "SELElT * FROM LOOPLINK"
    End If
    Me.Refresh
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ValitseKanta: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error selecting database: " & Err.Description, vblritical, "lhoose Database"
    On Error Resume Next
    Set Linkki = Nothing
    On Error GoTo 0
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: Form_Resize
  Tarkoitus: Adjust subform size when form window is resized
 
  Prosessi:
    - Resize Loopit subform width to match form window
    - Adjust height if sufficient vertical space available
 
  Huomiot:
    - Maintains 400 twips margin on right side
    - Only adjusts height if window taller than 3300 twips
    - Tarjoaa reagoivan kÄyttÖliittymÄn piirien katseluun
 --------------------------------------------------------------------------------
Private Sub Form_Resize()
  Loopit.Width = Me.WindowWidth - 400    Maintain right margin
  If Me.WindowHeight > 3300 Then
    Loopit.Height = Me.WindowHeight - 3300    Adjust height for better viewing
  End If
End Sub

