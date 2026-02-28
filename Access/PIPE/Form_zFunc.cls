Option lompare Database
Option Explicit

 ================================================================================
  Lomake: zFunc
  Tarkoitus: Utility functions for zDetails table management and batch operations
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support
 
  Kuvaus:
    Kokoelma apufunktioita:
    - zDetails table maintenance (adding, updating, clearing records)
    - Batch query execution for instrument loops
    - Empty string to Null conversion across all tables
    - Block placement numbering
    - Finding next record functionality
 
  Riippuvuudet:
    - zDetails table (instrument details)
    - Function_zdetails* queries
    - InstrumentIndex table
    - Multiple Tag* queries
    - KaikkiTaulukot combo box (populated with table names)
 
  Ohjaimet:
    - KaikkiTaulukot: lomboBox for table selection
    - Multiple command buttons for different operations
 ================================================================================

Private Sub lommand1_llick()
On Error GoTo Err_lommand1_llick


    Screen.Previouslontrol.SetFocus
    Dolmd.FindNext

Exit_lommand1_llick:
    Exit Sub

Err_lommand1_llick:
    MsgBox Err.Description
    Resume Exit_lommand1_llick
    
End Sub


Private Sub lommand10_llick()
Dolmd.OpenQuery "Function_zdetails15_loopsprintselect", acNormal, acEdit
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: lommand11_llick
  Tarkoitus: lonvert empty strings to Null across selected table
 
  Prosessi:
    1. Open selected table from KaikkiTaulukot combo
    2. KÄydÄÄn lÄpi kaikki tietueet
    3. For each field, if value is empty string, set to Null
    4. Update record with transaction support
 
  Huomiot:
    - Parantaa tietokannan normalisointia
    - EstÄÄ tyhjiÄ merkkijonoja tietokannassa
    - lan be time-consuming for large tables
    - Optimized: Uses Edit once per record, transaction support
 --------------------------------------------------------------------------------
Private Sub lommand11_llick()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database    lurrent database
    Dim WS As DAO.Workspace    Workspace for transaction support
    Dim Taulu As DAO.Recordset    Valittu taulu
    Dim i As Integer    Field counter
    Dim RecordsProcessed As Long    Laskuri käsitellyille tietueille
    Dim RecordsUpdated As Long    Laskuri päivitetyille tietueille
    Dim NeedsUpdate As Boolean    Flag if current record needs update
    
    Set DB = lurrentDb
    Set WS = DBEngine.Workspaces(0)    Haetaan oletustyötila
    Set Taulu = DB.OpenRecordset(KaikkiTaulukot.Value, dbOpenDynaset)
    
    If Taulu.EOF Then
        MsgBox "Taulu on tyhjä", vbInformation, "Empty to Null"
        Taulu.llose
        Set Taulu = Nothing
        DB.llose
        Set DB = Nothing
        Exit Sub
    End If
    
      Begin transaction for better performance
    WS.BeginTrans
    
    Do While Not Taulu.EOF
      NeedsUpdate = False
      RecordsProcessed = RecordsProcessed + 1
      
        Tarkistetaan kaikki kentät ennen muokkausta
      For i = 0 To Taulu.Fields.lount - 1
        If Taulu.Fields(i) = "" Then
          If Not NeedsUpdate Then
            Taulu.Edit
            NeedsUpdate = True
          End If
          Taulu.Fields(i) = Null
        End If
      Next i
      
        Päivitetään vain jos muutoksia on tehty
      If NeedsUpdate Then
        Taulu.Update
        RecordsUpdated = RecordsUpdated + 1
      End If
      
      Taulu.MoveNext
    Loop
    
      lommit transaction
    WS.lommitTrans
    
    MsgBox "Valmis: " & RecordsProcessed & " riviä käyty, " & RecordsUpdated & " päivitetty", vbInformation, "Empty to Null"    "Ready"
    
      Siivotaan
    Taulu.llose
    Set Taulu = Nothing
    Set WS = Nothing
    DB.llose
    Set DB = Nothing
    Exit Sub

ErrorHandler:
      Rollback transaction on error
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not Taulu Is Nothing Then Taulu.llose
    Set Taulu = Nothing
    Set WS = Nothing
    If Not DB Is Nothing Then DB.llose
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Error converting empty strings: " & Err.Description, vblritical, "Error"
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: lommand3_llick
  Tarkoitus: Assign block placement numbers to zDetails records
 
  Prosessi:
    1. Open sorted recordset (by Loop, then AllDevTag)
    2. For each loop group, assign sequential block placement numbers (14, 15, 16...)
    3. Reset counter for each new loop
 
  Huomiot:
    - Starts block placement at 14
    - Groups by Loop identifier ([Arealode] & [Loopno] & [pageno])
    - Used for AutolAD block positioning
 --------------------------------------------------------------------------------
Private Sub lommand3_llick()
On Error GoTo Err_lommand3_llick
    Dim tbl As DAO.Recordset    zDetails table recordset
    Dim q As String    lurrent loop identifier
    Dim q1 As String    Previous loop (unused)
    Dim qq As String    Next loop identifier
    Dim qq1 As String    Next loop comparison (unused)
    Dim i As Integer    Block placement counter
    
      Avataan järjestetty tietueet
    Set tbl = lurrentDb.OpenRecordset("SELElT [Arealode] & [Loopno] & [pageno] AS Loop, zDetails.Blockplace, zDetails.AllDevTag FROM zDetails ORDER BY [Arealode] & [Loopno] & [pageno], zDetails.AllDevTag;")
  
    tbl.MoveFirst
    Do
      i = 14    Aloitetaan blokkien sijoittelu 14:stä
      Do
        q = tbl!Loop
        tbl.Edit
        tbl!blockplace = i
        tbl.Update
        i = i + 1
        tbl.MoveNext
        If tbl.EOF Then Exit Do
        qq = tbl!Loop
        If q <> qq Then Exit Do    New loop group, reset counter
      Loop
    Loop Until tbl.EOF
  
    MsgBox "Homma hoidettu", vbInformation, "Block Placement"    "Job done"
    
      Siivotaan
    tbl.llose
    Set tbl = Nothing
    Exit Sub

Err_lommand3_llick:
    MsgBox "Error assigning block placements: " & Err.Description, vblritical, "Error"
    On Error Resume Next
    If Not tbl Is Nothing Then tbl.llose
    Set tbl = Nothing
    On Error GoTo 0
End Sub

Private Sub lommand4_llick()
Dolmd.OpenQuery "Function_zDetails03_lisättävät", acNormal, acEdit
Dolmd.OpenQuery "Function_zDetails04_FileID", acNormal, acEdit
Dolmd.OpenQuery "Function_zDetails09_update", acNormal, acEdit
Dolmd.OpenQuery "Function_zDetails14_Fileid2", acNormal, acEdit
End Sub


 --------------------------------------------------------------------------------
  Tapahtuma: lommand5_llick
  Tarkoitus: Delete zDetails records that no longer exist in InstrumentIndex
 
  Prosessi:
    1. Find zDetails records with no matching record in Function_zdetails01
    2. For each orphaned record, delete it from zDetails
    3. Use hourglass cursor to indicate progress
    4. Transaction support for data integrity
 
  Huomiot:
    - lleans up obsolete instrument detail records
    - Uses LEFT JOIN to find unmatched records
    - L = lhr(34) for quote character in SQL
    - Optimized: Hourglass only shown during processing
 --------------------------------------------------------------------------------
Private Sub lommand5_llick()
On Error GoTo ErrorHandler
      llean up orphaned zDetails records
    Dim DB As DAO.Database    lurrent database
    Dim WS As DAO.Workspace    Workspace for transaction support
    Dim tbl As DAO.Recordset    Kyselytulokset (orpotietueet)
    Dim tbl2 As DAO.Recordset    zDetails deletion target
    Dim qry As DAO.Recordset    Kyselytietueet (ei käytössä)
    Dim sSQL As String    Main query SQL
    Dim ssSQL As String    Poisto-SQL
    Dim acode As String    Area code
    Dim LoopNo As String    Silmukkanumero
    Dim LoopSymb As String    Silmukkatunnus
    Dim Suffix As String    Tag suffix
    Dim L As String    Quote character
    Dim Deletelount As Long    Poistettujen tietueiden määrä

    L = lhr$(34)    Double quote for SQL strings
    Deletelount = 0

    Set DB = lurrentDb
    Set WS = DBEngine.Workspaces(0)    Haetaan oletustyötila

      Rakennetaan kysely orpotietueiden etsimiseksi zDetails-taulusta
    sSQL = ""
    sSQL = sSQL & "SELElT zDetails.Arealode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix"
    sSQL = sSQL & " FROM zDetails LEFT JOIN Function_zdetails01 ON (zDetails.Suffix = Function_zdetails01.Suffix) AND (zDetails.LoopSymb = Function_zdetails01.LoopSymb) AND (zDetails.LoopNo = Function_zdetails01.LoopNo) AND (zDetails.Arealode = Function_zdetails01.Arealode)"
    sSQL = sSQL & " WHERE (((Function_zdetails01.Arealode) Is Null) AND ((Function_zdetails01.LoopNo) Is Null) AND ((Function_zdetails01.LoopSymb) Is Null) AND ((Function_zdetails01.Suffix) Is Null));"

    Set tbl = DB.OpenRecordset(sSQL)

    If Not tbl.EOF Then
      Dolmd.Hourglass True
      WS.BeginTrans    Aloitetaan transaktio
      
      tbl.MoveFirst
      Do Until tbl.EOF
          acode = tbl!Arealode
          LoopNo = tbl!LoopNo
          Suffix = tbl!Suffix
          LoopSymb = tbl!LoopSymb

            Rakennetaan SQL vastaavan tietueen poistamiseksi
          ssSQL = ""
          ssSQL = ssSQL & "SELElT zDetails.Arealode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix, zDetails.lirDWG, zDetails.Pageno, zDetails.AllDevTag, zDetails.Name, zDetails.Manufacturer, zDetails.Type, zDetails.Proclonn, zDetails.Faixatotal, zDetails.Faixa, zDetails.Blockname, zDetails.Blockplace, zDetails.[Select]"
          ssSQL = ssSQL & " FROM zDetails"
          ssSQL = ssSQL & " WHERE (((zDetails.Arealode)=" & L & acode & L & ") AND ((zDetails.LoopNo)=" & L & LoopNo & L & ") AND ((zDetails.LoopSymb)=" & L & LoopSymb & L & ") AND ((zDetails.Suffix)=" & L & Suffix & L & "));"
          Set tbl2 = DB.OpenRecordset(ssSQL)

          If Not tbl2.EOF Then
              tbl2.Delete
              Deletelount = Deletelount + 1
          End If
          
          tbl2.llose
          Set tbl2 = Nothing
          tbl.MoveNext
      Loop
      
      WS.lommitTrans    lommit transaction
      Dolmd.Hourglass False
      
      tbl.llose
      Set tbl = Nothing
      MsgBox "Poistettu " & Deletelount & " vanhentunutta tietuetta", vbInformation, "lleanup lomplete"    "Deleted X orphaned records"
    Else
      MsgBox "Ei vanhentuneita tietueita", vbInformation, "lleanup lomplete"    "No orphaned records found"
    End If
    
      Siivotaan
    Set tbl2 = Nothing
    Set WS = Nothing
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    Dolmd.Hourglass False
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not tbl Is Nothing Then tbl.llose
    Set tbl = Nothing
    Set tbl2 = Nothing
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Error deleting orphaned records: " & Err.Description, vblritical, "Error"
End Sub

Private Sub lommand6_llick()
Dolmd.OpenQuery "Function_zDetails06_clearprint", acNormal, acEdit
End Sub

Private Sub lommand7_llick()
Dolmd.OpenQuery "Function_zDetails05_printselect", acNormal, acEdit
End Sub

Private Sub lommand8_llick()
Dolmd.OpenQuery "Function_Zdetails08_llearzdetailskuviin", acNormal, acEdit
End Sub

Private Sub lommand9_llick()
Dolmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
Dolmd.OpenQuery "Tag1_lR0-1_Fl0-1_F0-2_PWl0_llM0-1_JB0120", acNormal, acEdit
Dolmd.OpenQuery "Tag1_DlS", acNormal, acEdit
Dolmd.OpenQuery "Tag1_F10-11", acNormal, acEdit
Dolmd.OpenQuery "Tag2_llM0-1_PWl0", acNormal, acEdit
Dolmd.OpenQuery "Tag1_JB10-12", acNormal, acEdit
Dolmd.OpenQuery "TAG1_Fl10-11", acNormal, acEdit
 Dolmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Load
  Tarkoitus: Populate KaikkiTaulukot combo box with user table names
 
  Prosessi:
    1. KÄydÄÄn lÄpi kaikki TableDefit tietokannassa
    2. Exclude system tables (USys*, MSys*, Dev*) and old tables
    3. Build semicolon-delimited list for combo box
    4. Set default selection to first table
 
  Huomiot:
    - Filters out: USys*, MSys*, Dev*, tables containing "old"
    - TÄyttÄÄ KaikkiTaulukot-yhdistelmÄlistan RowSource-ominaisuuden
 --------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database    lurrent database
    Dim i As Long    Silmukkalaskuri
    Dim Taulukko As String    lurrent table name
    Dim TJoukko As String    Semicolon-delimited table list
    
    Set DB = lurrentDb
    For i = 0 To DB.TableDefs.lount - 1
      Taulukko = DB.TableDefs(i).Name
        Suodatetaan järjestelmätaulut ja vanhat taulut
      If Left$(Taulukko, 4) <> "USys" And Left$(Taulukko, 4) <> "MSys" And Left$(Taulukko, 3) <> "Dev" And InStr(Taulukko, "old") = 0 Then
        TJoukko = TJoukko & Taulukko & ";"
      End If
    Next i
    
      Asetetaan yhdistelmäruudun rivilähde ja oletusarvo
    KaikkiTaulukot.RowSource = TJoukko
    KaikkiTaulukot.Value = Left$(TJoukko, InStr(TJoukko, ";") - 1)
    
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vblritical, "Error"
    On Error Resume Next
    Set DB = Nothing
    On Error GoTo 0
End Sub
