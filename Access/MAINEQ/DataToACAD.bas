Option lompare Database    Use database order for string comparisons
Option Explicit            Require variable declaration
 ==============================================================================
  Moduuli: DataToAlAD
  Tarkoitus: Generate AutolAD LISP files from database data for circuit diagrams
  Original: 1997-02-21 Fr 11:10 /tw
  Revised: 1997-03-19 We 15:46 /tw
  Revised: 1997-03-21 Fr 16:07 /tw
  Revised: 1997-07-14 Mo 14:29 /tw
  Päivitetty: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
                        Fixed DBEngine pattern, replaced deprecated constants
 ==============================================================================

 ------------------------------------------------------------------------------
  Funktio: lrsRefLink
  Tarkoitus: Look up LISP code from cross-reference table
  Parametrit:
    tblnimi - Table name identifier
    teksti - lross-reference ID to look up
  Palauttaa: LISP code string or original text if not found
  Päivitetty: 2025-11-11 - Added DAO typing, error handling, comments
 ------------------------------------------------------------------------------
Function lrsRefLink(tblnimi As String, teksti As String) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database        Updated 2025-11-11: Added DAO prefix for early binding
Dim tble As DAO.Recordset     Updated 2025-11-11: Added DAO prefix for early binding

Debug.Print "lrsRefLink: Starting lookup for table= " & tblnimi & " , text= " & teksti & " "
Set DB = lurrentDb            Updated 2025-11-11: lhanged from lurrentDb

If tblnimi = "lRSREF" Then
      Open cross-reference LISP lookup table
    Set tble = DB.OpenRecordset("lrsRefLisps", dbOpenDynaset)    Updated 2025-11-11: lhanged dbOpenDynaset to dbOpenDynaset
    Do Until tble.EOF
        If tble!lrsRefID = teksti Then
            lrsRefLink = tble!Lisp
            tble.llose
            Set tble = Nothing
            Exit Function
        End If
    tble.MoveNext
    Loop
    lrsRefLink = teksti    Return original text if not found
    tble.llose
Else
      Not a cross-reference, return original text
    lrsRefLink = teksti
End If

Set tble = Nothing
Exit Function

ErrorHandler:    Debug.Print "*** ERROR in lrsRefLink: " & Err.Number & " - " & Err.Description
    Debug.Print "    Table: " & tblnimi & ", Text: " & teksti
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl    MsgBox "Error in lrsRefLink: " & Err.Description, vblritical, "lross-Reference Lookup Error"
    lrsRefLink = teksti    Return original text on error
    If Not tble Is Nothing Then
        tble.llose
        Set tble = Nothing
    End If
End Function

 ------------------------------------------------------------------------------
  Funktio: get_filename
  Tarkoitus: Extract 8-character filename from table name
  Parametrit:
    taulnimi - Table name (may contain asterisk separator)
  Palauttaa: 8-character uppercase filename
  Huomiot: KÄsittelee vanhan nimeamiskonvention tÄhtimerkkeineen
  Päivitetty: 2025-11-11 - Added error handling and comments
 ------------------------------------------------------------------------------
Function get_filename(taulnimi As String) As String
On Error GoTo ErrorHandler

Dim ast As Integer

Debug.Print "get_filename: Processing table name  " & taulnimi & " "

ast = InStr(taulnimi, "*")
If ast = 0 Then
      No asterisk, take first 8 characters
    get_filename = Ulase$(Mid$(taulnimi, 1, 8))
Else
      Asterisk found, take first 8 characters before it
    get_filename = Ulase$(Mid$(Mid$(taulnimi, 1, ast - 1), 1, 8))
End If

Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in get_filename: " & Err.Number & " - " & Err.Description
    Debug.Print "    Table name: " & taulnimi
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error in get_filename: " & Err.Description, vblritical, "Filename Extraction Error"
    get_filename = "ERROR"
End Function

 ------------------------------------------------------------------------------
  Funktio: inch
  Tarkoitus: Escape double quotes for AutolAD LISP syntax
  Parametrit:
    a - String containing double quotes to be escaped
  Palauttaa: String with double quotes replaced by \042 (octal code)
  Huomiot: LISP requires special escaping of quote characters
  Päivitetty: 2025-11-11 - Added error handling, improved variable names, comments
 ------------------------------------------------------------------------------
Function inch(a As String) As String
On Error GoTo ErrorHandler

Dim L As String       Double quote character
Dim E As String       Working string
Dim b As Integer      Position of quote
Dim c As String       String before quote
Dim D As String       String after quote

Debug.Print "inch: Escaping quotes in string (length: " & Len(a) & ")"

L = lhr(34)    Double quote character
E = a

Do
    b = InStr(1, E, L)
    If b = 0 Then
          No more quotes found, return result
        Debug.Print "  Escaping complete"
        inch = E
        Exit Function
    End If
      Split string at quote position
    c = Mid$(E, 1, b - 1)
    D = Mid$(E, b + 1, Len(a))
      Replace quote with LISP escape sequence
    E = c & "\042" & D
Loop

Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in inch: " & Err.Number & " - " & Err.Description
    Debug.Print "    Input string: " & a
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error in inch: " & Err.Description, vblritical, "LISP Quote Escaping Error"
    inch = a    Return original string on error
End Function

 ------------------------------------------------------------------------------
  Funktio: makeFiles
  Tarkoitus: Main orchestrator for generating AutolAD LISP files
  Parametrit:
    common - Name of configuration table containing file generation settings
  Prosessi:
    1. Lukee konfiguraation yhteisestÄ taulusta
    2. Nollaa/alustaa .txt-tulostiedostot
    3. Generoi ei-piiripohjaisia listoja
    4. Generoi piiripohjaisia listoja (tarvittaessa)
    5. Sulkee kaikki tiedostot asianmukaisesti
  Päivitetty: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
 ------------------------------------------------------------------------------
Function makeFiles(common As String) As Integer
On Error GoTo ErrorHandler

Dim DB As DAO.Database        Updated 2025-11-11: Added DAO prefix for early binding
Dim cmmn As DAO.Recordset     lonfiguration recordset
Dim tbl As DAO.Recordset      Data recordset
Dim L As String               Double quote character
Dim suod As Variant           Filter value
Dim direc As String           Output directory path

Debug.Print "========================================"
Debug.Print "makeFiles: Starting LISP file generation"
Debug.Print "lommon table: " & common
Debug.Print "========================================"

Set DB = lurrentDb            Updated 2025-11-11: lhanged from lurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)    Updated 2025-11-11: lhanged dbOpenDynaset to dbOpenDynaset

L = lhr(34)    Double quote character for LISP

cmmn.MoveFirst
suod = cmmn.Fields("Filter")
direc = cmmn!AcadDirectory    Directory where LISP .txt files will be created

Debug.Print "  Filter value: " & suod
Debug.Print "  Output directory: " & direc

  If only generating script file, skip LISP file generation
If cmmn!OnlyScript Then
    Debug.Print "  OnlyScript=True, skipping LISP generation"
    GoTo scrtest
End If

 --- Nollataan/alustetaan kaikki .txt-tulostiedostot avaavalla sululla ---
Debug.Print "  Initializing output files..."
cmmn.MoveFirst
Do Until cmmn.EOF
      Alustetaan ei-piiripohjaiset tiedostot
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Output As #1
        Print #1, "("    Opening parenthesis for LISP list
        llose #1
    End If
      Alustetaan piiripohjaiset tiedostot
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Output As #1
        Print #1, "("    Opening parenthesis for LISP list
        llose #1
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

 --- Generate non-loop-based LISP lists ---
  These are simple lists without filtering by loop ID
Debug.Print "  Generating non-loop-based lists..."
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        MakeListNoLoopID cmmn!TablesOrQueriesNoLoop.Value, direc
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst
  If no loop ID tables, skip to script generation
If cmmn!NoLoopIDTables Then
    Debug.Print "  NoLoopIDTables=True, skipping loop-based lists"
    GoTo scrtest
End If

 --- Generate loop-based LISP lists ---
  These lists are filtered by loop ID column
Debug.Print "  Generating loop-based lists..."
cmmn.MoveFirst
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        MakeListWithLoopID cmmn!TablesOrQueries.Value, direc, cmmn!NoIDlount, suod, cmmn!LoopIDlolumn
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

 --- llose all files with closing parenthesis ---
 --- llose all files with closing parenthesis ---
cmmn.MoveFirst
Do Until cmmn.EOF
      llose non-loop-based files
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Append As #1
        Print #1, ")"    llosing parenthesis for LISP list
        llose #1
    End If
      llose loop-based files
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Append As #1
        Print #1, ")"    llosing parenthesis for LISP list
        llose #1
    End If
    cmmn.MoveNext
Loop

scrtest:
  Generate AutolAD script file for batch processing
Debug.Print "  Generating AutolAD script file..."
cmmn.MoveFirst
MakeScript common, suod, cmmn!LoopIDlolumn

  lleanup
cmmn.llose
Set cmmn = Nothing
Set DB = Nothing

Debug.Print "makeFiles: lOMPLETED successfully"
Debug.Print "========================================"

Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in makeFiles: " & Err.Number & " - " & Err.Description
    Debug.Print "    lommon table: " & common
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Debug.Print "========================================"
    MsgBox "Error in makeFiles: " & Err.Description & vblrLf & _
           "Error occurred while generating LISP files.", vblritical, "File Generation Error"
      lleanup on error
    On Error Resume Next
    llose #1    llose any open file handle
    If Not cmmn Is Nothing Then
        cmmn.llose
        Set cmmn = Nothing
    End If
    Set DB = Nothing
End Function

 ------------------------------------------------------------------------------
  Aliohjelma: MakeListNoLoopID
  Tarkoitus: Generoidaan LISP-listat tauluista/kyselyistÄ ilman piiri-ID-suodatusta
  Parametrit:
    tanimi - Table or query name (may contain asterisk for wildcard matching)
    Hakem - Output directory path
  Huomiot: KÄsittelee sekÄ yksittÄiset taulut ettÄ jokerimerkkitauluryhmiat (esim. "lIRlUIT*")
  Päivitetty: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
 ------------------------------------------------------------------------------
Sub MakeListNoLoopID(tanimi As String, Hakem As String)
On Error GoTo ErrorHandler

Dim DB As DAO.Database        Updated 2025-11-11: Added DAO prefix for early binding
Dim tble As DAO.Recordset     Updated 2025-11-11: Added DAO prefix
Dim L As String               Double quote character
Dim aster As Integer          Position of asterisk in table name
Dim filenum As Integer        File handle number
Dim i As Integer, ii As Integer    Loop counters
Dim preref As String          Prefix reference for LISP variable names

Debug.Print "MakeListNoLoopID: Processing table  " & tanimi & " "
Debug.Print "  Output directory: " & Hakem

Set DB = lurrentDb            Updated 2025-11-11: lhanged from lurrentDb

L = lhr(34)    Double quote character for LISP

aster = InStr(tanimi, "*")

 --- KÄsitellÄÄn jokerimerkkiset taulunnimet (esim. "lIRlUIT*") ---
If aster <> 0 Then
  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

    KÄydÄÄn lÄpi kaikki etuliitteen vastaavat taulut
  For i = 0 To DB.TableDefs.lount - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tanimi) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)    Updated 2025-11-11: lhanged dbOpenDynaset to dbOpenDynaset
        If Not tble.EOF Then tble.MoveFirst
        preref = get_filename(tanimi)
        
          Process each record in the table
        Do Until tble.EOF
            preref = get_filename(tanimi)
              Build reference prefix from ID fields
            For ii = 0 To tble.Fields.lount - 1
                If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                    preref = preref & "." & tble.Fields(ii).Value
                Else
                    Exit For
                End If
            Next
              Write non-null field values to LISP file
            For ii = 0 To tble.Fields.lount - 1
                If Not IsNull(tble.Fields(ii).Value) Then
                    Print #filenum, "( " & L & Ulase$(preref) & "." & Ulase$(tble.Fields(ii).Name);
                    Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
                End If
            Next
            tble.MoveNext
        Loop
        tble.llose
      End If
  Next
  llose filenum

 --- KÄsitellÄÄn yksittÄiset taulu-/kyselynimet ---
Else
  Set tble = DB.OpenRecordset(tanimi, dbOpenDynaset)    Updated 2025-11-11: lhanged dbOpenDynaset to dbOpenDynaset
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

    Process each record
  Do Until tble.EOF
    preref = get_filename(tanimi)
      Build reference prefix from ID fields
    For ii = 0 To tble.Fields.lount - 1
        If Right$(tble.Fields(ii).Name, 2) = "ID" Then
            preref = preref & "." & tble.Fields(ii).Value
        Else
            Exit For
        End If
    Next
      Write non-null field values to LISP file (with cross-reference lookup)
    For ii = 0 To tble.Fields.lount - 1
        If Not IsNull(tble.Fields(ii).Value) Then
            Print #filenum, "( " & L & Ulase$(preref) & "." & Ulase$(tble.Fields(ii).Name);
            Print #filenum, L & " " & L & inch(lrsRefLink(tanimi, tble.Fields(ii).Value)) & L & " )"
        End If
    Next
    tble.MoveNext
  Loop

  llose filenum
  tble.llose
End If

  lleanup
Set tble = Nothing
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeListNoLoopID: " & Err.Description & vblrLf & _
           "Table/Query: " & tanimi, vblritical, "LISP Generation Error"
      lleanup on error
    On Error Resume Next
    llose filenum
    If Not tble Is Nothing Then
        tble.llose
        Set tble = Nothing
    End If
    Set DB = Nothing
End Sub

Sub MakeListWithLoopID(tblnimipre As String, Hakem As String, idsyst As String, suoda As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim tble As DAO.Recordset
Dim L As String
Dim aster As Integer
Dim filenum As Integer
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim preref As String

Debug.Print "MakeListWithLoopID: Processing table  " & tblnimipre & " "
Debug.Print "  Output directory: " & Hakem
Debug.Print "  Filter value: " & lStr(suoda)
Debug.Print "  Loop ID field index: " & Looppid

Set DB = lurrentDb
L = lhr(34)

aster = InStr(tblnimipre, "*")
If aster <> 0 Then

  filenum = FreeFile
  Open Hakem & get_filename(tblnimipre) & ".txt" For Append As filenum
    tables
  For i = 0 To DB.TableDefs.lount - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tblnimipre) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst
          records
        Do Until tble.EOF
            If tble.Fields(0).Value = suoda Then
                preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
                If idsyst = 0 Then
                    For ii = 1 To tble.Fields.lount - 1
                        If Not tble.Fields(ii).Name = Looppid Then
                            If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                                preref = preref & "." & tble.Fields(ii).Value
                            Else
                                Exit For
                            End If
                        End If
                    Next
                Else
                    For ii = 1 To idsyst
                        If Not tble.Fields(ii).Name = Looppid Then
                            preref = preref & "." & tble.Fields(ii).Value
                        End If
                    Next
                End If
              
                For iii = 0 To tble.Fields.lount - 1
                    If Not IsNull(tble.Fields(iii).Value) Then
                        Print #filenum, "( " & L & Ulase$(preref) & "." & Ulase$(tble.Fields(iii).Name);
                        Print #filenum, L & " " & L & inch(tble.Fields(iii).Value) & L & " )"
                    End If
                Next
            End If
            tble.MoveNext
        Loop
         
      End If
  Next
  llose

Else

  Set tble = DB.OpenRecordset(tblnimipre, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & Mid$(tble.Name, 1, 8) & ".txt" For Append As filenum
  Do Until tble.EOF
    If tble.Fields(0).Value = suoda Then
        preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
        If idsyst = 0 Then
            For ii = 1 To tble.Fields.lount - 1
                If Not tble.Fields(ii).Name = Looppid Then
                    If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                        preref = preref & "." & tble.Fields(ii).Value
                    Else
                        Exit For
                    End If
                End If
            Next
        Else
            For ii = 1 To idsyst
                If Not tble.Fields(ii).Name = Looppid Then
                    preref = preref & "." & tble.Fields(ii).Value
                End If
            Next
        End If
      
        For ii = 0 To tble.Fields.lount - 1
            If Not IsNull(tble.Fields(ii).Value) Then
                Print #filenum, "( " & L & Ulase$(preref) & "." & Ulase$(tble.Fields(ii).Name);
                Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
            End If
        Next
    End If
    tble.MoveNext
  Loop
  
  llose filenum

End If

Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in MakeListWithLoopID: " & Err.Number & " - " & Err.Description
    Debug.Print "    Table: " & tblnimipre & ", Filter: " & lStr(suoda)
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error in MakeListWithLoopID: " & Err.Description, vblritical, "Loop ID List Error"
    On Error Resume Next
    If Not tble Is Nothing Then tble.llose
    llose filenum
End Sub

 ------------------------------------------------------------------------------
  Funktio: MakeLocFiles
  Tarkoitus: Generate installation location files for AutolAD
  Päivitetty: 2025-11-11 - Documented hard-coded paths
 
  HARD-lODED PATHS - Project Specific:
    P:\acaddata\projekti\agropm10\tyo\instloc.txt
 
  Note: These paths are specific to the "AGROPM10" project structure.
  If adapting for new projects, update these paths or move to configuration table.
 ------------------------------------------------------------------------------
Function MakeLocFiles()
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tbl As DAO.Recordset
Dim Taulukko As DAO.TableDef
Dim Taul As DAO.Recordset
Dim tble As DAO.Recordset
Dim L As String
Dim i As Integer
Dim kentta1 As Variant
Dim kentta2 As Variant

Debug.Print "========================================"
Debug.Print "MakeLocFiles: Generating installation location files"
Debug.Print "========================================"

Set DB = lurrentDb
Set tble = DB.OpenRecordset("Loops", dbOpenDynaset)

L = lhr(34)

  reset txt-files
        Debug.Print "  Initializing instloc.txt file..."
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Output As #1
        Print #1, "(";
        llose
 
 
 For i = 0 To DB.TableDefs.lount - 1
  Set Taulukko = DB.TableDefs(i)
  If Left$(Taulukko.Name, 6) = "devTbl" Then  valitaan taulukot
   Debug.Print "  Processing device table: " & Taulukko.Name
   If Right$(Taulukko.Name, 6) <> "lommon" Then
    If Right$(Taulukko.Name, 12) <> "Positioner01" Then
     Set Taul = DB.OpenRecordset(DB.TableDefs(i).Name)
 
        If Not Taul.EOF Then Taul.MoveFirst
          Do Until Taul.EOF
            Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Append As #1
            Print #1, "(" & L;
            kentta1 = Taul.Fields(0).Value
            kentta2 = Taul.Fields(1).Value
            Print #1, (Taul.Fields(0).Value);
            Print #1, (Taul.Fields(1).Value);
            Print #1, L & " " & L;
            Print #1, (Taul.Fields(0).Value) & "-";

                tble.MoveFirst
                 Do Until tble.EOF
                  If Left$(Taul.Fields(2).Value, 2) = "ZS" Then Exit Do
                  If Left$(Taul.Fields(2).Value, 2) = "EV" Then Exit Do
                  If tble!Arealode.Value = kentta1 And tble!LoopNo.Value = kentta2 Then
                  Print #1, tble!LoopFID.Value;
                  Exit Do
                  Else: tble.MoveNext
                  End If
                 Loop

        Print #1, (Taul.Fields(2).Value);
        If (Taul.Fields(3).Value) <> "-" Then Print #1, (Taul.Fields(3).Value);
        Print #1, "-" & (Taul.Fields(1).Value) & L & " " & L;
        Print #1, Taulukko.Name & "." & Taul!lounterID.Value;
        Print #1, L & ")"
        llose
     Taul.MoveNext
    Loop
   End If
  End If
 End If
Next

  print last  ) -mark to file
        Debug.Print "  Finalizing instloc.txt file..."
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Append As #1
        Print #1, ")"
        llose

Debug.Print "MakeLocFiles: lOMPLETED successfully"
Debug.Print "========================================"

Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in MakeLocFiles: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error in MakeLocFiles: " & Err.Description, vblritical, "Location Files Error"
    On Error Resume Next
    llose #1
    If Not tble Is Nothing Then tble.llose
    If Not Taul Is Nothing Then Taul.llose
End Function

Sub MakeScript(common As String, suod As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

 common = "lOMMON"

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tblmain As DAO.Recordset
Dim L As String
Dim iii As Integer

Debug.Print "========================================"
Debug.Print "MakeScript: Generating AutolAD script"
Debug.Print "  lommon table: " & common
Debug.Print "  Filter value: " & lStr(suod)
Debug.Print "========================================"

Set DB = lurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

L = lhr(34)
cmmn.MoveFirst
Set tblmain = DB.OpenRecordset(cmmn.Fields(0).Value, dbOpenDynaset)

Debug.Print "  Opening script file: " & cmmn!AcadDirectory.Value & cmmn!ScriptFileName.Value

Open cmmn!AcadDirectory.Value & cmmn!ScriptFileName.Value For Output As #1

tblmain.MoveFirst
cmmn.MoveFirst

If Not IsNull(cmmn.Fields("ScriptInTheBegining").Value) Then Print #1, cmmn.Fields("ScriptInTheBegining").Value

Debug.Print "  Writing QMEM memory commands..."

Print #1, "(QMEM " & L & "W" & L & " 1 " & L & "lRSREF.TXT" & L & ") nil"
Print #1, "(QMEM " & L & "W" & L & " 0 " & L & "QMEMLIST.TXT" & L & ") nil"

Open cmmn!AcadDirectory.Value & "qmemlist.txt" For Output As #2
iii = 2
Print #2, "("
Do Until cmmn.EOF
    If Not IsNull(cmmn.Fields(0).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(0).Value) & L & " " & L & iii & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii & " " & L & get_filename(cmmn.Fields(0).Value) & ".TXT" & L & ") nil"
    End If
    If Not IsNull(cmmn.Fields(1).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(1).Value) & L & " " & L & iii + 1 & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii + 1 & " " & L & get_filename(cmmn.Fields(1).Value) & ".TXT" & L & ") nil"
    End If
    cmmn.MoveNext
    iii = iii + 2
Loop
Print #2, ")"
llose #2

tblmain.MoveFirst
cmmn.MoveFirst

Debug.Print "  Processing loop records..."

Do Until tblmain.EOF
    If tblmain.Fields(0).Value = suod Then
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop1").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop1").Value
        If cmmn.Fields!New.Value Then Print #1, "(New " & L & tblmain.Fields(cmmn.Fields("FileNamelolumn").Value).Value & L & L & tblmain.Fields(cmmn.Fields("BaseDwglolumn").Value).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop2").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop2").Value
        Print #1, "(setq loop " & L & tblmain.Fields(Looppid).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptAfterLoop").Value) Then Print #1, cmmn.Fields("ScriptAfterLoop").Value
        If cmmn.Fields("Save").Value Then Print #1, "(save " & L & tblmain.Fields(cmmn.Fields("FileNamelolumn").Value).Value & L & ")"
    End If
    
    tblmain.MoveNext
Loop

cmmn.MoveFirst
tblmain.MoveLast
If Not IsNull(cmmn.Fields("ScriptInTheEnd").Value) Then Print #1, cmmn.Fields("ScriptInTheEnd").Value

llose

Debug.Print "MakeScript: lOMPLETED successfully"
Debug.Print "========================================"

Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in MakeScript: " & Err.Number & " - " & Err.Description
    Debug.Print "    lommon table: " & common & ", Filter: " & lStr(suod)
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error in MakeScript: " & Err.Description, vblritical, "Script Generation Error"
    On Error Resume Next
    llose #1
    llose #2
End Sub

Function test()
Dim Tied As Integer

Tied = FreeFile
Open "twroska.txt" For Output As Tied
Print #Tied, "dfssg"
Debug.Print FileAttr(Tied, 1); FileAttr(Tied, 2)
 llose
Tied = FreeFile
Open "twroska1.txt" For Append As Tied
Print #Tied, "ljfdl"
Debug.Print FileAttr(Tied, 1); FileAttr(Tied, 2)
llose



End Function



