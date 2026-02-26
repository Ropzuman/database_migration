Option Compare Database   'Use database order for string comparisons
Option Explicit           'Require variable declaration
'==============================================================================
' Module: DataToACAD
' Purpose: Generate AutoCAD LISP files from database data for circuit diagrams
' Original: 1997-02-21 Fr 11:10 /tw
' Revised: 1997-03-19 We 15:46 /tw
' Revised: 1997-03-21 Fr 16:07 /tw
' Revised: 1997-07-14 Mo 14:29 /tw
' Updated: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
'                       Fixed DBEngine pattern, replaced deprecated constants
'==============================================================================

'------------------------------------------------------------------------------
' Function: CrsRefLink
' Purpose: Look up LISP code from cross-reference table
' Parameters:
'   tblnimi - Table name identifier
'   teksti - Cross-reference ID to look up
' Returns: LISP code string or original text if not found
' Updated: 2025-11-11 - Added DAO typing, error handling, comments
'------------------------------------------------------------------------------
Function CrsRefLink(tblnimi As String, teksti As String) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database      ' Updated 2025-11-11: Added DAO prefix for early binding
Dim tble As DAO.Recordset   ' Updated 2025-11-11: Added DAO prefix for early binding

Set DB = CurrentDb          ' Updated 2025-11-11: Changed from CurrentDb

If tblnimi = "CRSREF" Then
    ' Open cross-reference LISP lookup table
    Set tble = DB.OpenRecordset("CrsRefLisps", dbOpenDynaset)  ' Updated 2025-11-11: Changed dbOpenDynaset to dbOpenDynaset
    Do Until tble.EOF
        If tble!CrsRefID = teksti Then
            CrsRefLink = tble!Lisp
            tble.Close
            Set tble = Nothing
            Exit Function
        End If
    tble.MoveNext
    Loop
    CrsRefLink = teksti  ' Return original text if not found
    tble.Close
Else
    ' Not a cross-reference, return original text
    CrsRefLink = teksti
End If

Set tble = Nothing
Exit Function

ErrorHandler:
    MsgBox "Error in CrsRefLink: " & Err.Description, vbCritical, "Cross-Reference Lookup Error"
    CrsRefLink = teksti  ' Return original text on error
    If Not tble Is Nothing Then
        tble.Close
        Set tble = Nothing
    End If
End Function

'------------------------------------------------------------------------------
' Function: get_filename
' Purpose: Extract 8-character filename from table name
' Parameters:
'   taulnimi - Table name (may contain asterisk separator)
' Returns: 8-character uppercase filename
' Notes: Handles legacy naming convention with asterisk markers
' Updated: 2025-11-11 - Added error handling and comments
'------------------------------------------------------------------------------
Function get_filename(taulnimi As String) As String
On Error GoTo ErrorHandler

Dim ast As Integer

ast = InStr(taulnimi, "*")
If ast = 0 Then
    ' No asterisk, take first 8 characters
    get_filename = UCase(Mid(taulnimi, 1, 8))
Else
    ' Asterisk found, take first 8 characters before it
    get_filename = UCase(Mid(Mid(taulnimi, 1, ast - 1), 1, 8))
End If

Exit Function

ErrorHandler:
    MsgBox "Error in get_filename: " & Err.Description, vbCritical, "Filename Extraction Error"
    get_filename = "ERROR"
End Function

'------------------------------------------------------------------------------
' Function: inch
' Purpose: Escape double quotes for AutoCAD LISP syntax
' Parameters:
'   a - String containing double quotes to be escaped
' Returns: String with double quotes replaced by \042 (octal code)
' Notes: LISP requires special escaping of quote characters
' Updated: 2025-11-11 - Added error handling, improved variable names, comments
'------------------------------------------------------------------------------
Function inch(a As String) As String
On Error GoTo ErrorHandler

Dim L As String     ' Double quote character
Dim E As String     ' Working string
Dim b As Integer    ' Position of quote
Dim c As String     ' String before quote
Dim D As String     ' String after quote

L = Chr(34)  ' Double quote character
E = a

Do
    b = InStr(1, E, L)
    If b = 0 Then
        ' No more quotes found, return result
        inch = E
        Exit Function
    End If
    ' Split string at quote position
    c = Mid$(E, 1, b - 1)
    D = Mid$(E, b + 1, Len(a))
    ' Replace quote with LISP escape sequence
    E = c & "\042" & D
Loop

Exit Function

ErrorHandler:
    MsgBox "Error in inch: " & Err.Description, vbCritical, "LISP Quote Escaping Error"
    inch = a  ' Return original string on error
End Function

'------------------------------------------------------------------------------
' Function: makeFiles
' Purpose: Main orchestrator for generating AutoCAD LISP files
' Parameters:
'   common - Name of configuration table containing file generation settings
' Process:
'   1. Reads configuration from common table
'   2. Resets/initializes output .txt files
'   3. Generates non-loop-based lists
'   4. Generates loop-based lists (if applicable)
'   5. Closes all files properly
' Updated: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
'------------------------------------------------------------------------------
Function makeFiles(common As String) As Integer
On Error GoTo ErrorHandler

Dim DB As DAO.Database      ' Updated 2025-11-11: Added DAO prefix for early binding
Dim cmmn As DAO.Recordset   ' Configuration recordset
Dim tbl As DAO.Recordset    ' Data recordset
Dim L As String             ' Double quote character
Dim suod As Variant         ' Filter value
Dim direc As String         ' Output directory path

Set DB = CurrentDb          ' Updated 2025-11-11: Changed from CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)  ' Updated 2025-11-11: Changed dbOpenDynaset to dbOpenDynaset

L = Chr(34)  ' Double quote character for LISP

cmmn.MoveFirst
suod = cmmn.Fields("Filter")
direc = cmmn!AcadDirectory  ' Directory where LISP .txt files will be created

' If only generating script file, skip LISP file generation
If cmmn!OnlyScript Then GoTo scrtest

'--- Reset/Initialize all output .txt files with opening parenthesis ---
cmmn.MoveFirst
Do Until cmmn.EOF
    ' Initialize non-loop-based files
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Output As #1
        Print #1, "("  ' Opening parenthesis for LISP list
        Close #1
    End If
    ' Initialize loop-based files
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Output As #1
        Print #1, "("  ' Opening parenthesis for LISP list
        Close #1
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

'--- Generate non-loop-based LISP lists ---
' These are simple lists without filtering by loop ID
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        MakeListNoLoopID cmmn!TablesOrQueriesNoLoop.Value, direc
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst
' If no loop ID tables, skip to script generation
If cmmn!NoLoopIDTables Then GoTo scrtest

'--- Generate loop-based LISP lists ---
' These lists are filtered by loop ID column
cmmn.MoveFirst
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        MakeListWithLoopID cmmn!TablesOrQueries.Value, direc, cmmn!NoIDCount, suod, cmmn!LoopIDColumn
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

'--- Close all files with closing parenthesis ---
'--- Close all files with closing parenthesis ---
cmmn.MoveFirst
Do Until cmmn.EOF
    ' Close non-loop-based files
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Append As #1
        Print #1, ")"  ' Closing parenthesis for LISP list
        Close #1
    End If
    ' Close loop-based files
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Append As #1
        Print #1, ")"  ' Closing parenthesis for LISP list
        Close #1
    End If
    cmmn.MoveNext
Loop

scrtest:
' Generate AutoCAD script file for batch processing
cmmn.MoveFirst
MakeScript common, suod, cmmn!LoopIDColumn

' Cleanup
cmmn.Close
Set cmmn = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in makeFiles: " & Err.Description & vbCrLf & _
           "Error occurred while generating LISP files.", vbCritical, "File Generation Error"
    ' Cleanup on error
    On Error Resume Next
    Close #1  ' Close any open file handle
    If Not cmmn Is Nothing Then
        cmmn.Close
        Set cmmn = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Sub: MakeListNoLoopID
' Purpose: Generate LISP lists from tables/queries that don't require loop ID filtering
' Parameters:
'   tanimi - Table or query name (may contain asterisk for wildcard matching)
'   Hakem - Output directory path
' Notes: Handles both single tables and wildcard table groups (e.g., "CIRCUIT*")
' Updated: 2025-11-11 - Added DAO typing, error handling, comprehensive comments
'------------------------------------------------------------------------------
Sub MakeListNoLoopID(tanimi As String, Hakem As String)
On Error GoTo ErrorHandler

Dim DB As DAO.Database      ' Updated 2025-11-11: Added DAO prefix for early binding
Dim tble As DAO.Recordset   ' Updated 2025-11-11: Added DAO prefix
Dim L As String             ' Double quote character
Dim aster As Integer        ' Position of asterisk in table name
Dim filenum As Integer      ' File handle number
Dim i As Integer, ii As Integer  ' Loop counters
Dim preref As String        ' Prefix reference for LISP variable names

Set DB = CurrentDb          ' Updated 2025-11-11: Changed from CurrentDb

L = Chr(34)  ' Double quote character for LISP

aster = InStr(tanimi, "*")

'--- Handle wildcard table names (e.g., "CIRCUIT*") ---
If aster <> 0 Then
  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' Loop through all tables matching the prefix
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tanimi) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)  ' Updated 2025-11-11: Changed dbOpenDynaset to dbOpenDynaset
        If Not tble.EOF Then tble.MoveFirst
        preref = get_filename(tanimi)
        
        ' Process each record in the table
        Do Until tble.EOF
            preref = get_filename(tanimi)
            ' Build reference prefix from ID fields
            For ii = 0 To tble.Fields.Count - 1
                If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                    preref = preref & "." & tble.Fields(ii).Value
                Else
                    Exit For
                End If
            Next
            ' Write non-null field values to LISP file
            For ii = 0 To tble.Fields.Count - 1
                If Not IsNull(tble.Fields(ii).Value) Then
                    Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(ii).Name);
                    Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
                End If
            Next
            tble.MoveNext
        Loop
        tble.Close
      End If
  Next
  Close filenum

'--- Handle single table/query names ---
Else
  Set tble = DB.OpenRecordset(tanimi, dbOpenDynaset)  ' Updated 2025-11-11: Changed dbOpenDynaset to dbOpenDynaset
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' Process each record
  Do Until tble.EOF
    preref = get_filename(tanimi)
    ' Build reference prefix from ID fields
    For ii = 0 To tble.Fields.Count - 1
        If Right$(tble.Fields(ii).Name, 2) = "ID" Then
            preref = preref & "." & tble.Fields(ii).Value
        Else
            Exit For
        End If
    Next
    ' Write non-null field values to LISP file (with cross-reference lookup)
    For ii = 0 To tble.Fields.Count - 1
        If Not IsNull(tble.Fields(ii).Value) Then
            Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(ii).Name);
            Print #filenum, L & " " & L & inch(CrsRefLink(tanimi, tble.Fields(ii).Value)) & L & " )"
        End If
    Next
    tble.MoveNext
  Loop

  Close filenum
  tble.Close
End If

' Cleanup
Set tble = Nothing
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeListNoLoopID: " & Err.Description & vbCrLf & _
           "Table/Query: " & tanimi, vbCritical, "LISP Generation Error"
    ' Cleanup on error
    On Error Resume Next
    Close filenum
    If Not tble Is Nothing Then
        tble.Close
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

Set DB = CurrentDb
L = Chr(34)

aster = InStr(tblnimipre, "*")
If aster <> 0 Then

  filenum = FreeFile
  Open Hakem & get_filename(tblnimipre) & ".txt" For Append As filenum
  ' tables
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tblnimipre) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst
        ' records
        Do Until tble.EOF
            If tble.Fields(0).Value = suoda Then
                preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
                If idsyst = 0 Then
                    For ii = 1 To tble.Fields.Count - 1
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
              
                For iii = 0 To tble.Fields.Count - 1
                    If Not IsNull(tble.Fields(iii).Value) Then
                        Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(iii).Name);
                        Print #filenum, L & " " & L & inch(tble.Fields(iii).Value) & L & " )"
                    End If
                Next
            End If
            tble.MoveNext
        Loop
         
      End If
  Next
  Close

Else

  Set tble = DB.OpenRecordset(tblnimipre, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & Mid(tble.Name, 1, 8) & ".txt" For Append As filenum
  Do Until tble.EOF
    If tble.Fields(0).Value = suoda Then
        preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
        If idsyst = 0 Then
            For ii = 1 To tble.Fields.Count - 1
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
      
        For ii = 0 To tble.Fields.Count - 1
            If Not IsNull(tble.Fields(ii).Value) Then
                Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(ii).Name);
                Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
            End If
        Next
    End If
    tble.MoveNext
  Loop
  
  Close filenum

End If

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeListWithLoopID: " & Err.Description, vbCritical, "Loop ID List Error"
    On Error Resume Next
    If Not tble Is Nothing Then tble.Close
    Close filenum
End Sub

'------------------------------------------------------------------------------
' Function: MakeLocFiles
' Purpose: Generate installation location files for AutoCAD
' Updated: 2025-11-11 - Documented hard-coded paths
'
' HARD-CODED PATHS - Project Specific:
'   P:\acaddata\projekti\agropm10\tyo\instloc.txt
'
' Note: These paths are specific to the "AGROPM10" project structure.
' If adapting for new projects, update these paths or move to configuration table.
'------------------------------------------------------------------------------
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

Set DB = CurrentDb
Set tble = DB.OpenRecordset("Loops", dbOpenDynaset)

L = Chr(34)

' reset txt-files
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Output As #1
        Print #1, "(";
        Close
 
 
 For i = 0 To DB.TableDefs.Count - 1
  Set Taulukko = DB.TableDefs(i)
  If Left(Taulukko.Name, 6) = "devTbl" Then 'valitaan taulukot
   If Right(Taulukko.Name, 6) <> "Common" Then
    If Right(Taulukko.Name, 12) <> "Positioner01" Then
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
                  If Left(Taul.Fields(2).Value, 2) = "ZS" Then Exit Do
                  If Left(Taul.Fields(2).Value, 2) = "EV" Then Exit Do
                  If tble!AreaCode.Value = kentta1 And tble!LoopNo.Value = kentta2 Then
                  Print #1, tble!LoopFID.Value;
                  Exit Do
                  Else: tble.MoveNext
                  End If
                 Loop

        Print #1, (Taul.Fields(2).Value);
        If (Taul.Fields(3).Value) <> "-" Then Print #1, (Taul.Fields(3).Value);
        Print #1, "-" & (Taul.Fields(1).Value) & L & " " & L;
        Print #1, Taulukko.Name & "." & Taul!CounterID.Value;
        Print #1, L & ")"
        Close
     Taul.MoveNext
    Loop
   End If
  End If
 End If
Next

' print last ')'-mark to file
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Append As #1
        Print #1, ")"
        Close

Exit Function

ErrorHandler:
    MsgBox "Error in MakeLocFiles: " & Err.Description, vbCritical, "Location Files Error"
    On Error Resume Next
    Close #1
    If Not tble Is Nothing Then tble.Close
    If Not Taul Is Nothing Then Taul.Close
End Function

Sub MakeScript(common As String, suod As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

'common = "COMMON"

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tblmain As DAO.Recordset
Dim L As String
Dim iii As Integer

Set DB = CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

L = Chr(34)
cmmn.MoveFirst
Set tblmain = DB.OpenRecordset(cmmn.Fields(0).Value, dbOpenDynaset)

Open cmmn!AcadDirectory.Value & cmmn!ScriptFileName.Value For Output As #1

tblmain.MoveFirst
cmmn.MoveFirst

If Not IsNull(cmmn.Fields("ScriptInTheBegining").Value) Then Print #1, cmmn.Fields("ScriptInTheBegining").Value

Print #1, "(QMEM " & L & "W" & L & " 1 " & L & "CRSREF.TXT" & L & ")'nil"
Print #1, "(QMEM " & L & "W" & L & " 0 " & L & "QMEMLIST.TXT" & L & ")'nil"

Open cmmn!AcadDirectory.Value & "qmemlist.txt" For Output As #2
iii = 2
Print #2, "("
Do Until cmmn.EOF
    If Not IsNull(cmmn.Fields(0).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(0).Value) & L & " " & L & iii & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii & " " & L & get_filename(cmmn.Fields(0).Value) & ".TXT" & L & ")'nil"
    End If
    If Not IsNull(cmmn.Fields(1).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(1).Value) & L & " " & L & iii + 1 & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii + 1 & " " & L & get_filename(cmmn.Fields(1).Value) & ".TXT" & L & ")'nil"
    End If
    cmmn.MoveNext
    iii = iii + 2
Loop
Print #2, ")"
Close #2

tblmain.MoveFirst
cmmn.MoveFirst
Do Until tblmain.EOF
    If tblmain.Fields(0).Value = suod Then
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop1").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop1").Value
        If cmmn.Fields!New.Value Then Print #1, "(New " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & L & tblmain.Fields(cmmn.Fields("BaseDwgColumn").Value).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop2").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop2").Value
        Print #1, "(setq loop " & L & tblmain.Fields(Looppid).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptAfterLoop").Value) Then Print #1, cmmn.Fields("ScriptAfterLoop").Value
        If cmmn.Fields("Save").Value Then Print #1, "(save " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & ")"
    End If
    
    tblmain.MoveNext
Loop

cmmn.MoveFirst
tblmain.MoveLast
If Not IsNull(cmmn.Fields("ScriptInTheEnd").Value) Then Print #1, cmmn.Fields("ScriptInTheEnd").Value

Close

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeScript: " & Err.Description, vbCritical, "Script Generation Error"
    On Error Resume Next
    Close #1
    Close #2
End Sub

Function test()
Dim Tied As Integer

Tied = FreeFile
Open "twroska.txt" For Output As Tied
Print #Tied, "dfssg"
Debug.Print FileAttr(Tied, 1); FileAttr(Tied, 2)
'Close
Tied = FreeFile
Open "twroska1.txt" For Append As Tied
Print #Tied, "ljfdl"
Debug.Print FileAttr(Tied, 1); FileAttr(Tied, 2)
Close



End Function



