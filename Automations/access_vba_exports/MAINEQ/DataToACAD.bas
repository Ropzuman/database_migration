Attribute VB_Name = "DataToACAD"
Option Compare Database   'Use database order for string comparisons
                           
' Circuit and Jb -dwgs
' 1997-02-21 Fr 11:10 /tw
' 1997-03-19 We 15:46 /tw
' 1997-03-21 Fr 16:07 /tw
' 1997-07-14 Mo 14:29 /tw

Function CrsRefLink(tblnimi, teksti)

Dim DB As Database
Dim tble As Recordset
Set DB = DBEngine.Workspaces(0).Databases(0)

If tblnimi = "CRSREF" Then
    Set tble = DB.OpenRecordset("CrsRefLisps", DB_OPEN_DYNASET)
    Do Until tble.EOF
        If tble!CrsRefID = teksti Then
            CrsRefLink = tble!Lisp
            Exit Function
        End If
    tble.MoveNext
    Loop
    CrsRefLink = teksti

Else
    CrsRefLink = teksti
End If

End Function

Function get_filename(taulnimi)

ast = InStr(taulnimi, "*")
If ast = 0 Then
    get_filename = UCase(Mid(taulnimi, 1, 8))
Else
  get_filename = UCase(Mid(Mid(taulnimi, 1, ast - 1), 1, 8))

End If

End Function

Function inch(a)
L = Chr(34)
E = a
Do
    b = InStr(1, E, L)
    If b = 0 Then
        inch = E
        Exit Function
    End If
    c = Mid$(E, 1, b - 1)
    D = Mid$(E, b + 1, Len(a))
    E = c & "\042" & D
Loop

End Function

Function makeFiles(common)

'common = "COMMON"

Dim DB As Database
Dim cmmn As Recordset
Dim tbl As Recordset

Set DB = DBEngine.Workspaces(0).Databases(0)
Set cmmn = DB.OpenRecordset(common, DB_OPEN_DYNASET)

L = Chr(34)

cmmn.MoveFirst
suod = cmmn.Fields("Filter")
direc = cmmn!AcadDirectory

If cmmn!OnlyScript Then GoTo scrtest

' reset txt-files
cmmn.MoveFirst
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Output As #1
        Print #1, "("
        Close
    End If
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Output As #1
        Print #1, "("
        Close
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

' print no-loop -lists to files
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        MakeListNoLoopID cmmn!TablesOrQueriesNoLoop.Value, direc
    End If
    cmmn.MoveNext
Loop


cmmn.MoveFirst
If cmmn!NoLoopIDTables Then GoTo scrtest
'Set tbl = db.OpenRecordset(cmmn.TablesOrQueries.value, db_open_dynaset)

' print loop based lists to files

'tbl.MoveFirst
'Do Until tbl.eof
  ' print loop to files
'  If tbl.fields(0).value = suod Then
    ' print datas to files
    cmmn.MoveFirst
    Do Until cmmn.EOF
        If Not IsNull(cmmn!TablesOrQueries.Value) Then
            MakeListWithLoopID cmmn!TablesOrQueries.Value, direc, cmmn!NoIDCount, suod, cmmn!LoopIDColumn
        End If
        cmmn.MoveNext
    Loop

cmmn.MoveFirst

' print last ')'-mark to files
cmmn.MoveFirst
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Append As #1
        Print #1, ")"
        Close
    End If
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Append As #1
        Print #1, ")"
        Close
    End If
    cmmn.MoveNext
Loop

scrtest:
cmmn.MoveFirst
MakeScript common, suod, cmmn!LoopIDColumn

End Function

Sub MakeListNoLoopID(tanimi, Hakem)

Dim DB As Database
Dim tble As Recordset
Set DB = DBEngine.Workspaces(0).Databases(0)

L = Chr(34)

aster = InStr(tanimi, "*")
If aster <> 0 Then
  
  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tanimi) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, DB_OPEN_DYNASET)
        If Not tble.EOF Then tble.MoveFirst
        preref = get_filename(tanimi)
        Do Until tble.EOF
            preref = get_filename(tanimi)
            For ii = 0 To tble.Fields.Count - 1
                If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                    preref = preref & "." & tble.Fields(ii).Value
                Else
                    Exit For
                End If
            Next
            For ii = 0 To tble.Fields.Count - 1
                If Not IsNull(tble.Fields(ii).Value) Then
                    Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(ii).Name);
                    Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
                End If
            Next
            tble.MoveNext
        Loop
      End If
  Next
  Close
  

Else

Set tble = DB.OpenRecordset(tanimi, DB_OPEN_DYNASET)
If Not tble.EOF Then tble.MoveFirst

filenum = FreeFile
Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

'Print #filenum, "( "
Do Until tble.EOF
    preref = get_filename(tanimi)
    For ii = 0 To tble.Fields.Count - 1
        If Right$(tble.Fields(ii).Name, 2) = "ID" Then
            preref = preref & "." & tble.Fields(ii).Value
        Else
            Exit For
        End If
    Next
    For ii = 0 To tble.Fields.Count - 1
        If Not IsNull(tble.Fields(ii).Value) Then
            Print #filenum, "( " & L & UCase(preref) & "." & UCase(tble.Fields(ii).Name);
            Print #filenum, L & " " & L & inch(CrsRefLink(tanimi, tble.Fields(ii).Value)) & L & " )"
        End If
    Next
    tble.MoveNext
Loop
'Print #filenum, " )"

Close filenum

End If

End Sub

Sub MakeListWithLoopID(tblnimipre, Hakem, idsyst, suoda, Looppid)

Dim DB As Database
Dim tble As Recordset
Set DB = DBEngine.Workspaces(0).Databases(0)
L = Chr(34)

aster = InStr(tblnimipre, "*")
If aster <> 0 Then

  filenum = FreeFile
  Open Hakem & get_filename(tblnimipre) & ".txt" For Append As filenum
  ' tables
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tblnimipre) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, DB_OPEN_DYNASET)
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

  Set tble = DB.OpenRecordset(tblnimipre, DB_OPEN_DYNASET)
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

End Sub

Function MakeLocFiles()

Dim DB As Database
Dim cmmn As Recordset
Dim tbl As Recordset
Dim Taulukko As TableDef
Dim Taul As Recordset
Dim tble As Recordset

Set DB = DBEngine.Workspaces(0).Databases(0)
Set tble = DB.OpenRecordset("Loops", DB_OPEN_DYNASET)

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

End Function

Sub MakeScript(common, suod, Looppid)

'common = "COMMON"

Dim DB As Database
Dim cmmn As Recordset
Dim tblmain As Recordset

Set DB = DBEngine.Workspaces(0).Databases(0)
Set cmmn = DB.OpenRecordset(common, DB_OPEN_DYNASET)

L = Chr(34)
cmmn.MoveFirst
Set tblmain = DB.OpenRecordset(cmmn.Fields(0).Value, DB_OPEN_DYNASET)

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


End Sub

Function test()

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

