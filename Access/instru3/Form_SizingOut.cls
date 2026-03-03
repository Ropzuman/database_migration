Option Compare Database
Option Explicit

'================================================================================
' Form: SizingOut
' Purpose: Export query results to CSV file with Finnish number formatting
' Updated: 2025-11-11 - Added VBA7/64-bit support
'
' Description:
'   Exports database query or table to semicolon-delimited CSV file:
'   - User selects query/table, output directory, and filename
'   - Converts decimal commas to periods (Finnish → international format)
'   - Wraps all values in quotes for proper CSV formatting
'   - Opens output folder in Windows Explorer after export
'   - Handles null values appropriately
'
' Dependencies:
'   - FileSystemObject (Microsoft Scripting Runtime)
'   - Query/table specified in Taulukko control
'
' Controls:
'   - Taulukko: TextBox for query/table name (default: qryValveSizingOut)
'   - Hakem: TextBox for output directory path
'   - FileName: TextBox for output filename (without .csv extension)
'   - Command8: Button to execute export
'================================================================================

'--------------------------------------------------------------------------------
' Event: Command8_Click
' Purpose: Export selected query/table to CSV file
'
' Process:
'   1. Validate output directory exists
'   2. Open source query/table as recordset
'   3. Create output CSV file
'   4. Iterate through all records:
'      - Write first field (quoted)
'      - Write remaining fields with semicolon separator
'      - Convert commas to periods in numeric values
'      - Handle null values as empty strings
'   5. Close file and open containing folder
'
' CSV Format:
'   - Semicolon (;) delimited
'   - All values quoted ("value")
'   - Decimal separator: period (.) not comma (,)
'   - Line ending: CRLF (Windows standard)
'
' Notes:
'   - Ensures directory path ends with backslash
'   - Handles Finnish decimal format (comma) conversion
'   - Opens output folder for user convenience
'   - Comprehensive error handling with cleanup
'--------------------------------------------------------------------------------
Private Sub Command8_Click()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Current database
    Dim Taulu As DAO.Recordset  ' Source data recordset
    Dim FSO As New FileSystemObject  ' File system object for file creation
    Dim Tiedosto As Object  ' Output CSV file handle
    Dim Path As String  ' Output directory path
    Dim i As Integer  ' Field counter
    Dim Osoitin As Long  ' Position of comma in value string
    Dim Arvo As String  ' Current field value
    
    Set DB = CurrentDb
    Path = Me.Hakem.Value
    
    ' Ensure path ends with backslash
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    ' Validate output directory exists
    If Not FSO.FolderExists(Path) Then
        MsgBox "Hakemistoa: " & Path & " ei ollut", vbCritical, "Export Error"  ' "Directory does not exist"
        Exit Sub
    End If
    
    ' Open source query/table
    Set Taulu = DB.OpenRecordset("SELECT * FROM " & Me.Taulukko.Value)
    
    ' Create output CSV file
    Set Tiedosto = FSO.CreateTextFile(Path & Me.FileName.Value & ".csv")
    
    ' Export all records
    Do While Not Taulu.EOF
        ' Write first field (quoted)
        Tiedosto.Write """" & Taulu.Fields(0).Value & """"
        
        ' Write remaining fields with semicolon separator
        For i = 1 To Taulu.Fields.Count - 1
            ' Handle null values
            If IsNull(Taulu.Fields(i).Value) Then
                Arvo = ""
            Else
                Arvo = Taulu.Fields(i).Value
                
                ' Convert decimal comma to period (Finnish → international format)
                Osoitin = InStr(Arvo, ",")
                If Osoitin <> 0 Then
                    Arvo = Left$(Arvo, Osoitin - 1) & "." & Mid$(Arvo, Osoitin + 1)
                End If
            End If
            
            ' Write field value (quoted, semicolon-separated)
            Tiedosto.Write ";""" & Arvo & """"
        Next i
        
        ' End record with line break
        Tiedosto.Write vbCrLf
        Taulu.MoveNext
    Loop
    
    ' Cleanup resources
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    
    ' Notify user and open output folder
    MsgBox "Valmis", vbInformation, "Export Complete"  ' "Ready"
    Call Shell("EXPLORER " & Path, vbNormalFocus)
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vbCritical, "Export Error"
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' Event: Form_Load
' Purpose: Initialize form with default values
'
' Defaults:
'   - Taulukko (Query/Table): qryValveSizingOut
'   - Hakem (Directory): User must specify (commented example: "d:\tilap\")
'
' Notes:
'   - Default query focuses on valve sizing export
'   - Directory path intentionally left blank (was "d:\tilap\")
'--------------------------------------------------------------------------------
Private Sub Form_Load()
  Me.Taulukko.Value = "qryValveSizingOut"  ' Oletusky sely venttiilien mitoitusdata vientiin
  'Hakem.VALUE = "d:\tilap\"  ' Example directory (commented out - user specified)
End Sub
