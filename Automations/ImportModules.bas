Attribute VB_Name = "ImportModules"
Option Explicit

' ImportModules.bas
' Import VBA modules from files into current database
' 
' USAGE:
' 1. Open your database (MAINEQ.accdb)
' 2. Press Alt+F11 (VBA Editor)
' 3. File -> Import File -> Select this ImportModules.bas
' 4. In Immediate Window (Ctrl+G), run: ImportAllModules
' 5. Follow prompts for source folder

Sub ImportAllModules()
    Dim sourcePath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbProj As Object
    Dim comp As Object
    Dim imported As Long
    Dim skipped As Long
    
    ' Get source folder
    sourcePath = InputBox("Enter full path to folder with .bas and .cls files:", _
                         "Import Modules", _
                         "C:\database_migration\Access\MAINEQ")
    
    If sourcePath = "" Then
        MsgBox "Cancelled by user", vbInformation
        Exit Sub
    End If
    
    ' Verify folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sourcePath) Then
        MsgBox "Folder not found: " & sourcePath, vbCritical
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(sourcePath)
    Set vbProj = Application.VBE.ActiveVBProject
    
    imported = 0
    skipped = 0
    
    Debug.Print "=== IMPORTING MODULES FROM: " & sourcePath & " ==="
    Debug.Print ""
    
    ' Process all files
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Or LCase(Right(file.Name, 4)) = ".cls" Then
            
            Dim moduleName As String
            Dim fileContent As String
            Dim cleanCode As String
            Dim isBasFile As Boolean
            
            moduleName = Left(file.Name, Len(file.Name) - 4)
            isBasFile = LCase(Right(file.Name, 4)) = ".bas"
            
            ' Read and clean the file content
            fileContent = ReadTextFile(file.Path)
            
            If isBasFile Then
                cleanCode = StripBasHeaders(fileContent)
            Else
                cleanCode = StripClassHeaders(fileContent)
            End If
            
            ' Try to find existing component
            Set comp = Nothing
            On Error Resume Next
            Set comp = vbProj.VBComponents(moduleName)
            On Error GoTo 0
            
            If comp Is Nothing Then
                ' Component doesn't exist - create it
                On Error Resume Next
                If isBasFile Then
                    Set comp = vbProj.VBComponents.Add(1) ' vbext_ct_StdModule
                    comp.Name = moduleName
                Else
                    ' For .cls files, check if it's a Form/Report or standalone class
                    If Left(moduleName, 5) = "Form_" Or Left(moduleName, 7) = "Report_" Then
                        ' This is a form/report that doesn't exist yet
                        Debug.Print "⚠ SKIPPED: " & moduleName & " - Form/Report not found in database"
                        Debug.Print "   Create the form/report first, then re-import"
                        skipped = skipped + 1
                        On Error GoTo 0
                        GoTo NextFile
                    Else
                        ' Standalone class module
                        Set comp = vbProj.VBComponents.Add(2) ' vbext_ct_ClassModule
                        comp.Name = moduleName
                    End If
                End If
                
                If Err.Number <> 0 Then
                    Debug.Print "✗ FAILED to create: " & moduleName & " - " & Err.Description
                    skipped = skipped + 1
                    On Error GoTo 0
                    GoTo NextFile
                End If
                On Error GoTo 0
                Debug.Print "  Created new: " & moduleName
            End If
            
            ' Now update the code
            On Error Resume Next
            ' Clear old code
            If comp.CodeModule.CountOfLines > 0 Then
                comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
            End If
            
            ' Add new code
            If cleanCode <> "" Then
                comp.CodeModule.AddFromString cleanCode
            End If
            
            If Err.Number = 0 Then
                Debug.Print "✓ Updated: " & moduleName & " (" & comp.CodeModule.CountOfLines & " lines)"
                imported = imported + 1
            Else
                Debug.Print "✗ FAILED: " & moduleName & " - " & Err.Description
                skipped = skipped + 1
            End If
            On Error GoTo 0
            
NextFile:
        End If
    Next file
    
    Debug.Print ""
    Debug.Print "=== IMPORT COMPLETE ==="
    Debug.Print "Imported: " & imported
    Debug.Print "Failed: " & skipped
    Debug.Print ""
    
    MsgBox "Import complete!" & vbCrLf & vbCrLf & _
           "Imported: " & imported & vbCrLf & _
           "Failed: " & skipped & vbCrLf & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details.", _
           vbInformation, "Import Complete"
    
    Set fso = Nothing
End Sub

' Helper function to read text file
Function ReadTextFile(filePath As String) As String
    Dim fso As Object
    Dim stream As Object
    Dim content As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = CreateObject("ADODB.Stream")
    
    ' Use ADODB.Stream for proper UTF-8 handling
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    content = stream.ReadText
    stream.Close
    
    ReadTextFile = content
    Set stream = Nothing
    Set fso = Nothing
End Function

' Helper function to strip .bas file headers
Function StripBasHeaders(fileContent As String) As String
    Dim lines() As String
    Dim i As Long
    Dim startIndex As Long
    Dim line As String
    
    lines = Split(fileContent, vbCrLf)
    If UBound(lines) = 0 Then
        lines = Split(fileContent, vbLf) ' Try Unix line endings
    End If
    
    startIndex = 0
    
    ' Skip ONLY Attribute VB_Name lines at the start
    ' Don't strip anything else - files are already clean
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        
        ' Only skip Attribute VB_Name line
        If line Like "Attribute VB_Name *" Then
            startIndex = i + 1
        Else
            ' Hit first non-attribute line, stop
            Exit For
        End If
    Next i
    
    ' Join remaining lines
    Dim result As String
    result = ""
    For i = startIndex To UBound(lines)
        If result <> "" Then result = result & vbCrLf
        result = result & lines(i)
    Next i
    
    StripBasHeaders = result
End Function

' Helper function to strip .cls file headers
Function StripClassHeaders(fileContent As String) As String
    ' Our .cls files are already clean - no VERSION/BEGIN/END headers
    ' Just return the content as-is
    ' If exported from Access with headers, they would need stripping,
    ' but our files are hand-edited and clean
    StripClassHeaders = fileContent
End Function

Sub RemoveAllModules()
    ' Helper: Remove all standard modules and class modules
    ' WARNING: Use with caution!
    
    Dim vbProj As Object
    Dim comp As Object
    Dim i As Long
    Dim removed As Long
    
    If MsgBox("This will remove ALL standard modules and class modules!" & vbCrLf & _
              "Forms and Reports will NOT be removed." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbExclamation, "Remove All Modules") = vbNo Then
        Exit Sub
    End If
    
    Set vbProj = Application.VBE.ActiveVBProject
    removed = 0
    
    ' Loop backwards to avoid index issues
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        
        ' Only remove standard modules (Type 1) and class modules (Type 2)
        ' Don't touch forms (Type 100) or reports
        If comp.Type = 1 Or comp.Type = 2 Then
            Debug.Print "Removing: " & comp.Name
            vbProj.VBComponents.Remove comp
            removed = removed + 1
        End If
    Next i
    
    MsgBox "Removed " & removed & " modules.", vbInformation
End Sub

Sub ListCurrentModules()
    ' Helper: List all current modules with their types
    
    Dim vbProj As Object
    Dim comp As Object
    Dim typeDesc As String
    Dim standardCount As Long
    Dim classCount As Long
    Dim formCount As Long
    Dim reportCount As Long
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    Debug.Print "=== CURRENT VBA COMPONENTS ==="
    Debug.Print ""
    
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case 1
                typeDesc = "Standard Module (.bas)"
                standardCount = standardCount + 1
            Case 2
                typeDesc = "Class Module (.cls)"
                classCount = classCount + 1
            Case 100
                typeDesc = "Form (with code)"
                formCount = formCount + 1
            Case 101
                typeDesc = "Report (with code)"
                reportCount = reportCount + 1
            Case Else
                typeDesc = "Other (Type " & comp.Type & ")"
        End Select
        
        Debug.Print "  " & comp.Name & " - " & typeDesc
    Next comp
    
    Debug.Print ""
    Debug.Print "SUMMARY:"
    Debug.Print "  Standard Modules: " & standardCount
    Debug.Print "  Class Modules: " & classCount
    Debug.Print "  Forms: " & formCount
    Debug.Print "  Reports: " & reportCount
    Debug.Print "  Total: " & vbProj.VBComponents.Count
    Debug.Print ""
    
    MsgBox "Components listed in Immediate Window (Ctrl+G)", vbInformation
End Sub
