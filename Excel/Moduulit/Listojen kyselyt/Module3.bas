'''
' Module3.vba - Apumakro LINKING-sheetin näkyvyyden vaihtamiseen
' Käytetään LINKING-sheetin näyttämiseen/piilottamiseen työkirjassa debuggausta tai käyttäjän mieltymyksiä varten.
'''
Sub Linking()
    ' Vaihtaa LINKING-sheetin näkyvyyden työkirjassa.
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim linkedSheetFound As Boolean
    linkedSheetFound = False
    
    Debug.Print Format(Now, "hh:mm:ss") & " [Linking] Vaihdetaan LINKING-sheetin näkyvyyttä"
    
    For i = 1 To Sheets.Count
        If LCase(Sheets(i).Name) = "linking" Then
            linkedSheetFound = True
            If Sheets(i).Visible = True Then
                Sheets(i).Visible = False
                Debug.Print "  LINKING-sheet piilotettu"
            Else
                Sheets(i).Visible = True
                Debug.Print "  LINKING-sheet näkyvissä"
            End If
            Exit For
        End If
    Next i
    
    If Not linkedSheetFound Then
        Debug.Print "  VAROITUS: LINKING-sheetiä ei löytynyt"
        MsgBox "LINKING-sheetiä ei löytynyt työkirjasta.", vbInformation, "Sheet puuttuu"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print Format(Now, "hh:mm:ss") & " [Linking ERROR] " & Err.Number & ": " & Err.Description
    MsgBox "Virhe LINKING-sheetin näkyvyyden vaihdossa:" & vbCrLf & Err.Description, vbCritical, "Virhe"
    Err.Clear
End Sub