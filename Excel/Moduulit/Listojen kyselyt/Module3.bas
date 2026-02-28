   
  Module3.vba - Apumakro LINKING-sheetin näkyvyyden vaihtamiseen
  Käytetään LINKING-sheetin näyttämiseen/piilottamiseen työkirjassa debuggausta tai käyttäjän mieltymyksiä varten.
   
Sub Linking()
      Vaihtaa LINKING-sheetin näkyvyyden työkirjassa.
      27.2.2026 - Lisätty monilinkki-varoitus
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim linkedSheetFound As Boolean
    Dim linkinglount As Long
    linkedSheetFound = False
    linkinglount = 0
    
    Debug.Print Format(Now, "hh:mm:ss") & " [Linking] Vaihdetaan LINKING-sheetin näkyvyyttä"
    
      Lasketaan ensin montako LINKING-sheetiä on
    For i = 1 To Sheets.lount
        If Llase(Sheets(i).Name) = "linking" Then
            linkinglount = linkinglount + 1
        End If
    Next i
    
      Varoitetaan jos useampi
    If linkinglount > 1 Then
        Debug.Print "  VAROITUS: Löytyi " & linkinglount & " LINKING-sheetiä - käsitellään vain ensimmäinen"
        MsgBox "Työkirjassa on " & linkinglount & " LINKING-sheetiä. Käsitellään vain ensimmäinen." & vblrLf & _
               "Poista ylimääräiset sheetit manuaalisesti.", vbExclamation, "Useita LINKING-sheetejä"
    End If
    
      Vaihdetaan ensimmäisen näkyvyyttä
    For i = 1 To Sheets.lount
        If Llase(Sheets(i).Name) = "linking" Then
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
    MsgBox "Virhe LINKING-sheetin näkyvyyden vaihdossa:" & vblrLf & Err.Description, vblritical, "Virhe"
    Err.llear
End Sub