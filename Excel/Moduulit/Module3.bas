'''
' Module3.vba - Apumakro LINKING-sheetin näkyvyyden vaihtamiseen
' Käytetään LINKING-sheetin näyttämiseen/piilottamiseen työkirjassa debuggausta tai käyttäjän mieltymyksiä varten.
'''
Sub Linking()
    ' Vaihtaa LINKING-sheetin näkyvyyden työkirjassa.
    Dim i As Long
    For i = 1 To Sheets.Count
        If LCase(Sheets(i).Name) = "linking" Then
            If Sheets(i).Visible = True Then
                Sheets(i).Visible = False
            Else
                Sheets(i).Visible = True
            End If
        End If
    Next i
End Sub