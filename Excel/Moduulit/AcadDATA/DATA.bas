Option Explicit

' Päivitetty 2025-10-26: 64-bittinen yhteensopivuus, parannettu virheenkäsittely, suorituskykyoptimoint
' Taulukon kaksoisnapsautustapahtuma: Etsii ja zoomaa AutoCAD-blokin piirustuksessa
' Muutokset: Integer → Long (64-bitti), varhainen sidonta → myöhäinen sidonta (yhteensopivuus)

' ============================================================================
' AutoCAD-vakiot – tarvitaan myöhäistä sidontaa varten
' ============================================================================
' Myöhäistä sidontaa käytettäessä (Object AcadApplication, AcadEntity jne. sijaan)
' AutoCAD-tyyppikirjastoa ei viitata, joten sisäänrakennetut vakiot eivät ole käytettävissä.
' Ne on määriteltävä manuaalisesti niiden numeroarvoilla.
' Lähde: Autodesk AutoCAD ActiveX/VBA Reference Documentation
' ============================================================================

Private Const acModelSpace As Long = 1              ' Malliavaruus (ei paperitila)
Private Const acMax As Long = 3                     ' Maksimoi ikkuna


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
    Dim oACAD As Object ' AcadApplication (myöhäinen sidonta – yhteensopivuus)
    Dim Entity As Object ' AcadEntity
    Dim Avataan As Boolean
    Dim OK As Boolean
    Dim Doku As String
    Dim Tiedosto As String
    Dim MinPoint As Variant
    Dim MaxPoint As Variant
    Dim i As Long ' Muutettu Integer → Long 64-bittistä yhteensopivuutta varten
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' Tarkistetaan, että rivillä on dataa
    If Cells(Target.Row, 1).Value = "" Then
        MsgBox "Ei kuvaa valitulla rivillä", vbInformation, "Etsi blokki"
        Cancel = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Yhdistetään käynnissä olevaan AutoCAD-instanssiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Etsi blokki"
        Cancel = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' Haetaan piirustuksen nimi solusta
    Doku = LCase(Cells(Target.Row, 2).Value) & ".dwg"
    
    ' Tarkistetaan, onko oikea piirustus avoinna
    If oACAD.Preferences.System.SingleDocumentMode Then
        ' SDI-tila – vain yksi piirustus voi olla auki kerrallaan
        If LCase(oACAD.ActiveDocument.Name) <> Doku Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        Else
            OK = True
        End If
    Else
        ' MDI-tila – useita piirustuksia voi olla auki yhtä aikaa
        For i = 0 To oACAD.Documents.Count - 1
            If LCase(oACAD.Documents(i).Name) = Doku Then
                oACAD.Documents(i).Activate 
                OK = True
                Exit For
            End If
        Next i
        
        If Not OK Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        End If
    End If
    
    ' Avataan piirustus, jos käyttäjä hyväksyi
    If Avataan Then
        Tiedosto = Cells(Target.Row, 1).Value & "\" & Doku
        
        On Error Resume Next
        oACAD.Documents.Open Tiedosto
        
        If Err.Number = 0 Then
            OK = True
        Else
            MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Tiedosto, vbCritical, "Etsi blokki"
            OK = False
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Zoomataan entiteetin luo, jos piirustus on avoinna
    If OK Then
        oACAD.ActiveDocument.ActiveSpace = acModelSpace
        Set Entity = oACAD.ActiveDocument.HandleToObject(Cells(Target.Row, 4).Value)
        
        ' Haetaan rajauslaatikko ja zoomaataan entiteettiin
        Entity.GetBoundingBox MinPoint, MaxPoint
        oACAD.ActiveDocument.WindowState = acMax
        oACAD.ZoomWindow MinPoint, MaxPoint
        ' Luotettava zoom-ulos myöhäisessä sidonnassa: kokeillaan arvoa 1, sitten 3
        SafeZoomScaled oACAD, 0.5
        
        ' Aktivoidaan AutoCAD-ikkuna etualalle
        On Error Resume Next
        AppActivate oACAD.Caption, True
        On Error GoTo ErrHandler
    End If
    
Cleanup:
    ' Vapautetaan COM-objektit
    Set Entity = Nothing
    Set oACAD = Nothing
    Cancel = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Etsi blokki"
    Resume Cleanup
End Sub