Option Explicit

' Turvallinen apufunktio ZoomScaled-kutsulle myöhäistä sidontaa käytettäessä.
' Eri AutoCAD-versiot voivat käyttää eri numeroarvoa AcZoomScaledRelative-vakiolle (1 tai 3),
' ja myöhäinen sidonta voi nostaa virheen "Invalid argument type in ZoomScaled".
' Funktio kokeilee molemmat arvojärjestyksessä ja ohittaa harmattomat virheet.
Public Sub SafeZoomScaled(ByVal app As Object, ByVal factor As Double)
    On Error Resume Next
    ' Kokeillaan ensin arvoa 1 (yleisimmin acZoomScaledRelative)
    app.ZoomScaled CDbl(factor), CLng(1)
    If Err.Number <> 0 Then
        Err.Clear
        ' Vaihtoehto: arvo 3 (käytetty joissakin Autodeskin dokumentaatioissa)
        app.ZoomScaled CDbl(factor), CLng(3)
    End If
    ' Jos molemmat epäonnistuvat, jätetään näkymä ennalleen
    Err.Clear
    On Error GoTo 0
End Sub
