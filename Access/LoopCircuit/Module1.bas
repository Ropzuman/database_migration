Option Compare Database
Option Explicit

' Esimerkki mukautetusta syötteen validoinnista ja viestinkutsusta
' Päivitetty 2025-10-22: Tyyppiturvallinen koodi, siivottu logiikka
' Päivitetty 2026-03-07: Fail-fast / return early -rakenne — poistettu Hadouken-sisennys
Sub CustomMessage()
    Dim strInput As String
    Dim n As Long

    ' Validointisilmukka fail-fast-periaatteella — ei sisäkkäisiä If-Else-haaroja
    Do
        strInput = InputBox("Enter a number between 1 and 10.")
        If strInput = "" Then Exit Sub ' Käyttäjä peruutti syötön

        ' Tarkistus 1: Onko syöte numeerinen?
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKCancel, "Error!") = vbCancel Then Exit Sub
            ' Jatketaan silmukkaan uudella yrityksellä
        Else
            n = CLng(strInput)
            ' Tarkistus 2: Onko luku sallitulla välillä 1–10?
            If n < 1 Or n > 10 Then
                If MsgBox("Number outside range." & vbCrLf & _
                          "You entered a number that is less than 1 or greater than 10.", _
                          vbOKCancel, "Error!") = vbCancel Then Exit Sub
                ' Jatketaan silmukkaan uudella yrityksellä
            Else
                ' Validointi onnistui — poistutaan silmukasta
                MsgBox "You entered valid number: " & n, vbInformation
                Exit Do
            End If
        End If
    Loop
End Sub
