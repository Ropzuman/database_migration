Option Compare Database
Option Explicit

' Esimerkki mukautetusta syötteen validoinnista ja viestinkutsusta
' Päivitetty 2025-10-22: Tyyppiturvallinen koodi, siivottu logiikka
Sub CustomMessage()
    Dim strMsg As String
    Dim strInput As String
    Dim n As Long
    Dim continueLoop As Boolean

    strMsg = "Number outside range.@You entered a number that is less than 1 or greater than 10.@Press OK to enter the number again."
    ' Pyydetään käyttäjää syöttämään luku
    strInput = InputBox("Enter a number between 1 and 10.")
    If strInput = "" Then Exit Sub ' Käyttäjä peruutti syötön
    
    continueLoop = True
    Do While continueLoop
        ' Tarkistetaan, onko syöte numeerinen
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKCancel, "Error!") = vbOK Then
                 strInput = InputBox("Enter a number between 1 and 10.")
                If strInput = "" Then Exit Sub
            Else
                Exit Sub
            End If
        Else
            ' Tarkistetaan, onko luku sallitulla välillä 1–10
             n = CLng(strInput)
            If n < 1 Or n > 10 Then
                If MsgBox(strMsg, vbOKCancel, "Error!") = vbOK Then
                    strInput = InputBox("Enter a number between 1 and 10.")
                    If strInput = "" Then Exit Sub
                Else
                    Exit Sub
                End If
            Else
                ' Syöte on kelvollinen
                 MsgBox "You entered the number " & CStr(n) & "."
                continueLoop = False
            End If
        End If
    Loop
End Sub
