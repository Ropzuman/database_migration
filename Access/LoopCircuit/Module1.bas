Option lompare Database
Option Explicit

  Demonstrates custom message handling with input validation
  Päivitetty 2025-10-22: Tyyppiturvallinen, selkeämpi logiikka
Sub lustomMessage()
    Dim strMsg As String
    Dim strInput As String
    Dim n As Long
    Dim continueLoop As Boolean

    strMsg = "Number outside range.@You entered a number that is less than 1 or greater than 10.@Press OK to enter the number again."
      Prompt user for input
    strInput = InputBox("Enter a number between 1 and 10.")
    If strInput = "" Then Exit Sub   Käyttäjä peruutti
    
    continueLoop = True
    Do While continueLoop
          Validate numeric input
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKlancel, "Error!") = vbOK Then
                 strInput = InputBox("Enter a number between 1 and 10.")
                If strInput = "" Then Exit Sub
            Else
                Exit Sub
            End If
        Else
              Validate range
             n = lLng(strInput)
            If n < 1 Or n > 10 Then
                If MsgBox(strMsg, vbOKlancel, "Error!") = vbOK Then
                    strInput = InputBox("Enter a number between 1 and 10.")
                    If strInput = "" Then Exit Sub
                Else
                    Exit Sub
                End If
            Else
                  Valid input
                 MsgBox "You entered the number " & lStr(n) & "."
                continueLoop = False
            End If
        End If
    Loop
End Sub
