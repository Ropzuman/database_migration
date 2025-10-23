Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

' Demonstrates custom message handling with input validation
' Updated 2025-10-22: Type safety, cleaner logic
Sub CustomMessage()
    Dim strMsg As String
    Dim strInput As String
    Dim n As Long
    Dim continueLoop As Boolean

    strMsg = "Number outside range.@You entered a number that is less than 1 or greater than 10.@Press OK to enter the number again."
    
    ' Prompt user for input
    strInput = InputBox("Enter a number between 1 and 10.")
    If strInput = "" Then Exit Sub ' User cancelled
    
    continueLoop = True
    Do While continueLoop
        ' Validate numeric input
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKCancel, "Error!") = vbOK Then
                strInput = InputBox("Enter a number between 1 and 10.")
                If strInput = "" Then Exit Sub
            Else
                Exit Sub
            End If
        Else
            ' Validate range
            n = CLng(strInput)
            If n < 1 Or n > 10 Then
                If MsgBox(strMsg, vbOKCancel, "Error!") = vbOK Then
                    strInput = InputBox("Enter a number between 1 and 10.")
                    If strInput = "" Then Exit Sub
                Else
                    Exit Sub
                End If
            Else
                ' Valid input
                MsgBox "You entered the number " & CStr(n) & "."
                continueLoop = False
            End If
        End If
    Loop
End Sub
