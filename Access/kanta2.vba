Option Compare Database

Sub CustomMessage()
    Dim strMsg As String, strInput As String

    ' Initialize string.
    strMsg = "Number outside range.@You entered " _
        & "a number that is less than 1 or greater " _
        & "than 10.@Press OK to enter the number " _
        & "again."
    ' Prompt user for input.
    strInput = InputBox("Enter a number between 1 " _
        & "and 10.")
    ' Determine if user chose "Cancel".
    If strInput = "" Then Exit Sub

    ' Validate numeric input
    Do
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKCancel, "Error!") = vbOK Then
                strInput = InputBox("Enter a number between 1 and 10.")
                If strInput = "" Then Exit Sub
            Else
                Exit Sub
            End If
        Else
            Dim n As Long
            n = CLng(strInput)
            If n < 1 Or n > 10 Then
                If MsgBox(strMsg, vbOKCancel, "Error!") = vbOK Then
                    strInput = InputBox("Enter a number between 1 and 10.")
                    If strInput = "" Then Exit Sub
                Else
                    Exit Sub
                End If
            Else
                MsgBox "You entered the number " & CStr(n) & "."
                Exit Do
            End If
        End If
    Loop
End Sub