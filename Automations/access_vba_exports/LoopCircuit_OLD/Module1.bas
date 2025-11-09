Attribute VB_Name = "Module1"
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
    If strInput <> "" Then
        ' Test value of user input.
        Do While strInput < 0 Or strInput > 10
            If MsgBox(strMsg, vbOKCancel, "Error!") = _
                    vbOK Then
                strInput = InputBox("Enter a number between 1 and 10.")
            Else
                Exit Sub
            End If
        Loop
        ' Display user's correct input.
        MsgBox "You entered the number " & strInput & "."
    Else
        Exit Sub
    End If
End Sub


