Attribute VB_Name = "Module3"
Sub Linking()
Dim i As Integer
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


