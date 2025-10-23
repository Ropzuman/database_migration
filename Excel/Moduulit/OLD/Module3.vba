'''
' Module3.vba - Utility macro for toggling visibility of the LINKING sheet
' Used to show/hide the LINKING sheet in the workbook for debugging or user preference.
'''
Sub Linking()
    ' Toggles the visibility of the LINKING sheet in the workbook.
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