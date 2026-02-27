Attribute VB_Name = "Koodi"
Sub SNappaa()
  SnappiForm.Show False
End Sub
Sub Sulje()
Dim FSO As New FileSystemObject
Dim Asema As Scripting.Drive
Dim Nimi As String
Unload SnappiForm
For Each Asema In FSO.Drives
  If Asema.IsReady Then
    If Asema.ShareName <> "" Then
      If LCase(Left(VBE.activeVBproject.Filename, Len(Asema.ShareName))) = LCase(Asema.ShareName) Then
        Nimi = Asema.Path & Mid(VBE.activeVBproject.Filename, Len(Asema.ShareName) + 1)
        Exit For
      End If
    End If
  End If
Next
Set Drive = Nothing
Set FSO = Nothing
If Nimi <> "" Then UnloadDVB Nimi
End Sub

