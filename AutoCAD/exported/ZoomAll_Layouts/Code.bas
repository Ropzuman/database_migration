Attribute VB_Name = "Code"
Sub Ekstentti()
Dim i As Integer
On Error Resume Next
For i = 0 To ActiveDocument.Layouts.Count - 1
  ActiveDocument.ActiveLayout = ActiveDocument.Layouts(i)
  ActiveDocument.MSpace = False
  Application.ZoomAll
Next i
Err.Clear
On Error GoTo 0
End Sub
