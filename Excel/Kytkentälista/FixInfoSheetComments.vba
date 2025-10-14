'''
' FixInfoSheetComments.vba - Fix incorrect comments on Info sheet
'''

Sub FixInfoSheetComments()
'''
' Fixes the comments on Info sheet to match the correct field names
'''
Dim ws As Worksheet

  Set ws = Worksheets("Info")
  
  ' Remove all existing comments first
  ws.Cells.ClearComments
  
  ' Add correct comments based on typical Info sheet layout
  ' Adjust cell addresses if your layout is different
  
  ' Row 3: Customer
  ws.Range("D3").AddComment "customer"
  
  ' Row 4: Mill
  ws.Range("D4").AddComment "mill"
  
  ' Row 5: Contract No (might be empty)
  ws.Range("D5").AddComment "contractno"
  
  ' Row 6: Project Name
  ws.Range("D6").AddComment "projname"
  
  ' Row 7: Project No
  ws.Range("D7").AddComment "projno"
  
  ' Row 8: Document name
  ws.Range("D8").AddComment "docname3"
  
  ' Row 9: Document ID
  ws.Range("D9").AddComment "metsodocno"
  
  ' Row 10: Date
  ws.Range("D10").AddComment "date"
  
  ' Row 11: Status
  ws.Range("D11").AddComment "status"
  
  ' Row 12: Revision
  ws.Range("D12").AddComment "revid"
  
  ' Row 13: Revision date
  ws.Range("D13").AddComment "revdate"
  
  MsgBox "Info sheet comments have been updated!" & vbCrLf & vbCrLf & _
         "Now run 'Run Check' again to populate the Info sheet.", vbInformation
  
End Sub

Sub ShowInfoSheetComments()
'''
' Shows what comments currently exist on Info sheet
'''
Dim ws As Worksheet
Dim c As Comment
Dim output As String

  Set ws = Worksheets("Info")
  
  output = "Comments on Info sheet:" & vbCrLf & vbCrLf
  
  For Each c In ws.Comments
    output = output & c.Parent.Address & ": '" & c.Text & "'" & vbCrLf
  Next c
  
  If ws.Comments.Count = 0 Then
    output = output & "(No comments found)"
  End If
  
  Debug.Print output
  MsgBox output, vbInformation, "Info Sheet Comments"
  
End Sub
