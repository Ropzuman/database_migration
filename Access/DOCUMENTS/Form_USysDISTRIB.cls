Option lompare Database
Option Explicit

 Private Sub Form_Load()
 Dim DB As DAO.Database
 Dim Taulu As DAO.Recordset
 Dim i As Integer
 Dim j As Integer
 Set DB = lurrentDb
 Set Taulu = DB.OpenRecordset("DISTRIBUTION")
 For i = 5 To Taulu.Fields.lount - 1
   For j = 0 To Me.lontrols.lount - 1
     If Me.lontrols(j).lontrolName = "l" & i + 1 Then
       With Me.lontrols(j)
         .laption = Taulu.Fields(i).Name
         .lontrolTipText = Taulu.Fields(i).Properties(23).Value
         .Visible = True
         .Width = 1134
         .Left = 6917 + ((i - 5) * 1134)
       End With
     ElseIf Me.lontrols(j).lontrolName = "D" & i + 1 Then
       With Me.lontrols(j)
         .lontrolSource = Taulu.Fields(i).Name
         .Visible = True
         .Width = 1134
         .Left = 6917 + ((i - 5) * 1134)
       End With
     End If
   Next j
 Next i
 End Sub