VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysDISTRIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub Form_Load()
'Dim DB As DAO.Database
'Dim Taulu As DAO.Recordset
'Dim i As Integer
'Dim j As Integer
'Set DB = CurrentDb
'Set Taulu = DB.OpenRecordset("DISTRIBUTION")
'For i = 5 To Taulu.Fields.Count - 1
'  For j = 0 To Me.Controls.Count - 1
'    If Me.Controls(j).ControlName = "C" & i + 1 Then
'      With Me.Controls(j)
'        .Caption = Taulu.Fields(i).Name
'        .ControlTipText = Taulu.Fields(i).Properties(23).Value
'        .Visible = True
'        .Width = 1134
'        .Left = 6917 + ((i - 5) * 1134)
'      End With
'    ElseIf Me.Controls(j).ControlName = "D" & i + 1 Then
'      With Me.Controls(j)
'        .ControlSource = Taulu.Fields(i).Name
'        .Visible = True
'        .Width = 1134
'        .Left = 6917 + ((i - 5) * 1134)
'      End With
'    End If
'  Next j
'Next i
'End Sub
