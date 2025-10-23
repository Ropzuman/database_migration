VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MotorTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
If [MotorManuf] = "SIEMENS" Then
    Nayta1
    If [MountType] <> "B3" Then
        Nayta
    Else
        Piilota
    End If
Else
    Piilota
    Piilota1
End If

End Sub
Private Sub Piilota()
    Me.SiemensConstrCodeLastPosition.Visible = False
    Me.Label95.Visible = False
End Sub
Private Sub Piilota1()
    'Me.Options_SubForm.Visible = False
End Sub
Private Sub Nayta()
    Me.SiemensConstrCodeLastPosition.Visible = True
    Me.Label95.Visible = True
End Sub
Private Sub Nayta1()
    'Me.Options_SubForm.Visible = True
End Sub

Private Sub Form_Resize()
  If Me.WindowWidth > 1000 Then
    Motors_Subform.Width = Me.WindowWidth - 500
  End If
  'If Me.WindowWidth > 1000 Then
  '  Options_SubForm.Width = Me.WindowWidth - 1000
 ' End If
  'If Me.WindowHeight > 6000 Then
  '  Motors_Subform.Height = Me.WindowHeight - 5700
  'End If
End Sub

Private Sub RevHistory_Click()
  MuutaRevisio
End Sub

Private Sub Revisio_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
'  TRevHist = IIf(IsNull(Rev.Value), "", Rev.Value)
'  TRevDesc = IIf(IsNull(RevHistory.Value), "", RevHistory.Value)
'  DoCmd.OpenForm "USysRevText", , , , , acDialog
'  If TRevHist <> "" And TRevDesc <> "" Then
'    Rev.Value = TRevHist
'    RevHistory.Value = TRevDesc
'  End If
'  Me.Refresh
End Sub

