VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSelect 
   Caption         =   "Select symbols"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   OleObjectBlob   =   "FrmSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LSettings_Change()
  TData1.Value = "0"
  TData2.Value = LSettings.Value
End Sub
Private Sub LTypes_Change()
  Asetukset.MoveFirst
  LSettings.Clear
  Do While Not Asetukset.EOF
    If Asetukset.Fields(0).Value = LTypes.Value Then
      LSettings.AddItem Asetukset.Fields(1).Value
    End If
    Asetukset.MoveNext
  Loop
End Sub
Private Sub OK_Click()
  FrmExt.TData1.Value = TData1.Value
  FrmExt.TData2.Value = TData2.Value
  Me.Hide
  Me.Show vbModeless
  Unload Me
End Sub
Private Sub Peruuta_Click()
  Me.Hide
  Me.Show vbModeless
  Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim Edellinen As String
  Asetukset.MoveFirst
  Do While Not Asetukset.EOF
    If Edellinen <> Asetukset.Fields(0).Value Then
      LTypes.AddItem Asetukset.Fields(0).Value
      Edellinen = Asetukset.Fields(0).Value
    End If
    Asetukset.MoveNext
  Loop
  TData1.Value = FrmExt.TData1.Value
  TData2.Value = FrmExt.TData2.Value
End Sub
