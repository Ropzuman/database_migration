VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmManValve 
   Caption         =   "Manual Valve Numbering - 28.4.2005"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   OleObjectBlob   =   "FrmManValve.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmManValve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OK_Click()
  'Me.Visible = False
  NumeroiVenttiilit
  Unload Me
End Sub
Private Sub Peruuta_Click()
  Unload Me
End Sub
Private Sub NumeroiVenttiilit()
Dim Joukko As AcadSelectionSet
Dim Pisteet() As Double
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim i As Long
Dim oBlock As AcadBlockReference
'sort
FilterType(0) = 2       'ObjecName
FilterData(0) = "UP050,UP051,UP052,UP053,UP054,UP082,UP084,UP085,UP086,UP087"  'Block Names

Set Joukko = ActiveDocument.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
For i = 0 To Joukko.Count - 1
  Set oBlock = Joukko(i)
  oBlock.AddAttribute
Next i

End Sub

Private Sub TextBox1_Change()

End Sub
