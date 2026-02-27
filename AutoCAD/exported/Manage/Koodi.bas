Attribute VB_Name = "Koodi"
Public oDbx As New AXDB15Lib.AxDbDocument
Sub StartUp()
  Shell "REGSVR32 /s """ & Application.Path & "\axdb15.dll""", 0
  frmTitleManage.Show
End Sub
