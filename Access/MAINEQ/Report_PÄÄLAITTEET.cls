VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_PƒƒLAITTEET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
Dim xlApp As Excel.Application
Dim Alue As Excel.Range
Dim WB As Excel.Workbook
Dim Tiedot As Recordset
Dim FSO As New FileSystemObject
Dim Rivi As Integer
Dim Sar As Integer
Set Tiedot = CurrentDB.OpenRecordset("_qryMAINEQCust")
Rivi = 9
  
  If MsgBox("Haluatko generoida Excel listan?" & vbCrLf & "Generointi saattaa kest‰‰ noin minuutin.", vbYesNo, "P‰‰laiteluettelo") = vbYes Then
    SysCmd acSysCmdSetStatus, "K‰ynnist‰‰ Exceli‰..."
    Set xlApp = CreateObject("Excel.Application")
    Set WB = xlApp.Workbooks.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MainEqTEMPLATE.xls", , True)
    SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin..."
    xlApp.ScreenUpdating = False
    xlApp.Sheets(1).Cells.NumberFormat = "@"
    Tiedot.MoveLast
    Tiedot.MoveFirst
    Do While Not Tiedot.EOF
      SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin... Tietue " & Rivi - 8 & "/" & Tiedot.RecordCount
      For Sar = 1 To 13
        WB.Sheets(1).Cells(Rivi, Sar).Value = Tiedot.Fields(Sar - 1).Value
      Next Sar
      Rivi = Rivi + 1
      Tiedot.MoveNext
    Loop
    Set Alue = WB.Sheets(1).Range(WB.Sheets(1).Cells(9, 1), WB.Sheets(1).Cells(Rivi - 1, 13))
    With Alue.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Alue.RowHeight = 18.75
    Alue.VerticalAlignment = xlVAlignTop
    Alue.HorizontalAlignment = xlCenter
    WB.Sheets(1).Range(WB.Sheets(1).Cells(9, 3), WB.Sheets(1).Cells(Rivi - 1, 3)).HorizontalAlignment = xlLeft
    Set Alue = Nothing
    xlApp.ScreenUpdating = True
    xlApp.Visible = True
    WB.SaveAs FSO.GetSpecialFolder(TemporaryFolder).Path & "\P‰‰laiteluettelo.xls"
    MsgBox "P‰‰laiteluettelo generoitu Exceliin.", vbInformation, "P‰‰laiteluettelo"
  End If
Set WB = Nothing
Set Tiedot = Nothing
Set xlApp = Nothing
End Sub

