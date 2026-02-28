Option Compare Database
Option Explicit

'==============================================================================
' Report: MOOTTORIT (Motors)
' Purpose: Export motor list to Excel using template
' Updated: 2025-11-11 - Added DAO typing, documented hard-coded paths
'
' HARD-CODED PATHS - Santa Fe Project 220018:
'   Template: N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MotorTEMPLATE.xls
'   Output: N:\whldata\Projekti\Santa Fe 220018\Sahko\SAHKOSUUNN\1A  Customer File\VL  3  Motor and Instrument List\MOTOR LIST\A4-2090-210-03-0101.xls
'
' Note: These paths are specific to the Santa Fe 220018 project.
' For new projects, update these paths in the Report_Open procedure.
'==============================================================================

Private Sub Report_Open(Cancel As Integer)
Dim xlApp As Excel.Application
Dim Alue As Excel.Range
Dim WB As Excel.Workbook
Dim Tiedot As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Dim FSO As New FileSystemObject
Dim Rivi As Integer
Dim Sar As Integer
Set Tiedot = CurrentDb.OpenRecordset("_qryMOTOR")
Rivi = 10
  
  If MsgBox("Haluatko generoida Excel listan?" & vbCrLf & "Generointi saattaa kestää noin minuutin.", vbYesNo, "Moottorilista") = vbYes Then
    SysCmd acSysCmdSetStatus, "Käynnistää Exceliä..."
    Set xlApp = CreateObject("Excel.Application")
    Set WB = xlApp.Workbooks.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MotorTEMPLATE.xls", , True)
    SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin..."
    xlApp.ScreenUpdating = False
    Do While Not Tiedot.EOF
      SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin... Tietue " & Rivi - 9 & "/" & Tiedot.RecordCount
      For Sar = 1 To 13
        WB.Sheets(1).Cells(Rivi, Sar).Value = Tiedot.Fields(Sar - 1).Value
      Next Sar
      Rivi = Rivi + 1
      Tiedot.MoveNext
    Loop
    Set Alue = WB.Sheets(1).Range(WB.Sheets(1).Cells(10, 1), WB.Sheets(1).Cells(Rivi - 1, 13))
    With Alue.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Alue.RowHeight = 27.75
    Alue.VerticalAlignment = xlVAlignTop
    Alue.HorizontalAlignment = xlCenter
    WB.Sheets(1).Range(WB.Sheets(1).Cells(10, 2), WB.Sheets(1).Cells(Rivi - 1, 2)).HorizontalAlignment = xlLeft
    Set Alue = Nothing
    xlApp.ScreenUpdating = True
    xlApp.Visible = True
    'WB.SaveAs FSO.GetSpecialFolder(TemporaryFolder).Path & "\A4-2090-210-03-0101.xls"
    WB.SaveAs "N:\whldata\Projekti\Santa Fe 220018\Sahko\SAHKOSUUNN\1A  Customer File\VL  3  Motor and Instrument List\MOTOR LIST\A4-2090-210-03-0101.xls"
    MsgBox "Moottorilista Generoitu Exceliin.", vbInformation, "Moottorilista"
  End If
Set WB = Nothing
Set Tiedot = Nothing
Set xlApp = Nothing
End Sub

