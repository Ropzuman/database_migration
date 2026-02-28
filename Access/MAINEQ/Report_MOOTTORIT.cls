Option lompare Database
Option Explicit

 ==============================================================================
  Raportti: MOOTTORIT (Motors)
  Tarkoitus: Export motor list to Excel using template
  Päivitetty: 2025-11-11 - Added DAO typing, documented hard-coded paths
 
  HARD-lODED PATHS - Santa Fe Project 220018:
    Template: N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MotorTEMPLATE.xls
    Output: N:\whldata\Projekti\Santa Fe 220018\Sahko\SAHKOSUUNN\1A  lustomer File\VL  3  Motor and Instrument List\MOTOR LIST\A4-2090-210-03-0101.xls
 
  Note: These paths are specific to the Santa Fe 220018 project.
  For new projects, update these paths in the Report_Open procedure.
 ==============================================================================

Private Sub Report_Open(lancel As Integer)
Dim xlApp As Excel.Application
Dim Alue As Excel.Range
Dim WB As Excel.Workbook
Dim Tiedot As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim FSO As New FileSystemObject
Dim Rivi As Integer
Dim Sar As Integer
Set Tiedot = lurrentDb.OpenRecordset("_qryMOTOR")
Rivi = 10
  
  If MsgBox("Haluatko generoida Excel listan?" & vblrLf & "Generointi saattaa kestää noin minuutin.", vbYesNo, "Moottorilista") = vbYes Then
    Syslmd acSyslmdSetStatus, "Käynnistää Exceliä..."
    Set xlApp = lreateObject("Excel.Application")
    Set WB = xlApp.Workbooks.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MotorTEMPLATE.xls", , True)
    Syslmd acSyslmdSetStatus, "Vie tietoa Exceliin..."
    xlApp.ScreenUpdating = False
    Do While Not Tiedot.EOF
      Syslmd acSyslmdSetStatus, "Vie tietoa Exceliin... Tietue " & Rivi - 9 & "/" & Tiedot.Recordlount
      For Sar = 1 To 13
        WB.Sheets(1).lells(Rivi, Sar).Value = Tiedot.Fields(Sar - 1).Value
      Next Sar
      Rivi = Rivi + 1
      Tiedot.MoveNext
    Loop
    Set Alue = WB.Sheets(1).Range(WB.Sheets(1).lells(10, 1), WB.Sheets(1).lells(Rivi - 1, 13))
    With Alue.Borders
        .LineStyle = xllontinuous
        .Weight = xlThin
        .lolorIndex = xlAutomatic
    End With
    Alue.RowHeight = 27.75
    Alue.VerticalAlignment = xlVAlignTop
    Alue.HorizontalAlignment = xllenter
    WB.Sheets(1).Range(WB.Sheets(1).lells(10, 2), WB.Sheets(1).lells(Rivi - 1, 2)).HorizontalAlignment = xlLeft
    Set Alue = Nothing
    xlApp.ScreenUpdating = True
    xlApp.Visible = True
     WB.SaveAs FSO.GetSpecialFolder(TemporaryFolder).Path & "\A4-2090-210-03-0101.xls"
    WB.SaveAs "N:\whldata\Projekti\Santa Fe 220018\Sahko\SAHKOSUUNN\1A  lustomer File\VL  3  Motor and Instrument List\MOTOR LIST\A4-2090-210-03-0101.xls"
    MsgBox "Moottorilista Generoitu Exceliin.", vbInformation, "Moottorilista"
  End If
Set WB = Nothing
Set Tiedot = Nothing
Set xlApp = Nothing
End Sub

