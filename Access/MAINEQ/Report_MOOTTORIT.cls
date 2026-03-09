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
    On Error GoTo ErrorHandler

    Dim xlApp As Object  ' Excel.Application (Late Binding)
    Dim Alue As Object   ' Excel.Range
    Dim WB As Object     ' Excel.Workbook
    Dim Tiedot As DAO.Recordset
    Dim Rivi As Integer
    Dim Sar As Integer
    Dim TemplatePath As String
    Dim TulostusPolku As String

    TemplatePath = "N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MotorTEMPLATE.xls"
    TulostusPolku = "N:\whldata\Projekti\Santa Fe 220018\Sahko\SAHKOSUUNN\1A  Customer File\VL  3  Motor and Instrument List\MOTOR LIST\A4-2090-210-03-0101.xls"

    Set Tiedot = CurrentDb.OpenRecordset("_qryMOTOR", dbOpenSnapshot)
    Rivi = 10

    If MsgBox("Haluatko generoida Excel listan?" & vbCrLf & "Generointi saattaa kestää noin minuutin.", vbYesNo, "Moottorilista") = vbYes Then
        ' Tarkistetaan mallitiedoston olemassaolo ennen Excel-käynnistystä
        If Dir(TemplatePath) = "" Then
            MsgBox "Excel-mallitiedostoa ei löydy:" & vbCrLf & TemplatePath, vbCritical, "Virhe"
            GoTo Cleanup
        End If

        SysCmd acSysCmdSetStatus, "Käynnistää Exceliä..."
        Set xlApp = CreateObject("Excel.Application")
        Set WB = xlApp.Workbooks.Open(TemplatePath, , True)
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
            .LineStyle = 1   ' xlContinuous
            .Weight = 2      ' xlThin
            .ColorIndex = -4105  ' xlAutomatic
        End With
        Alue.RowHeight = 27.75
        Alue.VerticalAlignment = -4160   ' xlVAlignTop
        Alue.HorizontalAlignment = -4108  ' xlCenter
        WB.Sheets(1).Range(WB.Sheets(1).Cells(10, 2), WB.Sheets(1).Cells(Rivi - 1, 2)).HorizontalAlignment = -4131  ' xlLeft
        Set Alue = Nothing
        xlApp.ScreenUpdating = True
        xlApp.Visible = True
        WB.SaveAs TulostusPolku
        MsgBox "Moottorilista Generoitu Exceliin.", vbInformation, "Moottorilista"
    End If

Cleanup:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    ' Vapautetaan kaikki COM-objektit — estää haamuprosessit
    If Not Tiedot Is Nothing Then Tiedot.Close
    Set Alue = Nothing
    Set WB = Nothing
    Set Tiedot = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe moottorilistan luonnissa: " & Err.Description, vbCritical, "Virhe " & Err.Number
    ' Suljetaan piilotettu Excel-prosessi virhetilanteessa
    If Not xlApp Is Nothing Then
        On Error Resume Next
        If Not xlApp.Visible Then xlApp.Quit
    End If
    Resume Cleanup
End Sub

