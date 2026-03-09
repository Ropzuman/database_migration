Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim xlApp As Object   ' Excel.Application (Late Binding)
    Dim Alue As Object    ' Excel.Range
    Dim WB As Object      ' Excel.Workbook
    Dim WS As Object      ' Excel.Worksheet
    Dim Tiedot As DAO.Recordset
    Dim Rivi As Integer
    Dim ViimeinenRivi As Integer
    Dim TemplatePath As String
    Dim FSO As Object

    TemplatePath = "N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\MainEqTEMPLATE.xls"
    Rivi = 9

    Set Tiedot = CurrentDb.OpenRecordset("_qryMAINEQCust", dbOpenSnapshot)

    If Tiedot.EOF Then
        MsgBox "Ei vietävää dataa kyselyssä _qryMAINEQCust.", vbInformation, "Päälaiteluettelo"
        Cancel = True
        GoTo Cleanup
    End If

    If MsgBox("Haluatko generoida Excel listan?" & vbCrLf & "Generointi saattaa kestää noin minuutin.", vbYesNo, "Päälaiteluettelo") = vbYes Then
        ' Tarkistetaan mallitiedoston olemassaolo ennen Excel-käynnistystä
        If Dir(TemplatePath) = "" Then
            MsgBox "Excel-mallitiedostoa ei löydy:" & vbCrLf & TemplatePath, vbCritical, "Virhe"
            GoTo Cleanup
        End If

        SysCmd acSysCmdSetStatus, "Käynnistää Exceliä..."
        Set xlApp = CreateObject("Excel.Application")
        xlApp.ScreenUpdating = False
        Set WB = xlApp.Workbooks.Open(TemplatePath, , True)
        Set WS = WB.Sheets(1)

        SysCmd acSysCmdSetStatus, "Vie tietoa Exceliin (CopyFromRecordset)..."
        WS.Cells.NumberFormat = "@"

        ' NOPEIN TAPA: pudotetaan koko Recordset kerralla ilman silmukoita
        WS.Cells(Rivi, 1).CopyFromRecordset Tiedot

        ' Lasketaan viimeinen rivi formatointia varten
        Tiedot.MoveLast
        ViimeinenRivi = Rivi + Tiedot.RecordCount - 1

        ' Muotoilu koko alueelle kerralla
        Set Alue = WS.Range(WS.Cells(Rivi, 1), WS.Cells(ViimeinenRivi, 13))
        With Alue.Borders
            .LineStyle = 1    ' xlContinuous
            .Weight = 2       ' xlThin
            .ColorIndex = -4105  ' xlAutomatic
        End With
        Alue.RowHeight = 18.75
        Alue.VerticalAlignment = -4160   ' xlVAlignTop
        Alue.HorizontalAlignment = -4108  ' xlCenter
        WS.Range(WS.Cells(Rivi, 3), WS.Cells(ViimeinenRivi, 3)).HorizontalAlignment = -4131  ' xlLeft
        Set Alue = Nothing

        xlApp.ScreenUpdating = True
        xlApp.Visible = True

        ' Tallennetaan väliaikaiskansioon
        Set FSO = CreateObject("Scripting.FileSystemObject")
        WB.SaveAs FSO.GetSpecialFolder(2).Path & "\Päälaiteluettelo.xls"
        MsgBox "Päälaiteluettelo generoitu Exceliin.", vbInformation, "Päälaiteluettelo"
    End If

Cleanup:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    ' Vapautetaan kaikki COM-objektit — estää haamuprosessit
    Set Alue = Nothing
    Set WS = Nothing
    If Not Tiedot Is Nothing Then Tiedot.Close
    Set Tiedot = Nothing
    Set WB = Nothing
    Set FSO = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe päälaiteluettelon luonnissa: " & Err.Description, vbCritical, "Virhe " & Err.Number
    ' Suljetaan piilotettu Excel-prosessi virhetilanteessa
    If Not xlApp Is Nothing Then
        On Error Resume Next
        If Not xlApp.Visible Then xlApp.Quit
    End If
    Resume Cleanup
End Sub

