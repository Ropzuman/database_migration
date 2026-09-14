Option Compare Database
Option Explicit

Private Const xlContinuous As Long = 1
Private Const xlThin As Long = 2
Private Const xlAutomatic As Long = -4105
Private Const xlVAlignTop As Long = -4160
Private Const xlCenter As Long = -4108
Private Const xlLeft As Long = -4131
Private Const TemporaryFolder As Long = 2

Private Sub Report_Open(Cancel As Integer)
Dim xlApp As Object
Dim Alue As Object
Dim WB As Object
Dim Tiedot As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix
Dim FSO As Object
Dim Rivi As Integer
Dim Sar As Integer

On Error GoTo ErrorHandler

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Tiedot = CurrentDb.OpenRecordset("_qryMAINEQCust")
Rivi = 9
  
  If MsgBox("Haluatko generoida Excel listan?" & vbCrLf & "Generointi saattaa kestää noin minuutin.", vbYesNo, "Päälaiteluettelo") = vbYes Then
    SysCmd acSysCmdSetStatus, "Käynnistää Exceliä..."
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
    WB.SaveAs FSO.GetSpecialFolder(TemporaryFolder).Path & "\Päälaiteluettelo.xls"
    MsgBox "Päälaiteluettelo generoitu Exceliin.", vbInformation, "Päälaiteluettelo"
  End If
Set WB = Nothing
Tiedot.Close
Set Tiedot = Nothing
Set xlApp = Nothing
Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not WB Is Nothing Then Set WB = Nothing
    If Not Tiedot Is Nothing Then Tiedot.Close: Set Tiedot = Nothing
    If Not xlApp Is Nothing Then Set xlApp = Nothing
    On Error GoTo 0
    MsgBox "Virhe raportin generoinnissa: " & Err.Description, vbExclamation, "Päälaiteluettelo"
End Sub

