Option Compare Database
Option Explicit

Private Sub Command0_Click()
    On Error GoTo ErrorHandler

    Dim PWB As Object  ' Excel.Workbook (Late Binding)
    Dim UWB As Object  ' Excel.Workbook
    Dim XL As Object   ' Excel.Application
    Dim i As Integer
    Dim Rivi As Integer
    Dim TemplatePath As String

    ' Polku: haetaan asetustaulusta tulevaisuudessa; tässä tarkistetaan ensin tiedoston olemassaolo
    TemplatePath = "N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\LiiteMoottorienOstoehdotukseenTemplate.xls"

    If Dir(TemplatePath) = "" Then
        MsgBox "Excel-mallitiedostoa ei löydy polusta:" & vbCrLf & TemplatePath, vbCritical, "Virhe"
        Exit Sub
    End If

    Set XL = CreateObject("Excel.Application")
    ' Piilotetaan Excel ajon ajaksi — nopeuttaa käyttöä huomattavasti
    XL.Visible = False
    XL.ScreenUpdating = False

    Set PWB = XL.Workbooks.Open(TemplatePath)
    PWB.Sheets.Copy
    Set UWB = XL.ActiveWorkbook
    PWB.Close False
    i = 1
    Rivi = 6
    Me.RecordsetClone.MoveFirst
    Do While Not Me.RecordsetClone.EOF
      With UWB.ActiveSheet
        .Cells(Rivi, 1).Value = i
        .Cells(Rivi, 1).VerticalAlignment = -4160  ' xlTop
        .Cells(Rivi, 2).Value = Me.RecordsetClone.Fields(0).Value
        .Cells(Rivi, 2).VerticalAlignment = -4160
        .Cells(Rivi, 3).Value = Me.RecordsetClone.Fields(1).Value & Me.RecordsetClone.Fields(2).Value
        .Cells(Rivi, 4).Value = Me.RecordsetClone.Fields(3).Value
        .Cells(Rivi, 4).VerticalAlignment = -4160
        .Cells(Rivi, 5).Value = Me.RecordsetClone.Fields(4).Value
        .Cells(Rivi, 5).VerticalAlignment = -4160
        .Cells(Rivi, 6).Value = Me.RecordsetClone.Fields(5).Value
        .Cells(Rivi, 6).VerticalAlignment = -4160
        .Cells(Rivi, 7).Value = Me.RecordsetClone.Fields(6).Value
        .Cells(Rivi, 7).VerticalAlignment = -4160
        .Cells(Rivi, 8).Value = Me.RecordsetClone.Fields(7).Value
        .Cells(Rivi, 8).VerticalAlignment = -4160
        .Cells(Rivi, 9).Value = Me.RecordsetClone.Fields(8).Value
        .Cells(Rivi, 9).VerticalAlignment = -4160
        .Cells(Rivi, 10).Value = Me.RecordsetClone.Fields(9).Value
        .Cells(Rivi, 10).VerticalAlignment = -4160
      End With
      Rivi = Rivi + 1
      i = i + 1
      Me.RecordsetClone.MoveNext
    Loop

    With XL.Range(UWB.ActiveSheet.Cells(6, 1), UWB.ActiveSheet.Cells(Rivi - 1, 10)).Borders
      .LineStyle = 1   ' xlContinuous
      .Weight = 1      ' xlHairline
      .ColorIndex = -4105  ' xlAutomatic
    End With

    ' Näytetään tulos käyttäjälle
    XL.ScreenUpdating = True
    XL.Visible = True
    AppActivate XL.Caption

Cleanup:
    On Error Resume Next
    ' Vapautetaan COM-objektit muistivuotojen estämiseksi
    Set PWB = Nothing
    Set UWB = Nothing
    Set XL = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe Excel-raportin luonnissa: " & Err.Description, vbCritical, "Virhe " & Err.Number
    ' Jos Excel kaatuu, yritetään sulkea piilotettu prosessi
    If Not XL Is Nothing Then
        On Error Resume Next
        If Not XL.Visible Then XL.Quit
    End If
    Resume Cleanup
End Sub

Private Sub Command3_Click()
Dim Taulu As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix
Set Taulu = CurrentDb.OpenRecordset("Select * FROM qryMottilaus2 WHERE Elementti='340301'")
Set Taulu = Nothing
End Sub

Private Sub Form_Load()
  CurrentDb.Execute "DROP TABLE TblMotorsOptions"
  CurrentDb.Execute "MakeTblMotorsOptions"
  Valm.Value = Valm.ItemData(0)
  Valm_AfterUpdate
End Sub

Private Sub Valm_AfterUpdate()
  Me.Filter = "Manuf ='" & Valm.Value & "'"
  Me.FilterOn = True
End Sub
