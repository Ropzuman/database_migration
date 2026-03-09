Option Compare Database
Option Explicit

Private Sub Command0_Click()
Dim PWB As Excel.Workbook
Dim UWB As Excel.Workbook
Dim XL As Excel.Application
Dim i As Integer
Dim Rivi As Integer

Set XL = CreateObject("Excel.Application")
Set PWB = XL.Workbooks.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\Tools\LiiteMoottorienOstoehdotukseenTemplate.xls")
PWB.Sheets.Copy
Set UWB = XL.ActiveWorkbook
PWB.Close False
i = 1
Rivi = 6
Me.RecordsetClone.MoveFirst
Do While Not Me.RecordsetClone.EOF
  With UWB.ActiveSheet
    .Cells(Rivi, 1).Value = i
    .Cells(Rivi, 1).VerticalAlignment = xlTop
    .Cells(Rivi, 2).Value = Me.RecordsetClone.Fields(0).Value
    .Cells(Rivi, 2).VerticalAlignment = xlTop
    .Cells(Rivi, 3).Value = Me.RecordsetClone.Fields(1).Value & Me.RecordsetClone.Fields(2).Value
    .Cells(Rivi, 4).Value = Me.RecordsetClone.Fields(3).Value
    .Cells(Rivi, 4).VerticalAlignment = xlTop
    .Cells(Rivi, 5).Value = Me.RecordsetClone.Fields(4).Value
    .Cells(Rivi, 5).VerticalAlignment = xlTop
    .Cells(Rivi, 6).Value = Me.RecordsetClone.Fields(5).Value
    .Cells(Rivi, 6).VerticalAlignment = xlTop
    .Cells(Rivi, 7).Value = Me.RecordsetClone.Fields(6).Value
    .Cells(Rivi, 7).VerticalAlignment = xlTop
    .Cells(Rivi, 8).Value = Me.RecordsetClone.Fields(7).Value
    .Cells(Rivi, 8).VerticalAlignment = xlTop
    .Cells(Rivi, 9).Value = Me.RecordsetClone.Fields(8).Value
    .Cells(Rivi, 9).VerticalAlignment = xlTop
    .Cells(Rivi, 10).Value = Me.RecordsetClone.Fields(9).Value
    .Cells(Rivi, 10).VerticalAlignment = xlTop
  End With
  
  Rivi = Rivi + 1
  i = i + 1
  Me.RecordsetClone.MoveNext
Loop

With XL.Range(UWB.ActiveSheet.Cells(6, 1), UWB.ActiveSheet.Cells(Rivi - 1, 10)).Borders
  .LineStyle = xlContinuous
  .Weight = xlHairline
  .ColorIndex = xlAutomatic
End With

XL.Application.Visible = True
AppActivate XL.Application.Caption
Set PWB = Nothing
Set UWB = Nothing
Set XL = Nothing
End Sub

Private Sub Command3_Click()
Dim Taulu As Recordset
Set Taulu = CurrentDB.OpenRecordset("Select * FROM qryMottilaus2 WHERE Elementti='340301'")
Debug.Print Taulu.Fields("Teksti").Value & Taulu.Fields("CustPos").Value

Debug.Print Taulu.Fields("CustPos").Value
Set Taulu = Nothing
End Sub

Private Sub Form_Load()
  CurrentDB.Execute "DROP TABLE TblMotorsOptions"
  CurrentDB.Execute "MakeTblMotorsOptions"
  Valm.Value = Valm.ItemData(0)
  Valm_AfterUpdate
End Sub

Private Sub Valm_AfterUpdate()
  Me.Filter = "Manuf ='" & Valm.Value & "'"
  Me.FilterOn = True
End Sub
