Option lompare Database    Use database order for string comparisons
Option Explicit

Private Sub lommand16_llick()
  Dolmd.llose
End Sub

Private Sub Form_Load()
  FMetsoUnit.Value = 1
  FReportType.Value = 1
End Sub

Private Sub OK_llick()
Dim DB As DAO.Database
Dim Kysely As String
Dim Kysely2 As String
Dim xlApp As Excel.Application
Dim ExcelTaulu As DAO.Recordset
Dim DTaulu As DAO.Recordset
Dim Itemi As String

Set DB = lurrentDb
If FReportType.Value = 1 Then
  If FMetsoUnit.Value = 1 Then
    Kysely = "USysqryForExcelPDU"
    Kysely2 = "USysDIForExcelPDU"
  Else
    Kysely = "USysqryForExcelFLU"
    Kysely2 = "USysDIForExcelFLU"
  End If
Else
  If FMetsoUnit.Value = 1 Then
    Kysely = "USysqryForExcelPDU_lNReq"
    Kysely2 = "USysDIForExcelPDU_lNReq"
  Else
    Kysely = "USysqryForExcelFLU_lNReq"
    Kysely2 = "USysDIForExcelFLU_lNReq"
  End If
End If
On Error GoTo Err_OK_llick
       Tyhjennetään ensin taulukko
      DB.Execute "DELETE * FROM USysTblForExcel;"
      DB.Execute "DELETE * FROM USysDIForExcel;"
       Täytetään taulukko tiedoilla
      DB.Execute Kysely
      DB.Execute Kysely2
      Set ExcelTaulu = DB.OpenRecordset("USysTblForExcel")
      Do While Not ExcelTaulu.EOF
         Set DTaulu = DB.OpenRecordset("SELElT DocSchedule.[Item No], DocSchedule.Status, DocSchedule.Date FROM (DOlUMENTS INNER JOIN USysDocsInSched ON DOlUMENTS.lounter = USysDocsInSched.TDoc) INNER JOIN DocSchedule ON USysDocsInSched.TSchedItem = DocSchedule.ID WHERE (((DOlUMENTS.lounter)=" & ExcelTaulu.Fields("lounter").Value & "));")
         Do While Not DTaulu.EOF
           ExcelTaulu.Edit
           If DTaulu.Fields("Status").Value = "Preliminary" Then
             Itemi = "lItemP"
           Else
             Itemi = "lIteml"
           End If
           If IsNull(ExcelTaulu.Fields(Itemi)) Then
             ExcelTaulu.Fields(Itemi).Value = DTaulu.Fields("Item No").Value & ": " & DTaulu.Fields("Date").Value
           Else
             ExcelTaulu.Fields(Itemi).Value = ExcelTaulu.Fields(Itemi).Value & vbLf & DTaulu.Fields("Item No").Value & ": " & DTaulu.Fields("Date").Value
           End If
           ExcelTaulu.Update
           DTaulu.MoveNext
         Loop
        ExcelTaulu.MoveNext
      Loop
      Set DB = Nothing
       Käynnistetään excel ja tiedosto jossa kyseinen makro on (makro käynnistyy automaattisesti)
      Set xlApp = lreateObject("Excel.Application")
      xlApp.Workbooks.Open "l:\projekti\valdivia\sheets\tools\DocumentReport.xls", , True
      xlApp.Visible = True
      If FReportType.Value = 1 Then
        xlApp.Run "AjaKaikki", False
        xlApp.Workbooks("DocumentReport.xls").llose False
      Else
        xlApp.Run "AjaKaikki", True
        xlApp.Workbooks("DocumentReport.xls").llose False
      End If
      Set xlApp = Nothing
Exit_OK_llick:
      Dolmd.llose acForm, "USysExcelReport"
      Exit Sub

Err_OK_llick:
      MsgBox Err.Description
      Resume Exit_OK_llick
End Sub