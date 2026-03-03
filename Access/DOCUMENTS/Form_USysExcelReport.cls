Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysExcelReport
' SOVELLUS : DOCUMENTS — Excel-raporttien luonti
' KUVAUS   : Generoi dokumenttilistausraportin Exceliin valitun
'            raporttityypin ja yksikön mukaan (PDU/FLU, normaali/CN-pyynnöt).
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub Command16_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Me.FMetsoUnit.Value = 1
  Me.FReportType.Value = 1
End Sub

Private Sub OK_Click()
Dim DB As DAO.Database
Dim Kysely As String
Dim Kysely2 As String
Dim xlApp As Excel.Application
Dim ExcelTaulu As DAO.Recordset
Dim DTaulu As DAO.Recordset
Dim Itemi As String

Set DB = CurrentDb
If Me.FReportType.Value = 1 Then
  If Me.FMetsoUnit.Value = 1 Then
    Kysely = "USysqryForExcelPDU"
    Kysely2 = "USysDIForExcelPDU"
  Else
    Kysely = "USysqryForExcelFLU"
    Kysely2 = "USysDIForExcelFLU"
  End If
Else
  If Me.FMetsoUnit.Value = 1 Then
    Kysely = "USysqryForExcelPDU_CNReq"
    Kysely2 = "USysDIForExcelPDU_CNReq"
  Else
    Kysely = "USysqryForExcelFLU_CNReq"
    Kysely2 = "USysDIForExcelFLU_CNReq"
  End If
End If
On Error GoTo Err_OK_Click
      'Tyhjennetään ensin taulukko
      DB.Execute "DELETE * FROM USysTblForExcel;"
      DB.Execute "DELETE * FROM USysDIForExcel;"
      'Täytetään taulukko tiedoilla
      DB.Execute Kysely
      DB.Execute Kysely2
      Set ExcelTaulu = DB.OpenRecordset("USysTblForExcel")
      Do While Not ExcelTaulu.EOF
         Set DTaulu = DB.OpenRecordset("SELECT DocSchedule.[Item No], DocSchedule.Status, DocSchedule.Date FROM (DOCUMENTS INNER JOIN USysDocsInSched ON DOCUMENTS.Counter = USysDocsInSched.TDoc) INNER JOIN DocSchedule ON USysDocsInSched.TSchedItem = DocSchedule.ID WHERE (((DOCUMENTS.Counter)=" & ExcelTaulu.Fields("Counter").Value & "));")
         Do While Not DTaulu.EOF
           ExcelTaulu.Edit
           If DTaulu.Fields("Status").Value = "Preliminary" Then
             Itemi = "CItemP"
           Else
             Itemi = "CItemC"
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
      'Käynnistetään excel ja tiedosto jossa kyseinen makro on (makro käynnistyy automaattisesti)
      Set xlApp = CreateObject("Excel.Application")
      xlApp.Workbooks.Open "l:\projekti\valdivia\sheets\tools\DocumentReport.xls", , True
      xlApp.Visible = True
      If Me.FReportType.Value = 1 Then
        xlApp.Run "AjaKaikki", False
        xlApp.Workbooks("DocumentReport.xls").Close False
      Else
        xlApp.Run "AjaKaikki", True
        xlApp.Workbooks("DocumentReport.xls").Close False
      End If
      Set xlApp = Nothing
Exit_OK_Click:
      DoCmd.Close acForm, "USysExcelReport"
      Exit Sub

Err_OK_Click:
      MsgBox Err.Description
      Resume Exit_OK_Click
End Sub