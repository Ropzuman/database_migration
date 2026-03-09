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
  Dim xlApp As Object  ' Late binding — vakaampi kuin varhainen sidonta
  Dim ExcelTaulu As DAO.Recordset
  Dim DTaulu As DAO.Recordset
  Dim Itemi As String
  Dim ExcelPolku As String

  On Error GoTo ErrorHandler

  ExcelPolku = "l:\projekti\valdivia\sheets\tools\DocumentReport.xls"

  ' Varmistetaan tiedoston olemassaolo ennen Excelin käynnistystä
  If Dir(ExcelPolku) = "" Then
    MsgBox "Excel report template not found at: " & ExcelPolku, vbCritical
    Exit Sub
  End If

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

  Set DB = CurrentDb

  ' Tyhjennetään väliaikaistaulut ennen täyttöä
  DB.Execute "DELETE * FROM USysTblForExcel;"
  DB.Execute "DELETE * FROM USysDIForExcel;"
  ' Täytetään väliaikaistaulut raporttidatalla
  DB.Execute Kysely
  DB.Execute Kysely2

  Set ExcelTaulu = DB.OpenRecordset("USysTblForExcel")
  Do While Not ExcelTaulu.EOF
    Set DTaulu = DB.OpenRecordset("SELECT DocSchedule.[Item No], DocSchedule.Status, DocSchedule.Date FROM (DOCUMENTS INNER JOIN USysDocsInSched ON DOCUMENTS.Counter = USysDocsInSched.TDoc) INNER JOIN DocSchedule ON USysDocsInSched.TSchedItem = DocSchedule.ID WHERE (((DOCUMENTS.Counter)=" & CLng(ExcelTaulu.Fields("Counter").Value) & "));")
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
    DTaulu.Close
    Set DTaulu = Nothing
    ExcelTaulu.MoveNext
  Loop
  ExcelTaulu.Close
  Set ExcelTaulu = Nothing

  ' Käynnistetään Excel COM-automaatio raporttia varten
  Set xlApp = CreateObject("Excel.Application")
  ' Poistetaan Excelin omat häiritsevät varoitukset automaation ajaksi
  xlApp.DisplayAlerts = False
  xlApp.Workbooks.Open ExcelPolku, UpdateLinks:=False
  xlApp.Visible = True

  If Me.FReportType.Value = 1 Then
    xlApp.Run "AjaKaikki", False
  Else
    xlApp.Run "AjaKaikki", True
  End If
  xlApp.Workbooks("DocumentReport.xls").Close False

Cleanup:
  ' Varmistetaan aina Excel-prosessin puhdistus muistivuodon estämiseksi
  On Error Resume Next
  If Not DTaulu Is Nothing Then
    DTaulu.Close
    Set DTaulu = Nothing
  End If
  If Not ExcelTaulu Is Nothing Then
    ExcelTaulu.Close
    Set ExcelTaulu = Nothing
  End If
  Set DB = Nothing
  If Not xlApp Is Nothing Then
    xlApp.DisplayAlerts = True
    xlApp.Quit  ' Lopetetaan Excel-prosessi — estetään "haamuinstanssit"
    Set xlApp = Nothing
  End If
  DoCmd.Close acForm, "USysExcelReport"
  Exit Sub

ErrorHandler:
  MsgBox "Error communicating with Excel: " & Err.Description, vbCritical
  Resume Cleanup
End Sub