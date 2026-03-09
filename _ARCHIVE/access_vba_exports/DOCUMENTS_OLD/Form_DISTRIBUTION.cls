Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Add_Click() 'Lis�� Valitut Dokumentit valittuun Distribuutioon
  If Nz(TDistr.Value) = "" Then
    MsgBox "Select Distribution First!", vbCritical, "Error!"
  ElseIf Docs.ItemsSelected.Count = 0 Then
    MsgBox "Select Documents!", vbCritical, "Error!"
  Else
    If MsgBox("Have you checked revisions for documents that your are about to add?", vbYesNo, "Add Documents to Distribution") = vbYes Then
      Dim DB As Database
      Dim taulu As Recordset
      Dim Taulu2 As Recordset
      Dim i As Integer
      Dim j As Integer
      Dim TRev As String
      Dim TMetsoRev As String
      Dim KaikkiOK As Boolean
      KaikkiOK = True
      Set DB = CurrentDb
      Set taulu = DB.OpenRecordset("SELECT Date From USysDISTRIBUTION WHERE ID=" & TDistr.Value)
'      If Format(Taulu.Fields(0).Value, "#") < Format(Date, "#") Then
'        MsgBox "You can not add document to the old distribution!" & vbCrLf _
'             & "Date is now: " & Date & " and distributions date is: " & Taulu.Fields(0).Value, vbCritical, "Error!"
'      Else
        For i = 0 To Docs.ItemsSelected.Count - 1
          j = Docs.ItemData(Docs.ItemsSelected.Item(i))
          'Etsit��n dokumentti Documents taulukosta
          Set Taulu2 = DB.OpenRecordset("SELECT Rev, Counter, Status From DOCUMENTS WHERE Counter=" & j)
          If IsNull(Taulu2.Fields("Rev")) Then
            KaikkiOK = False
          Else
            'Tarkistetaan ettei dokumenttia olla viel� lis�tty t�h�n l�hetykseen
            Set taulu = DB.OpenRecordset("SELECT * FROM SentDocuments WHERE TDoc=" & j & " AND TDistr=" & TDistr.Value)
            If Not taulu.EOF Then  'Dokumentti on jo lis�tty
              taulu.Close
            Else 'dokumenttia ei ole viel� kyseisess� l�hetyksess�, joten lis�t��n se
              With taulu
                .AddNew
                .Fields(0) = j
                .Fields(1) = TDistr.Value
                .Fields(2) = HaeRevisio(Taulu2.Fields("Rev"))
                .Fields(3) = Taulu2.Fields("Status")
                .Update
                .Close
              End With
            End If
          End If
        Next i
'      End If
      If KaikkiOK = False Then
        MsgBox "At least one of the documents did not have Revision!" & vbCrLf & "Please check the revision and try again.", vbCritical, "Add to DISTRIBUTION"
      End If
      Set taulu = Nothing
      Set Taulu2 = Nothing
      DB.Close
      Set DB = Nothing
      USysDocs.Requery
    End If
  End If
End Sub

Private Sub BExit_Click()
  DoCmd.Close
End Sub
Private Sub Command23_Click()
'Tehd��n ensin kysely transmittaalia varten
  If TDistr.Value = "" Or IsNull(TDistr.Value) Then
    MsgBox "Select Distribution First!", vbCritical, "Error!"
  Else
    DoCmd.OpenReport "TRANSMITTAL", acViewPreview, , "ID=" & TDistr.Value
  End If
End Sub
Private Sub Command27_Click()
'Tehd��n uusi DISTRIBUUTIO
    Common = ""
    DoCmd.OpenForm "USysNewDISTRIBUTION"
    TDistr.Requery
    Me.Refresh
    If Common <> "" Then TDistr.Value = Common
End Sub
Private Sub ExcelRep(Ehto As String)
Dim xlApp As Excel.Application

Dim Polku As String
Polku = CurrentDb.Name
Polku = Left(Polku, Len(Polku) - Len(Dir(Polku)))
  
  Set xlApp = CreateObject("Excel.Application")
  SysCmd acSysCmdSetStatus, "Starting Excel..."
  xlApp.Workbooks.Open Polku & "DocumentReport.xls", , True
  xlApp.Visible = True
  SysCmd acSysCmdSetStatus, "Running Excelmacros..."
  xlApp.Run "GeneroiRaportti", Ehto
  xlApp.WindowState = xlMinimized
  xlApp.Workbooks("DocumentReport.xls").Close False
  MsgBox "Excel report generated." & vbCrLf & "You can now save it as you wish.", vbInformation, "Generate Excelreport"
  SysCmd acSysCmdClearStatus
  xlApp.WindowState = xlMaximized
  xlApp.ScreenUpdating = True
  Set xlApp = Nothing
End Sub
Private Sub Command34_Click()
If IsNull(TDistr.Value) Then
  MsgBox "Select distribution first!", vbCritical, "Excel report"
Else
  ExcelRep TDistr.Value
End If
End Sub

Private Sub Command36_Click()
DoCmd.OpenForm "Reserve Number"
End Sub

Private Sub EditDistrib_Click()
  'Muokataan vastaanottajia kyseisess� l�hetyksess�
  If TDistr.Value = "" Or IsNull(TDistr.Value) Then
    MsgBox "Select Distribution First!", vbCritical, "Error!"
  Else
    Common = TDistr.Value
    DoCmd.OpenForm "USysEditDISTRIBUTION", , , , , acDialog
    TDistr.Requery
    FRecip.Requery
  End If
End Sub
Private Sub Form_Load()
Dim taulu As Recordset
  Set taulu = CurrentDb.OpenRecordset("ProjInfo")
  tProjekti.Caption = taulu("Name")
  taulu.Close
  Set taulu = Nothing
End Sub

Private Sub Form_Resize()
    If Me.WindowWidth > 8000 Then
      Logo.Left = Me.WindowWidth - 2000
      VPalkki.Width = Me.WindowWidth - 100
      Docs.Width = Me.WindowWidth - 1100
      BExit.Left = Me.WindowWidth - 1730
      USysDocs.Width = Me.WindowWidth - 1100
      FRecip.Width = Me.WindowWidth - 1100
    End If
    If Me.WindowWidth > 9500 Then
      Info.Left = Me.WindowWidth - 4150
    End If
End Sub
