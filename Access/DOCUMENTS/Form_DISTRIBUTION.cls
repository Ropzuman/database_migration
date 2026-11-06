Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_DISTRIBUTION
' SOVELLUS : DOCUMENTS ó Distribuutioiden hallinta
' KUVAUS   : Hallinnoi l‰hetysten (distribuutioiden) luontia, dokumenttien
'            lis‰‰mist‰ distribuutioihin, TRANSMITTAL-raportin tulostusta
'            sek‰ Excel-raporttien generointia. K‰ytt‰‰ transaktioita
'            atomaarisuuden takaamiseksi.
' PƒIVITETTY: 2026-03-03
'==========================================================================

'''
' Add Documents to Distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support
'''
Private Sub Add_Click()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim Taulu2 As DAO.Recordset
    Dim i As Integer
    Dim j As Long
    Dim TRev As String
    Dim TMetsoRev As String
    Dim KaikkiOK As Boolean
    Dim TransaktioKaynnissa As Boolean
    Dim VirheenKuvaus As String
    
    On Error GoTo ErrorHandler
    
    ' Validation
    If Nz(Me.TDistr.Value) = "" Then
        MsgBox "Select Distribution First!", vbCritical, "Error!"
        Exit Sub
    ElseIf Me.Docs.ItemsSelected.Count = 0 Then
        MsgBox "Select Documents!", vbCritical, "Error!"
        Exit Sub
    End If
    
    If MsgBox("Have you checked revisions for documents that you are about to add?", vbYesNo, "Add Documents to Distribution") <> vbYes Then
        Exit Sub
    End If
    
    KaikkiOK = True
    Set DB = CurrentDb
    
' Aloitetaan transaktio atomaarisuuden takaamiseksi
    DBEngine.BeginTrans
    TransaktioKaynnissa = True

    For i = 0 To Me.Docs.ItemsSelected.Count - 1
        j = Me.Docs.ItemData(Me.Docs.ItemsSelected.Item(i))
        
        ' Find document in DOCUMENTS table
        Set Taulu2 = DB.OpenRecordset("SELECT Rev, Counter, Status From DOCUMENTS WHERE Counter=" & j, dbOpenSnapshot)
        
        If Not Taulu2.EOF Then
            If IsNull(Taulu2.Fields("Rev")) Then
                KaikkiOK = False
            Else
                ' Check if document already added to this distribution
                Set taulu = DB.OpenRecordset("SELECT * FROM SentDocuments WHERE TDoc=" & j & " AND TDistr=" & Me.TDistr.Value, dbOpenDynaset)
                
                If taulu.EOF Then  ' Document not yet in distribution, add it
                    taulu.AddNew
                    taulu.Fields(0) = j
                    taulu.Fields(1) = Me.TDistr.Value
                    taulu.Fields(2) = HaeRevisioPaikallinen(Taulu2.Fields("Rev"))
                    taulu.Fields(3) = Nz(Taulu2.Fields("Status"), "")
                    taulu.Update
                End If
                
                taulu.Close
                Set taulu = Nothing
            End If
        End If
        
        Taulu2.Close
        Set Taulu2 = Nothing
    Next i
    
' Vahvistetaan transaktio
    DBEngine.CommitTrans
    TransaktioKaynnissa = False

    If Not KaikkiOK Then
        MsgBox "Only documents with a revision could be added. At least one document had no revision and was not added." & vbCrLf & "Please check the missing revisions and try again.", vbExclamation, "Add to DISTRIBUTION"
    End If
    
Cleanup:
    If Not taulu Is Nothing Then
        taulu.Close
        Set taulu = Nothing
    End If
    If Not Taulu2 Is Nothing Then
        Taulu2.Close
        Set Taulu2 = Nothing
    End If
    If Not DB Is Nothing Then
        DB.Close
        Set DB = Nothing
    End If
    Me.USysDocs.Requery
    Exit Sub
    
ErrorHandler:
' Per‰ytet‰‰n transaktio virhetilanteessa
    VirheenKuvaus = Err.Description
    If TransaktioKaynnissa Then
        On Error Resume Next
        Err.Clear
        DBEngine.Rollback
        If Err.Number <> 0 Then VirheenKuvaus = VirheenKuvaus & vbCrLf & "Rollback failed: " & Err.Description
        On Error GoTo 0
    End If
    MsgBox "Error adding documents to distribution: " & VirheenKuvaus, vbExclamation, "Add Error"
    Resume Cleanup
End Sub

Private Function HaeRevisioPaikallinen(ByVal Revisio As Variant) As String
    Dim RevisioTeksti As String
    Dim ErotinPos As Long

    If IsNull(Revisio) Then
        HaeRevisioPaikallinen = ""
        Exit Function
    End If

    RevisioTeksti = CStr(Revisio)
    ErotinPos = InStr(RevisioTeksti & " ", " ")
    If ErotinPos > 1 Then
        HaeRevisioPaikallinen = Left$(RevisioTeksti, ErotinPos - 1)
    Else
        HaeRevisioPaikallinen = ""
    End If
End Function

Private Sub BExit_Click()
  DoCmd.Close
End Sub
Private Sub Command23_Click()
' Avataan TRANSMITTAL-raportti valitulle distribuutiolle
  If Me.TDistr.Value = "" Or IsNull(Me.TDistr.Value) Then
    MsgBox "Select Distribution First!", vbCritical, "Error!"
  Else
    DoCmd.OpenReport "TRANSMITTAL", acViewPreview, , "ID=" & Me.TDistr.Value
  End If
End Sub
Private Sub Command27_Click()
' Luodaan uusi distribuutio
    DoCmd.OpenForm "USysNewDISTRIBUTION"
    Me.TDistr.Requery
    Me.Refresh
End Sub
'''
' Generate Excel Report for Distribution or All Documents
' OPTIMIZED 2025-11-09: Added error handling, removed redundant CurrentDb call
' @param Ehto - Distribution ID filter (if empty, generates report for all documents)
'''
Private Sub ExcelRep(Ehto As String)
    Dim xlApp As Object
    Dim ExcelKayttajalla As Boolean
    Const ExcelPienennetty As Long = -4140
    Const ExcelSuurennettu As Long = -4137
    
    On Error GoTo ErrorHandler
    
    Set xlApp = CreateObject("Excel.Application")
    SysCmd acSysCmdSetStatus, "Starting Excel..."
    xlApp.Workbooks.Open "\\proense01\yhteinen\Data\Tools\Excel\DocumentReport.xls", , True
    xlApp.Visible = True
    ExcelKayttajalla = True
    SysCmd acSysCmdSetStatus, "Running Excelmacros..."
    xlApp.Run "GeneroiRaportti", Ehto
    xlApp.WindowState = ExcelPienennetty
    xlApp.Workbooks("DocumentReport.xls").Close False
    MsgBox "Excel report generated." & vbCrLf & "You can now save it as you wish.", vbInformation, "Generate Excelreport"
    SysCmd acSysCmdClearStatus
    xlApp.WindowState = ExcelSuurennettu
    xlApp.ScreenUpdating = True
    
Cleanup:
    If Not xlApp Is Nothing Then
        If Not ExcelKayttajalla Then
            On Error Resume Next
            xlApp.Quit
            On Error GoTo 0
        End If
        Set xlApp = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating Excel report: " & Err.Description, vbExclamation, "Excel Report Error"
    Resume Cleanup
End Sub
Private Sub Command34_Click()
If IsNull(Me.TDistr.Value) Then
  MsgBox "Select distribution first!", vbCritical, "Excel report"
Else
  ExcelRep Me.TDistr.Value
End If
End Sub

Private Sub Command36_Click()
DoCmd.OpenForm "Reserve Number"
End Sub

Private Sub EditDistrib_Click()
  ' Muokataan vastaanottajia kyseisess‰ l‰hetyksess‰
  If Me.TDistr.Value = "" Or IsNull(Me.TDistr.Value) Then
    MsgBox "Select Distribution First!", vbCritical, "Error!"
  Else
        DoCmd.OpenForm "USysEditDISTRIBUTION", , , , , acDialog, CStr(Me.TDistr.Value)
    Me.TDistr.Requery
    Me.FRecip.Requery
  End If
End Sub
'''
' Form Load Event - Initialize form with project information
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Public Sub Form_Load()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set DB = CurrentDb
    Set taulu = DB.OpenRecordset("ProjInfo", dbOpenSnapshot)
    
    If Not taulu.EOF Then
        Me.tProjekti.Caption = Nz(taulu("Name"), "")
    End If
    
Cleanup:
    If Not taulu Is Nothing Then
        taulu.Close
        Set taulu = Nothing
    End If
    If Not DB Is Nothing Then
        DB.Close
        Set DB = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading project information: " & Err.Description, vbExclamation, "Form Load Error"
    Resume Cleanup
End Sub

Private Sub Form_Resize()
    If Me.WindowWidth > 8000 Then
      Me.Logo.Left = Me.WindowWidth - 2000
      Me.VPalkki.Width = Me.WindowWidth - 100
      Me.Docs.Width = Me.WindowWidth - 1100
      Me.BExit.Left = Me.WindowWidth - 1730
      Me.USysDocs.Width = Me.WindowWidth - 1100
      Me.FRecip.Width = Me.WindowWidth - 1100
    End If
    If Me.WindowWidth > 9500 Then
      Me.Info.Left = Me.WindowWidth - 4150
    End If
End Sub