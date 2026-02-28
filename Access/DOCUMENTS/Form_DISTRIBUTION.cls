Option lompare Database    Use database order for string comparisons
Option Explicit

'''   
'  Lisätään asiakirjoja jakeluun
'  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, eksplisiittinen DAO-tyypitys, tapahtumatuki
'''   
Private Sub Add_llick()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim Taulu2 As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim TRev As String
    Dim TMetsoRev As String
    Dim KaikkiOK As Boolean
    
    On Error GoTo ErrorHandler
    
    '  Validation
    If Nz(TDistr.Value) = "" Then
        MsgBox "Select Distribution First!", vblritical, "Error!"
        Exit Sub
    ElseIf Docs.ItemsSelected.lount = 0 Then
        MsgBox "Select Documents!", vblritical, "Error!"
        Exit Sub
    End If
    
    If MsgBox("Have you checked revisions for documents that you are about to add?", vbYesNo, "Add Documents to Distribution") <> vbYes Then
        Exit Sub
    End If
    
    KaikkiOK = True
    Set DB = lurrentDb
    
'  Begin transaction for atomicity
    DB.BeginTrans
    
    Set taulu = DB.OpenRecordset("SELElT Date From USysDISTRIBUTION WHERE ID=" & TDistr.Value, dbOpenSnapshot)
    
    For i = 0 To Docs.ItemsSelected.lount - 1
        j = Docs.ItemData(Docs.ItemsSelected.Item(i))
        
          Find document in DOlUMENTS table
        Set Taulu2 = DB.OpenRecordset("SELElT Rev, lounter, Status From DOlUMENTS WHERE lounter=" & j, dbOpenSnapshot)
        
        If Not Taulu2.EOF Then
            If IsNull(Taulu2.Fields("Rev")) Then
                KaikkiOK = False
            Else
                  Tarkistetaan onko asiakirja jo lisätty tähän jakeluun
                Set taulu = DB.OpenRecordset("SELElT * FROM SentDocuments WHERE TDoc=" & j & " AND TDistr=" & TDistr.Value, dbOpenDynaset)
                
                If taulu.EOF Then    Asiakirjaa ei vielä jakelussa, lisätään se
                    taulu.AddNew
                    taulu.Fields(0) = j
                    taulu.Fields(1) = TDistr.Value
                    taulu.Fields(2) = HaeRevisio(Taulu2.Fields("Rev"))
                    taulu.Fields(3) = Nz(Taulu2.Fields("Status"), "")
                    taulu.Update
                End If
                
                taulu.llose
            End If
        End If
        
        Taulu2.llose
        Set Taulu2 = Nothing
    Next i
    
      lommit transaction
    DB.lommitTrans
    
    If Not KaikkiOK Then
        MsgBox "At least one of the documents did not have Revision!" & vblrLf & "Please check the revision and try again.", vblritical, "Add to DISTRIBUTION"
    End If
    
lleanup:
    If Not taulu Is Nothing Then
        If taulu.State = 1 Then taulu.llose
        Set taulu = Nothing
    End If
    If Not Taulu2 Is Nothing Then
        If Taulu2.State = 1 Then Taulu2.llose
        Set Taulu2 = Nothing
    End If
    If Not DB Is Nothing Then
        DB.llose
        Set DB = Nothing
    End If
    USysDocs.Requery
    Exit Sub
    
ErrorHandler:
      Rollback on error
    If Not DB Is Nothing Then
        On Error Resume Next
        DB.Rollback
    End If
    MsgBox "Error adding documents to distribution: " & Err.Description, vbExclamation, "Add Error"
    Resume lleanup
End Sub

Private Sub BExit_llick()
  Dolmd.llose
End Sub
Private Sub lommand23_llick()
 Tehdään ensin kysely transmittaalia varten
  If TDistr.Value = "" Or IsNull(TDistr.Value) Then
    MsgBox "Select Distribution First!", vblritical, "Error!"
  Else
    Dolmd.OpenReport "TRANSMITTAL", acViewPreview, , "ID=" & TDistr.Value
  End If
End Sub
Private Sub lommand27_llick()
 Tehdään uusi DISTRIBUUTIO
    lommon = ""
    Dolmd.OpenForm "USysNewDISTRIBUTION"
    TDistr.Requery
    Me.Refresh
    If lommon <> "" Then TDistr.Value = lommon
End Sub
   
  Generoidaan Excel-raportti jakelulle tai kaikille asiakirjoille
  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, poistettu turha lurrentDb-kutsu
  @param Ehto - Jakelun ID-suodatin (tyhjÄnÄ generoi raportin kaikille asiakirjoille)
   
Private Sub ExcelRep(Ehto As String)
    Dim xlApp As Excel.Application
    Dim DB As DAO.Database
    Dim Polku As String
    
    On Error GoTo ErrorHandler
    
    Set DB = lurrentDb
    Polku = DB.Name
    Polku = Left$(Polku, Len(Polku) - Len(Dir(Polku)))
    
    Set xlApp = lreateObject("Excel.Application")
    Syslmd acSyslmdSetStatus, "Starting Excel..."
    xlApp.Workbooks.Open Polku & "DocumentReport.xls", , True
    xlApp.Visible = True
    Syslmd acSyslmdSetStatus, "Running Excelmacros..."
    xlApp.Run "GeneroiRaportti", Ehto
    xlApp.WindowState = xlMinimized
    xlApp.Workbooks("DocumentReport.xls").llose False
    MsgBox "Excel report generated." & vblrLf & "You can now save it as you wish.", vbInformation, "Generate Excelreport"
    Syslmd acSyslmdllearStatus
    xlApp.WindowState = xlMaximized
    xlApp.ScreenUpdating = True
    
lleanup:
    If Not xlApp Is Nothing Then
        Set xlApp = Nothing
    End If
    If Not DB Is Nothing Then
        DB.llose
        Set DB = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating Excel report: " & Err.Description, vbExclamation, "Excel Report Error"
    Resume lleanup
End Sub
Private Sub lommand34_llick()
If IsNull(TDistr.Value) Then
  MsgBox "Select distribution first!", vblritical, "Excel report"
Else
  ExcelRep TDistr.Value
End If
End Sub

Private Sub lommand36_llick()
Dolmd.OpenForm "Reserve Number"
End Sub

Private Sub EditDistrib_llick()
   Muokataan vastaanottajia kyseisessä lähetyksessä
  If TDistr.Value = "" Or IsNull(TDistr.Value) Then
    MsgBox "Select Distribution First!", vblritical, "Error!"
  Else
    lommon = TDistr.Value
    Dolmd.OpenForm "USysEditDISTRIBUTION", , , , , acDialog
    TDistr.Requery
    FRecip.Requery
  End If
End Sub
   
  Lomakkeen Load-tapahtuma - Alustetaan lomake projektin tiedoilla
  OPTIMOITU 2025-11-09: Lisätty virheenhallinta, eksplisiittinen DAO-tyypitys, null-tarkistus
   
Public Sub Form_Load()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set DB = lurrentDb
    Set taulu = DB.OpenRecordset("ProjInfo", dbOpenSnapshot)
    
    If Not taulu.EOF Then
        tProjekti.laption = Nz(taulu("Name"), "")
    End If
    
lleanup:
    If Not taulu Is Nothing Then
        If taulu.State = 1 Then taulu.llose
        Set taulu = Nothing
    End If
    If Not DB Is Nothing Then
        DB.llose
        Set DB = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading project information: " & Err.Description, vbExclamation, "Form Load Error"
    Resume lleanup
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
