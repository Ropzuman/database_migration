Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_DOCUMENTS
' SOVELLUS : DOCUMENTS — Dokumenttirekisterin päälomake
' KUVAUS   : Päälomake dokumenttien hallintaan: uusien dokumenttien lisäys,
'            revisioiden muokkaus, tiedostojen avaus, distribuutioiden
'            hallinta sekä asiakkaan piirustunumeron varaus.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub Add_Distr_Click()
  CurRecord = Me.Counter.Value
  If IsNull(Me.Rev.Value) Then
    MsgBox "You have to put revision for document first!", vbCritical, "Error!"
  Else
    Common = ""
    DefRev = ""
    DocStatus = ""
    Common = HaeRevisio(Me.Rev.Value)
    DocStatus = Nz(Me.Status.Value)
    DoCmd.OpenForm "USysAddToDistr", , , , , acDialog
    Me.Lahetykset.Requery
  End If
End Sub
Private Sub AvaaTiedosto_Click()
' Avaa dokumenttiin liitetyn tiedoston tai näyttää tiedostolistan dialogi
Dim fso As New FileSystemObject
Dim Avattava As String
If IsNull(Me.File.Value) Then
  MsgBox "No filename defined", vbCritical, "Open File"
ElseIf IsNull(Me.WorkPath.Value) Then
  MsgBox "No path defined", vbCritical, "Open File"
Else
  If InStr(Me.File.Value, vbCrLf) Then
    Common = Me.Counter.Value
    DoCmd.OpenForm "USysOpenFile", , , , , acDialog
  Else
    Avattava = fso.BuildPath(Me.WorkPath.Value, Me.File.Value)
    If fso.FileExists(Avattava) Then
         Application.FollowHyperlink Avattava, , True, False
    Else
      MsgBox "File not found: " & Me.WorkPath.Value & Me.File.Value, vbCritical, "Open File"
    End If
  End If
End If
Set fso = Nothing
End Sub
Private Sub ClientDwg_Click()
  ' Varaa asiakkaan piirustunumero, jos ei vielä varattuna
  If IsNull(Me.ClientNo.Value) Then
    DoCmd.OpenForm "USysReserve", , , "Counter = " & Me.Counterfield.Value, , acDialog
    Me.Refresh
  Else
    MsgBox "You can't change reserved client number.", , "Client Document number"
  End If
End Sub
Private Sub Command58_Click()
  DoCmd.Close
End Sub
Private Sub LisaaUusi_Click()
  CurRecord = ""
  Me.Refresh
  Me.Filter = ""
  Me.FilterOn = False
  DoCmd.OpenForm "USysAddDocument", , , , , acDialog
  If CurRecord = "" Then
  Else
    Me.Requery
    Me.Refresh
    Me.Counterfield.SetFocus
    DoCmd.FindRecord CurRecord, , , , , acCurrent, True
    Me.LisaaUusi.SetFocus
  End If
End Sub
Private Sub CPath_Click()
' Avataan hakemiston valinta -dialogi käyttöpolun asettamiseksi
Dim Hakemisto As String
  Hakemisto = ValitseHakem(Me.hWnd, IIf(IsNull(Me.WorkPath.Value), "", Me.WorkPath.Value))
  If Hakemisto <> "" Then
    Me.WorkPath.Value = Hakemisto
    Me.Refresh
  End If
End Sub
Private Sub DetDwg_Click()
  Common = IIf(IsNull(Me.DetDwg.Value), "", Me.DetDwg.Value)
  CommonOts = "Detail drawing(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.DetDwg.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub File_Click()
  Common = Nz(Me.File.Value)
  CommonOts = "Files(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.File.Value = IIf(Common = "", Null, Common)
End Sub
'''
' Form Load Event - Initialize form with project information
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, removed redundant CurrentDb call
'''
Private Sub Form_Load()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set DB = CurrentDb
    Set taulu = DB.OpenRecordset("ProjInfo")
    
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
Private Sub Form_Unload(Cancel As Integer)
  Me.Filter = ""
  Me.FilterOn = False
End Sub
Private Sub Note_Click()
  Common = IIf(IsNull(Me.Note.Value), "", Me.Note.Value)
  CommonOts = "Note:"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.Note.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub OpenPath_Click()
' Avaa työhakemiston Windows Explorerissa
Dim fso As New FileSystemObject
If IsNull(Me.WorkPath.Value) Then
  MsgBox "No path defined", vbCritical, "Open Path"
Else
  If fso.FolderExists(Me.WorkPath.Value) Then
    Application.FollowHyperlink Me.WorkPath.Value
  Else
    MsgBox "Invalid path", vbCritical, "Open Path"
  End If
End If
Set fso = Nothing
End Sub
Private Sub Pos_Click()
  Common = IIf(IsNull(Me.Pos.Value), "", Me.Pos.Value)
  CommonOts = "Position(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.Pos.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub Rev_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  ' Avataan revisioeditori ja tallennetaan tulos Rev-kenttään
  Revisioteksti = Nz(Me.Rev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Me.Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub