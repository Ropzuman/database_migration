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
Private Const TV_DOC_CURRENT_COUNTER As String = "DOC_CurrentCounter"
Private Const TV_DOC_DIST_REV As String = "DOC_AddToDistr_Rev"
Private Const TV_DOC_DIST_STATUS As String = "DOC_AddToDistr_Status"
Private Const TV_DOC_NEW_COUNTER As String = "DOC_NewRecordCounter"
Private Const TV_DOC_COMMON_TEXT As String = "DOC_Common_Text"
Private Const TV_DOC_COMMON_TITLE As String = "DOC_Common_Title"
Private Const TV_DOC_COMMON_TYPE As String = "DOC_Common_Type"
Private Const TV_DOC_REV_TEXT As String = "DOC_Rev_Text"

Private Sub Add_Distr_Click()
  If IsNull(Me.Rev.Value) Then
    MsgBox "You have to put revision for document first!", vbCritical, "Error!"
  Else
    AsetaTempVar TV_DOC_CURRENT_COUNTER, CStr(Me.Counter.Value)
    AsetaTempVar TV_DOC_DIST_REV, HaeRevisioPaikallinen(Me.Rev.Value)
    AsetaTempVar TV_DOC_DIST_STATUS, Nz(Me.Status.Value, "")
    DoCmd.OpenForm "USysAddToDistr", , , , , acDialog
    Me.Lahetykset.Requery
  End If
End Sub
Private Sub AvaaTiedosto_Click()
' Avaa dokumenttiin liitetyn tiedoston tai näyttää tiedostolistan dialogi
Dim fso As Object
Dim Avattava As String

Set fso = CreateObject("Scripting.FileSystemObject")
If IsNull(Me.File.Value) Then
  MsgBox "No filename defined", vbCritical, "Open File"
ElseIf IsNull(Me.WorkPath.Value) Then
  MsgBox "No path defined", vbCritical, "Open File"
Else
  If InStr(Me.File.Value, vbCrLf) Then
    DoCmd.OpenForm "USysOpenFile", , , , , acDialog, CStr(Me.Counter.Value)
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
  AsetaTempVar TV_DOC_NEW_COUNTER, ""
  Me.Refresh
  Me.Filter = ""
  Me.FilterOn = False
  DoCmd.OpenForm "USysAddDocument", , , , , acDialog
  If LueTempVarTeksti(TV_DOC_NEW_COUNTER) = "" Then
  Else
    Me.Requery
    Me.Refresh
    Me.Counterfield.SetFocus
    DoCmd.FindRecord LueTempVarTeksti(TV_DOC_NEW_COUNTER), , , , , acCurrent, True
    Me.LisaaUusi.SetFocus
  End If
End Sub
Private Sub CPath_Click()
' Avataan hakemiston valinta -dialogi käyttöpolun asettamiseksi
Dim Hakemisto As String
  Hakemisto = ValitseHakemPaikallinen(IIf(IsNull(Me.WorkPath.Value), "", Me.WorkPath.Value))
  If Hakemisto <> "" Then
    Me.WorkPath.Value = Hakemisto
    Me.Refresh
  End If
End Sub
Private Sub DetDwg_Click()
  AsetaCommonPaikallinen IIf(IsNull(Me.DetDwg.Value), "", Me.DetDwg.Value), "Detail drawing(s):", 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.DetDwg.Value = IIf(HaeCommonPaikallinen() = "", Null, HaeCommonPaikallinen())
End Sub
Private Sub File_Click()
  AsetaCommonPaikallinen Nz(Me.File.Value), "Files(s):", 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.File.Value = IIf(HaeCommonPaikallinen() = "", Null, HaeCommonPaikallinen())
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
  AsetaCommonPaikallinen IIf(IsNull(Me.Note.Value), "", Me.Note.Value), "Note:", 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.Note.Value = IIf(HaeCommonPaikallinen() = "", Null, HaeCommonPaikallinen())
End Sub
Private Sub OpenPath_Click()
' Avaa työhakemiston Windows Explorerissa
Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")
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
  AsetaCommonPaikallinen IIf(IsNull(Me.Pos.Value), "", Me.Pos.Value), "Position(s):", 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Me.Pos.Value = IIf(HaeCommonPaikallinen() = "", Null, HaeCommonPaikallinen())
End Sub
Private Sub Rev_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  ' Avataan revisioeditori ja tallennetaan tulos Rev-kenttään
  AsetaTempVar TV_DOC_REV_TEXT, Nz(Me.Rev.Value, "")
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If LueTempVarTeksti(TV_DOC_REV_TEXT) <> "" Then
      Me.Rev.Value = LueTempVarTeksti(TV_DOC_REV_TEXT)
  End If
  Me.Refresh
End Sub

Private Function ValitseHakemPaikallinen(Optional ByVal AloitusPolku As String = "") As String
  ' Käytetään paikallista kansiovalitsinta, jotta käännös ei riipu ulkoisen moduulin näkyvyydestä.
  Dim Dlg As Object
  Dim Polku As String

  On Error GoTo ErrorHandler

  Set Dlg = Application.FileDialog(4)
  With Dlg
    .Title = "Choose path:"
    .AllowMultiSelect = False
    If AloitusPolku <> "" Then .InitialFileName = AloitusPolku
    If .Show <> -1 Then
      ValitseHakemPaikallinen = ""
      GoTo Cleanup
    End If
    Polku = CStr(.SelectedItems(1))
  End With

  If Polku <> "" Then
    If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
  End If
  ValitseHakemPaikallinen = Polku

Cleanup:
  Set Dlg = Nothing
  Exit Function

ErrorHandler:
  ValitseHakemPaikallinen = ""
  Resume Cleanup
End Function

Private Sub AsetaTempVar(ByVal Nimi As String, ByVal Arvo As String)
  On Error Resume Next
  TempVars.Remove Nimi
  On Error GoTo 0
  TempVars.Add Nimi, Arvo
End Sub

Private Function LueTempVarTeksti(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTeksti = Nz(TempVars(Nimi).Value, "")
  Exit Function

ErrorHandler:
  LueTempVarTeksti = ""
End Function

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

Private Sub AsetaCommonPaikallinen(ByVal Teksti As String, ByVal Otsikko As String, ByVal Tyyppi As Integer)
  AsetaTempVar TV_DOC_COMMON_TEXT, Teksti
  AsetaTempVar TV_DOC_COMMON_TITLE, Otsikko
  AsetaTempVar TV_DOC_COMMON_TYPE, CStr(Tyyppi)
End Sub

Private Function HaeCommonPaikallinen() As String
  HaeCommonPaikallinen = LueTempVarTeksti(TV_DOC_COMMON_TEXT)
End Function