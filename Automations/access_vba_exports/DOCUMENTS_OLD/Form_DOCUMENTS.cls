Option Compare Database
Option Explicit
Private Sub Add_Distr_Click()
  CurRecord = Counter.Value
  If IsNull(Rev.Value) Then
    MsgBox "You have to put revision for document first!", vbCritical, "Error!"
  Else
    Common = ""
    DefRev = ""
    DocStatus = ""
    Common = HaeRevisio(Rev.Value)
    DocStatus = Nz(Status.Value)
    DoCmd.OpenForm "USysAddToDistr", , , , , acDialog
    Lahetykset.Requery
  End If
End Sub
Private Sub AvaaTiedosto_Click()
'T�m� avaa m��ritellyn tiedosto
Dim FSO As New FileSystemObject
Dim Avattava As String
If IsNull(File.Value) Then
  MsgBox "No filename defined", vbCritical, "Open File"
ElseIf IsNull(WorkPath.Value) Then
  MsgBox "No path defined", vbCritical, "Open File"
Else
  If InStr(File.Value, vbCrLf) Then
    Common = Counter.Value
    DoCmd.OpenForm "USysOpenFile", , , , , acDialog
  Else
    Avattava = FSO.BuildPath(WorkPath.Value, File.Value)
    If FSO.FileExists(Avattava) Then
         Application.FollowHyperlink Avattava, , True, False
    Else 'Tiedostoa ei l�ytynyt
      MsgBox "File not found: " & WorkPath.Value & File.Value, vbCritical, "Open File"
    End If
  End If
End If
Set FSO = Nothing
End Sub
Private Sub ClientDwg_Click()
  If IsNull(ClientNo.Value) Then 'Numeroa ei viel� varattu, joten voidaan suorittaa varaus
    DoCmd.OpenForm "USysReserve", , , "Counter = " & Counterfield.Value, , acDialog
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
    Counterfield.SetFocus
    DoCmd.FindRecord CurRecord, , , , , acCurrent, True
    LisaaUusi.SetFocus
  End If
End Sub
Private Sub CPath_Click()
Dim Hakemisto As String
  Hakemisto = ValitseHakem(Me.hWnd, IIf(IsNull(WorkPath.Value), "", WorkPath.Value))
  If Hakemisto <> "" Then
    WorkPath.Value = Hakemisto
    Me.Refresh
  End If
End Sub
Private Sub DetDwg_Click()
  Common = IIf(IsNull(DetDwg.Value), "", DetDwg.Value)
  CommonOts = "Detail drawing(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  DetDwg.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub File_Click()
  Common = Nz(File.Value)
  CommonOts = "Files(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  File.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub Form_Load()
Dim DB As Database
Dim taulu As Recordset
Set DB = CurrentDb
Set taulu = CurrentDb.OpenRecordset("ProjInfo")
tProjekti.Caption = taulu("Name")
taulu.Close
DB.Close
Set taulu = Nothing
Set DB = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Me.Filter = ""
  Me.FilterOn = False
End Sub
Private Sub Note_Click()
  Common = IIf(IsNull(Note.Value), "", Note.Value)
  CommonOts = "Note:"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Note.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub OpenPath_Click()
Dim FSO As New FileSystemObject
If IsNull(WorkPath.Value) Then
  MsgBox "No path defined", vbCritical, "Open Path"
Else
  If FSO.FolderExists(WorkPath.Value) Then
    Application.FollowHyperlink WorkPath.Value
  Else 'Hakemistoa ei ole
    MsgBox "Invalid path", vbCritical, "Open Path"
  End If
End If
Set FSO = Nothing
End Sub
Private Sub Pos_Click()
  Common = IIf(IsNull(Pos.Value), "", Pos.Value)
  CommonOts = "Position(s):"
  CommonType = 1
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
  Pos.Value = IIf(Common = "", Null, Common)
End Sub
Private Sub Rev_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  Revisioteksti = Nz(Rev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
