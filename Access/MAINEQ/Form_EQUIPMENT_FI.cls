Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Button77_Click()
On Error GoTo Err_Button77_Click
    DoCmd.GoToRecord , , A_PREVIOUS
Exit_Button77_Click:
    Exit Sub

Err_Button77_Click:
    MsgBox VBA.Error$
    Resume Exit_Button77_Click
    
End Sub

Private Sub Button78_Click()
On Error GoTo Err_Button78_Click
    DoCmd.GoToRecord , , A_NEXT
Exit_Button78_Click:
    Exit Sub

Err_Button78_Click:
    MsgBox VBA.Error$
    Resume Exit_Button78_Click
    
End Sub

Private Sub Button79_Click()
On Error GoTo Err_Button79_Click
    DoCmd.GoToRecord , , A_NEWREC
Exit_Button79_Click:
    Exit Sub

Err_Button79_Click:
    MsgBox VBA.Error$
    Resume Exit_Button79_Click
    
End Sub

Private Sub Button80_Click()
On Error GoTo Err_Button80_Click
    DoCmd.GoToRecord , , A_NEWREC
Exit_Button80_Click:
    Exit Sub

Err_Button80_Click:
    MsgBox VBA.Error$
    Resume Exit_Button80_Click
    
End Sub

Private Sub Combo117_AfterUpdate()

    Dim stMainDocName As String
    Dim stMainLinkCriteria As String

    stMainDocName = "EQUIPMENT"
    stMainLinkCriteria = "(MetsoUnit= " & VBA.Chr(34) & Me.Combo117.Value & VBA.Chr(34) & ")"
    Me.Filter = stMainLinkCriteria
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70
    Me.Combo123.Value = ""
    Me.Combo123.Requery
        
Exit_Combo117_AfterUpdate:
    Exit Sub

Err_Combo117_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_Combo117_AfterUpdate

End Sub

Private Sub Combo117_GotFocus()
    If Me.Komento172.Caption = "Edit Maineq Data" Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub Combo117_LostFocus()
    If Me.Komento172.Caption = "Edit Maineq Data" Then
    Me.AllowEdits = False
    End If
End Sub

Private Sub Combo123_AfterUpdate()
    Dim stMainDocName As String
    Dim stMainLinkCriteria As String

    stMainDocName = "EQUIPMENT"
    stMainLinkCriteria = "((MetsoUnit= " & VBA.Chr(34) & Me.Combo117.Value & VBA.Chr(34) & ")and (Department= " & VBA.Chr(34) & Me.Combo123.Value & VBA.Chr(34) & "))"
    Me.Filter = stMainLinkCriteria
    Me.FilterOn = True
        
Exit_Combo117_AfterUpdate:
    Exit Sub

Err_Combo117_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_Combo117_AfterUpdate

End Sub

Private Sub EquipName_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

Private Sub ExtData_Click()
'Avaa formin, jossa on lisätiedot
'Dim stDocName As String
'Dim stDocNamePu As String
' EI KÄYTÖSSÄ
'stDocNamePu = "PumpsSubform"
'    If Me.EqGroup.Value = "Pumps" Then
'    Pumps
'    Else
'     If IsLoaded(stDocNamePu) Then
'        Forms!PumpsSubform.Visible = False
'        End If
'    End If

End Sub

'Private Sub Form_Current()
'Dim stDocName As String
'Dim stDocNamePu As String

'    stDocName = "zMotorsBeltDrives"
'    If IsLoaded(stDocName) Then
'    Forms!zMotorsBeltDrives.Visible = False
'    End If

'End Sub

Private Sub Pos_AfterUpdate()
suffix_afterupdate

End Sub

Private Sub Pos_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

Private Sub PosIndex_AfterUpdate()
   'Stop
   Dim frm As Form
   Set frm = Forms("equipment")
    
    frm("Posindex").Value = frm("Posindex").Value & frm("suffix").Value
End Sub

Private Sub suffix_afterupdate()
   'Stop
   Dim frm As Form
   Set frm = Forms("equipment")
    'If Not (IsNull(frm!Pos) And IsNull(frm!Suffix)) Then frm!PosIndex = frm!Pos & frm!Suffix

End Sub

Private Sub Suffix_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

Private Sub Combo123_GotFocus()
    If Me.Komento172.Caption = "Edit Maineq Data" Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub Combo123_LostFocus()
    If Me.Komento172.Caption = "Edit Maineq Data" Then
    Me.AllowEdits = False
    End If
End Sub



Private Sub Command96_Click()
On Error GoTo Err_Command96_Click


    DoCmd.GoToRecord , , acNext

Exit_Command96_Click:
    Exit Sub

Err_Command96_Click:
    MsgBox Err.Description
    Resume Exit_Command96_Click
    
End Sub
Private Sub Command97_Click()
On Error GoTo Err_Command97_Click


    DoCmd.GoToRecord , , acPrevious

Exit_Command97_Click:
    Exit Sub

Err_Command97_Click:
    MsgBox Err.Description
    Resume Exit_Command97_Click
    
End Sub
Private Sub Command98_Click()
On Error GoTo Err_Command98_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command98_Click:
    Exit Sub

Err_Command98_Click:
    MsgBox Err.Description
    Resume Exit_Command98_Click
    
End Sub
Private Sub Command99_Click()
On Error GoTo Err_Command99_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Command99_Click:
    Exit Sub

Err_Command99_Click:
    MsgBox Err.Description
    Resume Exit_Command99_Click
    
End Sub

Private Sub EqGroup_AfterUpdate()
Form_Current
End Sub

Private Sub EqType_AfterUpdate()
Dim Taulukko As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix
Dim Uusi As String
Set Taulukko = CurrentDb.OpenRecordset("SELECT EqSeq From MAINEQ WHERE Department = '" & Me.Department.Value & "' AND Eqtype='" & Me.EqType.Value & "' ORDER By EqSeq DESC;")

If Taulukko.RecordCount = 0 Then
  Uusi = "001"
Else
  Uusi = CStr(VAL(Taulukko.Fields(0)) + 1)
End If
If Len(Uusi) = 1 Then
  Uusi = "00" & Uusi
ElseIf Len(Uusi) = 2 Then
  Uusi = "0" & Uusi
End If

  EqSeq.Value = Uusi
  Set Taulukko = Nothing
  Form_Current
End Sub

Private Sub Form_Current()
Me.Department.Requery
  If Me.EqType = "23" Then ' Pumput
    Me.PUMPS_SubForm.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.AUXILIARY_SubForm.Visible = False
    Me.BeltConvFi_Subform.Visible = False
  ElseIf Me.EqType = "22" Then 'Tankit
    Me.PUMPS_SubForm.Visible = False
    Me.TANKS_SubForm.Visible = True
    Me.AUXILIARY_SubForm.Visible = False
    Me.BeltConvFi_Subform.Visible = False
  ElseIf Me.EqType = "25" Then 'Kuljettimet
    Me.BeltConvFi_Subform.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.PUMPS_SubForm.Visible = False
  Else
    Me.AUXILIARY_SubForm.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.PUMPS_SubForm.Visible = False
    Me.BeltConvFi_Subform.Visible = False
  End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form_Current
Me.AllowEdits = False
Me.Komento172.Caption = "Muokkaa tietoja"
End Sub

Private Sub Form_Resize()
  If Me.WindowWidth > 1000 Then
    BeltConvFi_Subform.Width = Me.WindowWidth - 550
    DRIVES_SubForm.Width = Me.WindowWidth - 600
  End If
  If Me.WindowHeight > 6000 Then
    BeltConvFi_Subform.Height = Me.WindowHeight - 5750
    DRIVES_SubForm.Height = Me.WindowHeight - 5900
    PUMPS_SubForm.Height = Me.WindowHeight - 5700
  End If
End Sub

Private Sub Komento172_Click()
If Me.Komento172.Caption = "Muokkaa tietoja" Then
Me.AllowEdits = True
Me.Komento172.Caption = "Talleta"
Me.EquipNameEng.SetFocus
Else
Me.AllowEdits = False
Me.Komento172.Caption = "Muokkaa tietoja"
Me.EquipNameEng.SetFocus
End If
End Sub

Private Sub MetsoUnit_AfterUpdate()
Me.Department.Requery
End Sub
Private Sub Komento173_Click()
On Error GoTo Err_Komento173_Click


    DoCmd.GoToRecord , , acNewRec
    Me.MetsoUnit.SetFocus

Exit_Komento173_Click:
    Exit Sub

Err_Komento173_Click:
    MsgBox Err.Description
    Resume Exit_Komento173_Click
    
End Sub

Private Sub RevHistory_Click()
  MuutaRevisio
End Sub

Private Sub Revisio_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  TRevHist = IIf(IsNull(Rev.Value), "", Rev.Value)
  TRevDesc = IIf(IsNull(RevHistory.Value), "", RevHistory.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If TRevHist <> "" And TRevDesc <> "" Then
    Rev.Value = TRevHist
    RevHistory.Value = TRevDesc
  End If
  Me.Refresh
End Sub
