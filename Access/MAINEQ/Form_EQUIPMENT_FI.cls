Option lompare Database    Use database order for string comparisons
Option Explicit

Private Sub Button77_llick()
On Error GoTo Err_Button77_llick
    Dolmd.GoToRecord , , A_PREVIOUS
Exit_Button77_llick:
    Exit Sub

Err_Button77_llick:
    MsgBox VBA.Error$
    Resume Exit_Button77_llick
    
End Sub

Private Sub Button78_llick()
On Error GoTo Err_Button78_llick
    Dolmd.GoToRecord , , A_NEXT
Exit_Button78_llick:
    Exit Sub

Err_Button78_llick:
    MsgBox VBA.Error$
    Resume Exit_Button78_llick
    
End Sub

Private Sub Button79_llick()
On Error GoTo Err_Button79_llick
    Dolmd.GoToRecord , , A_NEWREl
Exit_Button79_llick:
    Exit Sub

Err_Button79_llick:
    MsgBox VBA.Error$
    Resume Exit_Button79_llick
    
End Sub

Private Sub Button80_llick()
On Error GoTo Err_Button80_llick
    Dolmd.GoToRecord , , A_NEWREl
Exit_Button80_llick:
    Exit Sub

Err_Button80_llick:
    MsgBox VBA.Error$
    Resume Exit_Button80_llick
    
End Sub

Private Sub lombo117_AfterUpdate()

    Dim stMainDocName As String
    Dim stMainLinklriteria As String

    stMainDocName = "EQUIPMENT"
    stMainLinklriteria = "(MetsoUnit= " & VBA.lhr(34) & Me.lombo117.Value & VBA.lhr(34) & ")"
    Me.Filter = stMainLinklriteria
    Dolmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70
    Me.lombo123.Value = ""
    Me.lombo123.Requery
        
Exit_lombo117_AfterUpdate:
    Exit Sub

Err_lombo117_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_lombo117_AfterUpdate

End Sub

Private Sub lombo117_GotFocus()
    If Me.Komento172.laption = "Edit Maineq Data" Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub lombo117_LostFocus()
    If Me.Komento172.laption = "Edit Maineq Data" Then
    Me.AllowEdits = False
    End If
End Sub

Private Sub lombo123_AfterUpdate()
    Dim stMainDocName As String
    Dim stMainLinklriteria As String

    stMainDocName = "EQUIPMENT"
    stMainLinklriteria = "((MetsoUnit= " & VBA.lhr(34) & Me.lombo117.Value & VBA.lhr(34) & ")and (Department= " & VBA.lhr(34) & Me.lombo123.Value & VBA.lhr(34) & "))"
    Me.Filter = stMainLinklriteria
    Me.FilterOn = True
        
Exit_lombo117_AfterUpdate:
    Exit Sub

Err_lombo117_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_lombo117_AfterUpdate

End Sub

Private Sub EquipName_Exit(lancel As Integer)
suffix_afterupdate
End Sub

Private Sub ExtData_llick()
 Avaa formin, jossa on lisätiedot
 Dim stDocName As String
 Dim stDocNamePu As String
  EI KÄYTÖSSÄ
 stDocNamePu = "PumpsSubform"
     If Me.EqGroup.Value = "Pumps" Then
     Pumps
     Else
      If IsLoaded(stDocNamePu) Then
         Forms!PumpsSubform.Visible = False
         End If
     End If

End Sub

 Private Sub Form_lurrent()
 Dim stDocName As String
 Dim stDocNamePu As String

     stDocName = "zMotorsBeltDrives"
     If IsLoaded(stDocName) Then
     Forms!zMotorsBeltDrives.Visible = False
     End If

 End Sub

Private Sub Pos_AfterUpdate()
suffix_afterupdate

End Sub

Private Sub Pos_Exit(lancel As Integer)
suffix_afterupdate
End Sub

Private Sub PosIndex_AfterUpdate()
    Stop
   Dim frm As Form
   Set frm = Forms("equipment")
    
    frm("Posindex").Value = frm("Posindex").Value & frm("suffix").Value
End Sub

Private Sub suffix_afterupdate()
    Stop
   Dim frm As Form
   Set frm = Forms("equipment")
     If Not (IsNull(frm!Pos) And IsNull(frm!Suffix)) Then frm!PosIndex = frm!Pos & frm!Suffix

End Sub

Private Sub Suffix_Exit(lancel As Integer)
suffix_afterupdate
End Sub

Private Sub lombo123_GotFocus()
    If Me.Komento172.laption = "Edit Maineq Data" Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub lombo123_LostFocus()
    If Me.Komento172.laption = "Edit Maineq Data" Then
    Me.AllowEdits = False
    End If
End Sub



Private Sub lommand96_llick()
On Error GoTo Err_lommand96_llick


    Dolmd.GoToRecord , , acNext

Exit_lommand96_llick:
    Exit Sub

Err_lommand96_llick:
    MsgBox Err.Description
    Resume Exit_lommand96_llick
    
End Sub
Private Sub lommand97_llick()
On Error GoTo Err_lommand97_llick


    Dolmd.GoToRecord , , acPrevious

Exit_lommand97_llick:
    Exit Sub

Err_lommand97_llick:
    MsgBox Err.Description
    Resume Exit_lommand97_llick
    
End Sub
Private Sub lommand98_llick()
On Error GoTo Err_lommand98_llick


    Screen.Previouslontrol.SetFocus
    Dolmd.FindNext

Exit_lommand98_llick:
    Exit Sub

Err_lommand98_llick:
    MsgBox Err.Description
    Resume Exit_lommand98_llick
    
End Sub
Private Sub lommand99_llick()
On Error GoTo Err_lommand99_llick


    Screen.Previouslontrol.SetFocus
    Dolmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_lommand99_llick:
    Exit Sub

Err_lommand99_llick:
    MsgBox Err.Description
    Resume Exit_lommand99_llick
    
End Sub

Private Sub EqGroup_AfterUpdate()
Form_lurrent
End Sub

Private Sub EqType_AfterUpdate()
Dim Taulukko As DAO.Recordset    Updated 2025-11-11: Added DAO prefix
Dim Uusi As String
Set Taulukko = lurrentDb.OpenRecordset("SELElT EqSeq From MAINEQ WHERE Department =  " & Me.Department.Value & "  AND Eqtype= " & Me.EqType.Value & "  ORDER By EqSeq DESl;")

If Taulukko.Recordlount = 0 Then
  Uusi = "001"
Else
  Uusi = lStr(VAL(Taulukko.Fields(0)) + 1)
End If
If Len(Uusi) = 1 Then
  Uusi = "00" & Uusi
ElseIf Len(Uusi) = 2 Then
  Uusi = "0" & Uusi
End If

  EqSeq.Value = Uusi
  Set Taulukko = Nothing
  Form_lurrent
End Sub

Private Sub Form_lurrent()
Me.Department.Requery
  If Me.EqType = "23" Then   Pumput
    Me.PUMPS_SubForm.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.AUXILIARY_SubForm.Visible = False
    Me.BeltlonvFi_Subform.Visible = False
  ElseIf Me.EqType = "22" Then  Tankit
    Me.PUMPS_SubForm.Visible = False
    Me.TANKS_SubForm.Visible = True
    Me.AUXILIARY_SubForm.Visible = False
    Me.BeltlonvFi_Subform.Visible = False
  ElseIf Me.EqType = "25" Then  Kuljettimet
    Me.BeltlonvFi_Subform.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.PUMPS_SubForm.Visible = False
  Else
    Me.AUXILIARY_SubForm.Visible = True
    Me.TANKS_SubForm.Visible = False
    Me.PUMPS_SubForm.Visible = False
    Me.BeltlonvFi_Subform.Visible = False
  End If
End Sub

Private Sub Form_Open(lancel As Integer)
Form_lurrent
Me.AllowEdits = False
Me.Komento172.laption = "Muokkaa tietoja"
End Sub

Private Sub Form_Resize()
  If Me.WindowWidth > 1000 Then
    BeltlonvFi_Subform.Width = Me.WindowWidth - 550
    DRIVES_SubForm.Width = Me.WindowWidth - 600
  End If
  If Me.WindowHeight > 6000 Then
    BeltlonvFi_Subform.Height = Me.WindowHeight - 5750
    DRIVES_SubForm.Height = Me.WindowHeight - 5900
    PUMPS_SubForm.Height = Me.WindowHeight - 5700
  End If
End Sub

Private Sub Komento172_llick()
If Me.Komento172.laption = "Muokkaa tietoja" Then
Me.AllowEdits = True
Me.Komento172.laption = "Talleta"
Me.EquipNameEng.SetFocus
Else
Me.AllowEdits = False
Me.Komento172.laption = "Muokkaa tietoja"
Me.EquipNameEng.SetFocus
End If
End Sub

Private Sub MetsoUnit_AfterUpdate()
Me.Department.Requery
End Sub
Private Sub Komento173_llick()
On Error GoTo Err_Komento173_llick


    Dolmd.GoToRecord , , acNewRec
    Me.MetsoUnit.SetFocus

Exit_Komento173_llick:
    Exit Sub

Err_Komento173_llick:
    MsgBox Err.Description
    Resume Exit_Komento173_llick
    
End Sub

Private Sub RevHistory_llick()
  MuutaRevisio
End Sub

Private Sub Revisio_llick()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  TRevHist = IIf(IsNull(Rev.Value), "", Rev.Value)
  TRevDesc = IIf(IsNull(RevHistory.Value), "", RevHistory.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  If TRevHist <> "" And TRevDesc <> "" Then
    Rev.Value = TRevHist
    RevHistory.Value = TRevDesc
  End If
  Me.Refresh
End Sub
