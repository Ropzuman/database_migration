Option Compare Database
Option Explicit

' Painikkeen teksti, kun lomake on vain luku -tilassa
Private Const MUOKKAA_TEKSTI As String = "Edit Maineq Data"

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

' Suodattaa lomakkeen valitun yksikön mukaan ja tyhjentää osaston valinnan
Private Sub Combo117_AfterUpdate()
On Error GoTo Err_Combo117_AfterUpdate

    Me.Filter = "(MetsoUnit= " & VBA.Chr(34) & Replace(CStr(Nz(Me.Combo117.Value, "")), VBA.Chr(34), VBA.Chr(34) & VBA.Chr(34)) & VBA.Chr(34) & ")"
    Me.FilterOn = True
    Me.Combo123.Value = ""
    Me.Combo123.Requery

Exit_Combo117_AfterUpdate:
    Exit Sub

Err_Combo117_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_Combo117_AfterUpdate

End Sub

Private Sub Combo117_GotFocus()
    If Me.Komento172.Caption = MUOKKAA_TEKSTI Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub Combo117_LostFocus()
    If Me.Komento172.Caption = MUOKKAA_TEKSTI Then
    Me.AllowEdits = False
    End If
End Sub

' Suodattaa lomakkeen valitun yksikön ja osaston mukaan
Private Sub Combo123_AfterUpdate()
On Error GoTo Err_Combo123_AfterUpdate
    Dim Lainaus As String  ' Lainausmerkki
    Lainaus = VBA.Chr(34)

    Me.Filter = "((MetsoUnit= " & Lainaus & Replace(CStr(Nz(Me.Combo117.Value, "")), Lainaus, Lainaus & Lainaus) & Lainaus & ") and (Department= " & Lainaus & Replace(CStr(Nz(Me.Combo123.Value, "")), Lainaus, Lainaus & Lainaus) & Lainaus & "))"
    Me.FilterOn = True

Exit_Combo123_AfterUpdate:
    Exit Sub

Err_Combo123_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_Combo123_AfterUpdate

End Sub

Private Sub EquipName_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

' Ei käytössä: lisätietolomakkeen avaus on poistettu
Private Sub ExtData_Click()
End Sub

Private Sub Pos_AfterUpdate()
suffix_afterupdate

End Sub

Private Sub Pos_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

Private Sub PosIndex_AfterUpdate()
    ' Liitetään tunnuksen perään liite (suffix)
    Me!PosIndex.Value = Me!PosIndex.Value & Me!Suffix.Value
End Sub

' Ei toimintaa: PosIndex muodostetaan PosIndex_AfterUpdate-tapahtumassa
Private Sub suffix_afterupdate()
End Sub

Private Sub Suffix_Exit(Cancel As Integer)
suffix_afterupdate
End Sub

Private Sub Combo123_GotFocus()
    If Me.Komento172.Caption = MUOKKAA_TEKSTI Then
    Me.AllowEdits = True
    End If
End Sub

Private Sub Combo123_LostFocus()
    If Me.Komento172.Caption = MUOKKAA_TEKSTI Then
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
    DoCmd.RunCommand acCmdFind

Exit_Command99_Click:
    Exit Sub

Err_Command99_Click:
    MsgBox Err.Description
    Resume Exit_Command99_Click
    
End Sub

Private Sub EqGroup_AfterUpdate()
End Sub

' Ei toimintaa: alalomakkeiden näkyvyyttä ei ohjata tässä versiossa
Private Sub Form_Current()
End Sub

' Ehdottaa uudelle laitteelle seuraavaa juoksevaa numeroa (EqSeq) osastolla ja tyypillä
Private Sub EqType_AfterUpdate()
On Error GoTo Err_EqType_AfterUpdate
Dim Taulukko As DAO.Recordset  ' Suurin käytössä oleva numero
Dim Uusi As Long  ' Seuraava numero

  ' Suurin numero haetaan numeerisesti, jotta yli 999 -arvot ja tekstijärjestys eivät sotke tulosta
  Set Taulukko = CurrentDb.OpenRecordset("SELECT Max(Val(EqSeq)) AS Suurin FROM MAINEQ WHERE Department = '" & Replace(CStr(Nz(Me.Department.Value, "")), "'", "''") & "' AND Eqtype='" & Replace(CStr(Nz(Me.EqType.Value, "")), "'", "''") & "';", dbOpenSnapshot)
  Uusi = CLng(Nz(Taulukko.Fields(0).Value, 0)) + 1
  Taulukko.Close
  Set Taulukko = Nothing

  EqSeq.Value = Format$(Uusi, "000")
  Form_Current

Exit_EqType_AfterUpdate:
  Exit Sub

Err_EqType_AfterUpdate:
  MsgBox "Seuraavan numeron haku epäonnistui: " & Err.Description, vbExclamation
  On Error Resume Next
  If Not Taulukko Is Nothing Then Taulukko.Close
  Resume Exit_EqType_AfterUpdate
End Sub

Private Sub Form_Open(Cancel As Integer)
Form_Current
Me.AllowEdits = False
Me.Komento172.Caption = MUOKKAA_TEKSTI
End Sub

' Mukautetaan alalomakkeiden koko ikkunan kokoon
Private Sub Form_Resize()
  On Error Resume Next  ' Koonmuutos ei saa kaatua pienennettyyn ikkunaan
  If Me.WindowWidth > 1000 Then
    DRIVES_SubForm.Width = Me.WindowWidth - 300
    TANKS_SubForm.Width = Me.WindowWidth - 300
  End If
  If Me.WindowHeight > 6000 Then
    DRIVES_SubForm.Height = Me.WindowHeight - 5700
    PUMPS_SubForm.Height = Me.WindowHeight - 5700
    TANKS_SubForm.Height = Me.WindowHeight - 5700
  End If
End Sub

Private Sub Komento172_Click()
If Me.Komento172.Caption = MUOKKAA_TEKSTI Then
Me.AllowEdits = True
Me.Komento172.Caption = "Save"
Me.EquipNameEng.SetFocus
Else
Me.AllowEdits = False
Me.Komento172.Caption = MUOKKAA_TEKSTI
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
