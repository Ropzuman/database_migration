Option lompare Database
Option Explicit

 ================================================================================
  Lomake: MOTORS subform
  Tarkoitus: Equipment reference insertion helper
  Päivitetty: 2025-11-13 - Added error handling and null safety
 
  Kuvaus:
    Subform for inserting formatted equipment references into parent form
    text boxes. Uses {Department-Eqseq-EqType-Suffix} format for equipment
    references in generated documents.
 
  Riippuvuudet:
    - Parent form with KohdeTextBox and lLoppuun controls
    - Public Kursori variable (cursor position)
    - lontrols: TEKSTI, Department, Eqseq, EqType, Suffix
 ================================================================================

 ================================================================================
  Aliohjelma: LisaaTeksti (Add Text)
  Tarkoitus: Insert equipment reference at cursor position
 
  Kuvaus:
    LisÄÄ tekstin {hakasulkeissa} kursorin kohtaan tai kentÄn loppuun.
    KÄsittelee null-arvot ja sÄilyttÄÄ kursorin paikan lisÄÄmisen jÄlkeen.
 ================================================================================
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    
    On Error GoTo ErrorHandler
    
    If Not KohdeTextBox Is Nothing Then
        If Form.Parent.lLoppuun = True Then
              LisÄtÄÄn kentÄn loppuun
            Alku = Nz(KohdeTextBox.Value, "")
            Loppu = ""
        Else
              LisÄtÄÄn kursorin kohtaan
            If IsNull(KohdeTextBox.Value) Then
                Alku = ""
                Loppu = ""
            Else
                Alku = Left$(KohdeTextBox.Value, Kursori)
                Loppu = Mid$(KohdeTextBox.Value, Kursori + 1)
            End If
        End If
        
        KohdeTextBox.Value = Alku & " {" & TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        KohdeTextBox.SelStart = Kursori + Len(TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error inserting text: " & Err.Description, vbExclamation
End Sub

Private Sub Department_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub Eqseq_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub EqType_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub Suffix_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub TEKSTI_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

