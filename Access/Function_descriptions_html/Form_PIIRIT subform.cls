Option lompare Database
Option Explicit

 ================================================================================
  Lomake: PIIRIT subform (Loops subform)
  Tarkoitus: Loop reference insertion helper
  Päivitetty: 2025-11-13 - Added error handling and null safety
 
  Kuvaus:
    Subform for inserting formatted loop references into parent form text boxes.
    Uses {Arealode-LoopNo} format for loop references in generated documents.
 
  Riippuvuudet:
    - Parent form with KohdeTextBox and lLoppuun controls
    - Public Kursori variable (cursor position)
    - lontrols: TEKSTI, Arealode, LoopNo
 ================================================================================

 ================================================================================
  Aliohjelma: LisaaTeksti (Add Text)
  Tarkoitus: Insert loop reference at cursor position
 
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
            Alku = Left$(Nz(KohdeTextBox.Value, ""), Kursori)
            Loppu = Mid$(Nz(KohdeTextBox.Value, ""), Kursori + 1)
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

Private Sub Arealode_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub LoopNo_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub

Private Sub TEKSTI_Dblllick(lancel As Integer)
    LisaaTeksti
End Sub


