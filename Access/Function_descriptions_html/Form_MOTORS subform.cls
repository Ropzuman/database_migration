Option Compare Database
Option Explicit

'================================================================================
' Form: MOTORS subform
' Purpose: Equipment reference insertion helper
' Updated: 2025-11-13 - Added error handling and null safety
'
' Description:
'   Subform for inserting formatted equipment references into parent form
'   text boxes. Uses {Department-Eqseq-EqType-Suffix} format for equipment
'   references in generated documents.
'
' Dependencies:
'   - Parent form with KohdeTextBox and CLoppuun controls
'   - Public Kursori variable (cursor position)
'   - Controls: TEKSTI, Department, Eqseq, EqType, Suffix
'================================================================================

'================================================================================
' Sub: LisaaTeksti (Add Text)
' Purpose: Insert equipment reference at cursor position
'
' Description:
'   Inserts text in {brackets} at current cursor position or end of field.
'   Handles null values and maintains cursor position after insertion.
'================================================================================
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    
    On Error GoTo ErrorHandler
    
    If Not KohdeTextBox Is Nothing Then
        If Form.Parent.CLoppuun = True Then
            ' Append to end
            Alku = Nz(KohdeTextBox.Value, "")
            Loppu = ""
        Else
            ' Insert at cursor position
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

Private Sub Department_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub Eqseq_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub EqType_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub Suffix_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub TEKSTI_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

