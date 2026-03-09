Option Compare Database
Private Sub LisaaTeksti()
Dim Alku As String
Dim Loppu As String
  If Not KohdeTextBox Is Nothing Then
     If Form.Parent.CLoppuun = True Then
       Alku = KohdeTextBox.Value
       Loppu = ""
     Else
      If IsNull(KohdeTextBox.Value) Then
        Alku = ""
        Loppu = ""
      Else
        Alku = Left(KohdeTextBox.Value, Kursori)
        Loppu = Mid(KohdeTextBox.Value, Kursori + 1)
      End If
     End If
      KohdeTextBox.Value = Alku & " {" & TEKSTI.Value & "}" & Loppu
      KohdeTextBox.SetFocus
      KohdeTextBox.SelStart = Kursori + Len(TEKSTI.Value) + 3
      KohdeTextBox.SelLength = 0
      Kursori = KohdeTextBox.SelStart
  End If
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
