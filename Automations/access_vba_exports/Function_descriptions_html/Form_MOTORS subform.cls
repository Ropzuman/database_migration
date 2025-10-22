VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MOTORS subform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
