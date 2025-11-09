Option Compare Database

Private Sub Command3_Click()
  AvaaValittu False
End Sub

Private Sub Command4_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
Dim SQL As String
Dim Alue As String
Dim Linja As String
Dim Parametrit As Variant
  If IsNull(Me.OpenArgs) Then
     MsgBox "Avaa t�m� lomake putkilinjataulukon kautta," & vbCrLf & "painamalla Add-Ins valikosta 'Etsi kohde'", vbCritical, "Virhe"
  Else
    Parametrit = Split(Me.OpenArgs, ",")
    If UBound(Parametrit) = 1 Then
      Alue = Trim(Parametrit(0))
      Linja = Trim(Parametrit(1))
      SQL = "SELECT PATH, ImpFileID, Handles, DEP, FLOWSUBSTMAINGROUP AS FSMG, LINENO, FLOWSUBSTCODE AS FSC, DN, PN,  [FROM], [TO] FROM PIPELINEDATA "
      SQL = SQL & "WHERE DEP = '" & Alue & "' AND LINENO = '" & Linja & "'"
      Lista.RowSource = SQL
    End If
  End If
  If Lista.ListCount = 2 Then
    Lista.Selected(1) = True
    AvaaValittu True
  ElseIf Lista.ListCount < 2 Then
    MsgBox "Putkilinjatietoja ei ole tuotu virtauskaavioista", vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
  Else
    Lista.Selected(1) = True
  End If
End Sub
Private Sub AvaaValittu(Sulje As Boolean)
Dim Polku As String
Dim DWG As String
Dim Handle As String
Dim Info As String
    Polku = Lista.Column(0, Lista.ListIndex + 1)
    DWG = Lista.Column(1, Lista.ListIndex + 1)
    Handle = Lista.Column(2, Lista.ListIndex + 1)
    Info = Lista.Column(3, Lista.ListIndex + 1) & "-" & Lista.Column(5, Lista.ListIndex + 1)
    AvaaKuvasta Polku, DWG, Handle, Info
    If Sulje Then
      DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
    End If
End Sub
