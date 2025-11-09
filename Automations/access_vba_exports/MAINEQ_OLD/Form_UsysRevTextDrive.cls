Option Compare Database
Option Explicit
Dim Uusi As Boolean
Dim Eka As Boolean
Private Sub AddNew_Click()
Dim Viimeisin As String
    If Len(Rev.Value) <> 0 Then
      Viimeisin = Rev.Value
      If Left(Viimeisin, 1) = "0" Then
        Viimeisin = Mid(Viimeisin, 2, 1)
      End If
      Viimeisin = CStr(VAL(Viimeisin) + 1)
      If Len(Viimeisin) = 1 Then Viimeisin = "0" & Viimeisin
      Rev.Value = Viimeisin
      TOtsikko.Caption = "New Revision"
      TOtsikko2.Caption = "Uusi revisio"
      DateFld.Value = Date
      Uusi = True
      Talleta.SetFocus
      AddNew.Enabled = False
    Else
      MsgBox "Revision mark field was empty! Unable to determine next revision.", vbCritical, "Error!"
    End If
End Sub
Private Sub Command16_Click()
  DoCmd.Close
End Sub
Private Sub Form_Load()
Dim Revisio As String
Dim Osoitin As Long
Uusi = False
Eka = False
  
'  TRevHist = "a 12.01.2000/aa/bb/cc" & vbCrLf & "0 1.1.2000/aa/bb/ee"
'  TRevDesc = "a Updated" & vbCrLf & "0 Preliminary"
'  TRevHist = ""
'  TRevDesc = ""
  
    If Len(TRevHist) <> 0 Then
      Rev.Value = LCase(Left(TRevHist, InStr(TRevHist, " ") - 1))
      Revisio = Mid(TRevHist, 3)
      Osoitin = InStr(Revisio, "/")
      If Osoitin <> 0 Then
        DateFld.Value = Mid(Revisio, 1, Osoitin - 1)
        Revisio = Mid(Revisio, Osoitin + 1)
        Osoitin = InStr(Revisio, "/")
        If Osoitin <> 0 Then
          Drawn.Value = Mid(Revisio, 1, Osoitin - 1)
          Revisio = Mid(Revisio, Osoitin + 1)
          Osoitin = InStr(Revisio, "/")
          If Osoitin <> 0 Then
            Checked.Value = Mid(Revisio, 1, Osoitin - 1)
            Revisio = Mid(Revisio, Osoitin + 1)
            Osoitin = InStr(Revisio, "/")
            If Osoitin <> 0 Then
              Approved.Value = Mid(Revisio, 1, Osoitin - 1)
              Revisio = Mid(Revisio, Osoitin + 1)
              If InStr(Revisio, vbCrLf) Then
                Description.Value = Mid(Revisio, 1, InStr(Revisio, vbCrLf) - 1)
              Else
                Description.Value = Revisio
              End If
            End If
          Else
            Checked.Value = Revisio
          End If
        Else
          Drawn.Value = Revisio
        End If
      Else
        DateFld.Value = Revisio
      End If
    Else
      Rev.Value = "0"
      TOtsikko.Caption = "First Revision"
      TOtsikko2.Caption = "Ensimm�inen revisio"
      AddNew.Enabled = False
      Uusi = True
      Eka = True
    End If
End Sub
Private Sub Talleta_Click()
Dim Osoitin As Long

If IsNull(Rev.Value) Or Rev.Value = "" Or IsNull(DateFld.Value) Or DateFld.Value = "" _
    Or IsNull(Drawn.Value) Or Drawn.Value = "" Then
  MsgBox "Revision Mark, Date and Author fields can't be left empty!", vbCritical, "Warning!"
  Exit Sub
ElseIf IsNull(Checked.Value) Or IsNull(Approved.Value) Or IsNull(Description.Value) Then
  If MsgBox("Checker, Approver or Description fields are empty" & vbCrLf _
          & "Are your sure that you want to keep it that way?", vbYesNo, "Warning!") = vbNo Then
    Exit Sub
  End If
End If
  
If Uusi Then
  If Eka Then 'Eka revisio, joten ei laiteta Rivinvaihtoa loppuun
    TRevHist = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value
  Else
    TRevHist = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value & vbCrLf & TRevHist
  End If
Else 'Vanhan revision muutos
  Osoitin = InStr(TRevHist, vbCrLf)
  If Osoitin <> 0 Then
    TRevHist = Mid(TRevHist, Osoitin + 2)
    TRevHist = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value & vbCrLf & TRevHist
  Else
    TRevHist = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value
  End If
End If
  DoCmd.Close
End Sub
