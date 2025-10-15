VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysRevText_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Uusi As Boolean
Dim Eka As Boolean
Private Sub AddNew_Click()
Dim Viimeisin As String
    If Len(Rev.Value) <> 0 Then
      Viimeisin = Left(Rev.Value, 1)
      If Viimeisin = "0" Then
        Rev.Value = "A"
      Else
        Rev.Value = Chr(Asc(Viimeisin) + 1)
      End If
        TOtsikko.Caption = "New Revision"
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
  
  If Common = "ENSIMMAINEN" Then
    AddNew.Visible = False
  End If
    If Len(Revisionumero) <> 0 Then
      Rev.Value = Left(Revisionumero, InStr(Revisionumero, " ") - 1)
      Revisio = Mid(Revisionumero, InStr(Revisionumero, " ") + 1)
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
            If InStr(Revisio, vbCrLf) Then
              Approved.Value = Mid(Revisio, 1, InStr(Revisio, vbCrLf) - 1)
            Else
              Approved.Value = Revisio
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
      If Len(Revisioteksti) <> 0 Then
        Osoitin = InStr(Revisioteksti, vbCrLf)
        If Osoitin = 0 Then
          Description.Value = Mid(Revisioteksti, InStr(Revisioteksti, " ") + 1)
        Else
          Description.Value = Left(Revisioteksti, Osoitin - 1)
          Description.Value = Mid(Description.Value, InStr(Description.Value, " ") + 1)
          
        End If
      End If
    Else
      Rev.Value = "0"
      TOtsikko.Caption = "First Revision"
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
  If Eka Then 'Eka revisio, joten ei laiteta Revinvaithoa loppuun
    Revisionumero = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value
    Revisioteksti = Rev.Value & " " & Description.Value
  Else
    Revisionumero = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & vbCrLf & Revisionumero
    Revisioteksti = Rev.Value & " " & Description.Value & vbCrLf & Revisioteksti
  End If
Else 'Vanhan revision muutos
  Osoitin = InStr(Revisionumero, vbCrLf)
  If Osoitin <> 0 Then
    Revisionumero = Mid(Revisionumero, Osoitin + 2)
    Revisionumero = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & vbCrLf & Revisionumero
  Else
    Revisionumero = Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value
  End If
  Osoitin = InStr(Revisioteksti, vbCrLf)
  If Osoitin <> 0 Then
    Revisioteksti = Mid(Revisioteksti, Osoitin + 2)
    Revisioteksti = Rev.Value & " " & Description.Value & vbCrLf & Revisioteksti
  Else
    Revisioteksti = Rev.Value & " " & Description.Value
  End If
End If
  DoCmd.Close
End Sub
