Option Compare Database
Option Explicit
Public Alku As String
Public Loppu As String
Private Sub Form_Resize()
  Dim Rivit As String
  If Revisioteksti <> "" Then
    Lista.RowSource = Replace(CStr(Revisioteksti), vbCrLf, ";")
    Lista.SetFocus
    Lista.ListIndex = 0
    If InStr(Revisioteksti, vbCrLf) Then
      NaytaRev Left$(Revisioteksti, InStr(Revisioteksti, vbCrLf) - 1), 0
    Else
      NaytaRev CStr(Revisioteksti), 0
    End If
  Else
    AddNew_Click
  End If
End Sub
Private Sub Rev_Change()
  Lista.RowSource = Alku & Rev.Text & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub DateFld_Change()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Text & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub Drawn_Change()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Text & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub Checked_Change()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Text & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub Approved_Change()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Text & "/" & Description.Value & Loppu
End Sub
Private Sub Description_Change()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & Checked.Value & "/" & Approved.Value & "/" & Description.Text & Loppu
End Sub
Private Sub Lista_Click()
  If Lista.ListIndex > 0 Then
     TOtsikko.Caption = "Older Revision"
  Else
    If AddNew.Enabled Then
      TOtsikko.Caption = "Current Revision"
    Else
      TOtsikko.Caption = "New Revision"
    End If
  End If
  NaytaRev Nz(Lista.Value), Lista.ListIndex
End Sub
Private Sub Talleta_Click()
Dim koe As String
 If Lista.RowSource <> "" Then
   koe = Rev & Drawn & Checked & Approved & Description
   If InStr(koe, "/") Then
     MsgBox "Letter ""/"" is not allowed!", vbCritical, "Save revision"
     Exit Sub
   Else
     MRevRev = Rev.Value
     MRevDrawn = Drawn.Value
     MRevChecked = Checked.Value
     MRevApproved = Approved.Value
     MRevDescription = Description.Value
     Revisioteksti = Replace(Lista.RowSource, ";", vbCrLf)
   End If
 End If
 DoCmd.Close
End Sub
Private Sub AddNew_Click()
Dim UusiRev As String
  Alku = ""
  Loppu = Lista.RowSource
  If MRevRev = "" Then
    If Loppu <> "" Then
        MRevRev = Left$(Lista.RowSource, InStr(Lista.RowSource, " ") - 1)
        MRevRev = Chr(Asc(MRevRev) + 1)
    Else
      MRevRev = "0"
    End If
  End If
  
  UusiRev = MRevRev & " " & Date & "/" & MRevDrawn & "/" & MRevChecked & "/" & MRevApproved & "/" & MRevDescription
  Lista.RowSource = UusiRev & ";" & Lista.RowSource
  TOtsikko.Caption = "New Revision"
  Lista.SetFocus
  Lista.ListIndex = 0
  AddNew.Enabled = False
  NaytaRev UusiRev, 0
End Sub
Private Sub Sulje_Click()
  DoCmd.Close
End Sub
Private Sub NaytaRev(Tieto As String, Indeksi As Integer)
Dim Osoitin As Integer
Dim i As Integer
Tyhjenna
Alku = ""
Loppu = ""
For i = 0 To Indeksi - 1
  Alku = Alku & Lista.Column(0, i) & ";"
Next i
For i = Indeksi + 1 To Lista.ListCount
  Loppu = Loppu & ";" & Lista.Column(0, i)
Next i
'Revisio
Osoitin = InStr(Tieto, " ")
If Osoitin = 0 Then Exit Sub
Rev.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
'Päivämäärä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
DateFld.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
'Tekijä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Drawn.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
'Tarkastaja
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Checked.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
'Hyväksyjä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Approved.Value = Left$(Tieto, Osoitin - 1)
'Tieto = Mid(Tieto, Osoitin + 1)
'Kohde
'Osoitin = InStr(Tieto, "/")
'If Osoitin = 0 Then Exit Sub
'Kohde.VALUE = Left(Tieto, Osoitin - 1)
'Selitys
Description.Value = Mid$(Tieto, Osoitin + 1)
End Sub
Private Sub Tyhjenna()
  Rev.Value = ""
  'DateFld.VALUE = ""
  Drawn.Value = ""
  Checked.Value = ""
  Approved.Value = ""
  Description.Value = ""
End Sub
