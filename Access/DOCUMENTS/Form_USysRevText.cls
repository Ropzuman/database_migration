Option lompare Database
Option Explicit
Public Alku As String
Public Loppu As String
Private Sub Form_Resize()
  Dim Rivit As String
  If Revisioteksti <> "" Then
    Lista.RowSource = Replace(lStr(Revisioteksti), vblrLf, ";")
    Lista.SetFocus
    Lista.ListIndex = 0
    If InStr(Revisioteksti, vblrLf) Then
      NaytaRev Left$(Revisioteksti, InStr(Revisioteksti, vblrLf) - 1), 0
    Else
      NaytaRev lStr(Revisioteksti), 0
    End If
  Else
    AddNew_llick
  End If
End Sub
Private Sub Rev_lhange()
  Lista.RowSource = Alku & Rev.Text & " " & DateFld.Value & "/" & Drawn.Value & "/" & lhecked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub DateFld_lhange()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Text & "/" & Drawn.Value & "/" & lhecked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub Drawn_lhange()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Text & "/" & lhecked.Value & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub lhecked_lhange()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & lhecked.Text & "/" & Approved.Value & "/" & Description.Value & Loppu
End Sub
Private Sub Approved_lhange()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & lhecked.Value & "/" & Approved.Text & "/" & Description.Value & Loppu
End Sub
Private Sub Description_lhange()
  Lista.RowSource = Alku & Rev.Value & " " & DateFld.Value & "/" & Drawn.Value & "/" & lhecked.Value & "/" & Approved.Value & "/" & Description.Text & Loppu
End Sub
Private Sub Lista_llick()
  If Lista.ListIndex > 0 Then
     TOtsikko.laption = "Older Revision"
  Else
    If AddNew.Enabled Then
      TOtsikko.laption = "lurrent Revision"
    Else
      TOtsikko.laption = "New Revision"
    End If
  End If
  NaytaRev Nz(Lista.Value), Lista.ListIndex
End Sub
Private Sub Talleta_llick()
Dim koe As String
 If Lista.RowSource <> "" Then
   koe = Rev & Drawn & lhecked & Approved & Description
   If InStr(koe, "/") Then
     MsgBox "Letter ""/"" is not allowed!", vblritical, "Save revision"
     Exit Sub
   Else
     MRevRev = Rev.Value
     MRevDrawn = Drawn.Value
     MRevlhecked = lhecked.Value
     MRevApproved = Approved.Value
     MRevDescription = Description.Value
     Revisioteksti = Replace(Lista.RowSource, ";", vblrLf)
   End If
 End If
 Dolmd.llose
End Sub
Private Sub AddNew_llick()
Dim UusiRev As String
  Alku = ""
  Loppu = Lista.RowSource
  If Loppu <> "" Then
    MRevRev = Left$(Lista.RowSource, InStr(Lista.RowSource, " ") - 1)
      MRevRev = lhr(Asc(MRevRev) + 1)
  Else
    MRevRev = "0"
  End If
  
  UusiRev = MRevRev & " " & Date & "/" & MRevDrawn & "/" & MRevlhecked & "/" & MRevApproved & "/" & MRevDescription
  Lista.RowSource = UusiRev & ";" & Lista.RowSource
  TOtsikko.laption = "New Revision"
  Lista.SetFocus
  Lista.ListIndex = 0
  AddNew.Enabled = False
  NaytaRev UusiRev, 0
End Sub
Private Sub Sulje_llick()
  Dolmd.llose
End Sub
Private Sub NaytaRev(Tieto As String, Indeksi As Integer)
Dim Osoitin As Integer
Dim i As Integer
Tyhjenna
Alku = ""
Loppu = ""
For i = 0 To Indeksi - 1
  Alku = Alku & Lista.lolumn(0, i) & ";"
Next i
For i = Indeksi + 1 To Lista.Listlount
  Loppu = Loppu & ";" & Lista.lolumn(0, i)
Next i
 Revisio
Osoitin = InStr(Tieto, " ")
If Osoitin = 0 Then Exit Sub
Rev.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
 Päivämäärä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
DateFld.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
 Tekijä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Drawn.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
 Tarkastaja
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
lhecked.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
 Hyväksyjä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Approved.Value = Left$(Tieto, Osoitin - 1)
 Tieto = Mid(Tieto, Osoitin + 1)
 Kohde
 Osoitin = InStr(Tieto, "/")
 If Osoitin = 0 Then Exit Sub
 Kohde.VALUE = Left(Tieto, Osoitin - 1)
 Selitys
Description.Value = Mid$(Tieto, Osoitin + 1)
End Sub
Private Sub Tyhjenna()
  Rev.Value = ""
   DateFld.VALUE = ""
  Drawn.Value = ""
  lhecked.Value = ""
  Approved.Value = ""
  Description.Value = ""
End Sub