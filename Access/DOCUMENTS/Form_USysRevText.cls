Option Compare Database
Option Explicit
'==========================================================================
Private Const TV_DOC_REV_TEXT As String = "DOC_Rev_Text"
' LOMAKE   : Form_USysRevText
' SOVELLUS : DOCUMENTS — Revisiotekstin muokkaus
' KUVAUS   : Hallinnoi dokumentin revisiohistoriaa. Revisiot tallennetaan
'            puolipisteellä eroteltuina merkkijonoina, jotka parsitaan
'            kenttiin näytöllä editointia varten.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Public Alku As String
Public Loppu As String
Private PaikallinenRev As String
Private PaikallinenDrawn As String
Private PaikallinenChecked As String
Private PaikallinenApproved As String
Private PaikallinenDescription As String
Private PaikallinenRevisioteksti As String

' Revisiokenttien tila pidetään paikallisena, jotta lomake toimii ilman ulkoisia globaaleja.
Private Sub Form_Resize()
  Dim RevTeksti As String

  RevTeksti = HaeRevisioTekstiTyoskentely()
  If RevTeksti <> "" Then
    Me.Lista.RowSource = Replace(CStr(RevTeksti), vbCrLf, ";")
    Me.Lista.SetFocus
    Me.Lista.ListIndex = 0
    If InStr(RevTeksti, vbCrLf) Then
      NaytaRev Left$(RevTeksti, InStr(RevTeksti, vbCrLf) - 1), 0
    Else
      NaytaRev CStr(RevTeksti), 0
    End If
  Else
    AddNew_Click
  End If
End Sub
Private Sub Rev_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Text & " " & Me.DateFld.Value & "/" & Me.Drawn.Value & "/" & Me.Checked.Value & "/" & Me.Approved.Value & "/" & Me.Description.Value & Loppu
End Sub
Private Sub DateFld_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Value & " " & Me.DateFld.Text & "/" & Me.Drawn.Value & "/" & Me.Checked.Value & "/" & Me.Approved.Value & "/" & Me.Description.Value & Loppu
End Sub
Private Sub Drawn_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Value & " " & Me.DateFld.Value & "/" & Me.Drawn.Text & "/" & Me.Checked.Value & "/" & Me.Approved.Value & "/" & Me.Description.Value & Loppu
End Sub
Private Sub Checked_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Value & " " & Me.DateFld.Value & "/" & Me.Drawn.Value & "/" & Me.Checked.Text & "/" & Me.Approved.Value & "/" & Me.Description.Value & Loppu
End Sub
Private Sub Approved_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Value & " " & Me.DateFld.Value & "/" & Me.Drawn.Value & "/" & Me.Checked.Value & "/" & Me.Approved.Text & "/" & Me.Description.Value & Loppu
End Sub
Private Sub Description_Change()
  Me.Lista.RowSource = Alku & Me.Rev.Value & " " & Me.DateFld.Value & "/" & Me.Drawn.Value & "/" & Me.Checked.Value & "/" & Me.Approved.Value & "/" & Me.Description.Text & Loppu
End Sub
Private Sub Lista_Click()
  If Me.Lista.ListIndex > 0 Then
     Me.TOtsikko.Caption = "Older Revision"
  Else
    If Me.AddNew.Enabled Then
      Me.TOtsikko.Caption = "Current Revision"
    Else
      Me.TOtsikko.Caption = "New Revision"
    End If
  End If
  NaytaRev Nz(Me.Lista.Value), Me.Lista.ListIndex
End Sub
Private Sub Talleta_Click()
Dim koe As String
 ' Tarkistetaan että lista ei ole tyhjä ennen tallennusta
 If Me.Lista.RowSource <> "" Then
   koe = Me.Rev.Value & Me.Drawn.Value & Me.Checked.Value & Me.Approved.Value & Me.Description.Value
   If InStr(koe, "/") Then
     MsgBox "Letter ""/""  is not allowed!", vbCritical, "Save revision"
     Exit Sub
   Else
     PaikallinenRev = Nz(Me.Rev.Value, "")
     PaikallinenDrawn = Nz(Me.Drawn.Value, "")
     PaikallinenChecked = Nz(Me.Checked.Value, "")
     PaikallinenApproved = Nz(Me.Approved.Value, "")
     PaikallinenDescription = Nz(Me.Description.Value, "")
     PaikallinenRevisioteksti = Replace(Me.Lista.RowSource, ";", vbCrLf)
     AsetaTempVarPaikallinen TV_DOC_REV_TEXT, PaikallinenRevisioteksti
   End If
 End If
 DoCmd.Close
End Sub
Private Sub AddNew_Click()
Dim UusiRev As String
  Alku = ""
  Loppu = Me.Lista.RowSource
 ' Lasketaan uuden revision kirjain edellisen perusteella
  If Loppu <> "" Then
    PaikallinenRev = Left$(Me.Lista.RowSource, InStr(Me.Lista.RowSource, " ") - 1)
      PaikallinenRev = Chr(Asc(PaikallinenRev) + 1)
  Else
    PaikallinenRev = "0"
  End If
  ' Käytetään lomakkeen kenttien arvoja uuden revision oletustietoina.
  PaikallinenDrawn = Nz(Me.Drawn.Value, "")
  PaikallinenChecked = Nz(Me.Checked.Value, "")
  PaikallinenApproved = Nz(Me.Approved.Value, "")
  PaikallinenDescription = Nz(Me.Description.Value, "")
  
  UusiRev = PaikallinenRev & " " & Date & "/" & PaikallinenDrawn & "/" & PaikallinenChecked & "/" & PaikallinenApproved & "/" & PaikallinenDescription
  Me.Lista.RowSource = UusiRev & ";" & Me.Lista.RowSource
  Me.TOtsikko.Caption = "New Revision"
  Me.Lista.SetFocus
  Me.Lista.ListIndex = 0
  Me.AddNew.Enabled = False
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
' Rakennetaan Alku-merkkijono valitun rivin yläpuolisista riveistä
For i = 0 To Indeksi - 1
  Alku = Alku & Me.Lista.Column(0, i) & ";"
Next i
' Rakennetaan Loppu-merkkijono valitun rivin alapuolisista riveistä
For i = Indeksi + 1 To Me.Lista.ListCount
  Loppu = Loppu & ";" & Me.Lista.Column(0, i)
Next i
' Revisio
Osoitin = InStr(Tieto, " ")
If Osoitin = 0 Then Exit Sub
Me.Rev.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
' Päivämäärä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Me.DateFld.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
' Tekijä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Me.Drawn.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
' Tarkastaja
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Me.Checked.Value = Left$(Tieto, Osoitin - 1)
Tieto = Mid$(Tieto, Osoitin + 1)
' Hyväksyjä
Osoitin = InStr(Tieto, "/")
If Osoitin = 0 Then Exit Sub
Me.Approved.Value = Left$(Tieto, Osoitin - 1)
' Selitys
Me.Description.Value = Mid$(Tieto, Osoitin + 1)
End Sub
Private Sub Tyhjenna()
  Me.Rev.Value = ""
  Me.Drawn.Value = ""
  Me.Checked.Value = ""
  Me.Approved.Value = ""
  Me.Description.Value = ""
End Sub

Private Function HaeRevisioTekstiTyoskentely() As String
  ' Käytetään TempVars-välitystä revisiotekstin välittämiseen
  HaeRevisioTekstiTyoskentely = LueTempVarTekstiPaikallinen(TV_DOC_REV_TEXT)
End Function

Private Sub AsetaTempVarPaikallinen(ByVal Nimi As String, ByVal Arvo As String)
  On Error Resume Next
  TempVars.Remove Nimi
  On Error GoTo 0
  TempVars.Add Nimi, Arvo
End Sub

Private Function LueTempVarTekstiPaikallinen(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTekstiPaikallinen = Nz(TempVars(Nimi).Value, "")
  Exit Function

ErrorHandler:
  LueTempVarTekstiPaikallinen = ""
End Function