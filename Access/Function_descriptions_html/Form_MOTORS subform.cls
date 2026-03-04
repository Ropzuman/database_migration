Option Compare Database
Option Explicit

'================================================================================
' Lomake: MOTORS subform (Moottorit-alilomake)
' Tarkoitus: Laiteviittauksen lisäämisapuri päälomakkeen tekstikenttiin
' Päivitetty: 2025-11-13 — Virheenkäsittely ja null-suojaus lisätty
'             2026-03-03 — Me.-etuliite, SelStart-suojaus (Virhe 2185)
'
' Kuvaus:
'   Alilomake laiteviittausten lisäämiseksi päälomakkeen tekstikenttiin.
'   Käyttää {Osasto-Laitenro-Laitetyyppi-Liite}-muotoa.
'
' Riippuvuudet:
'   - Päälomake, jossa KohdeTextBox ja CLoppuun-ohjaimet
'   - Julkinen Kursori-muuttuja (kursorin sijainti)
'   - Ohjaimet: TEKSTI, Department, Eqseq, EqType, Suffix
'================================================================================

'================================================================================
' Aliohjelma: LisaaTeksti
' Tarkoitus: Lisää laiteviittaus kursorin kohdalle tai kentän loppuun
'
' Kuvaus:
'   Lisää tekstin {hakasulkeissa} nykyiseen kursorin sijaintiin tai kentän
'   loppuun. Käsittelee null-arvot ja yllapitää kursorin sijainnin.
' Huom: SelStart vaatii fokuksen — kietotaan virheensuojauksen sisään (Virhe 2185)
'================================================================================
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    
    On Error GoTo ErrorHandler
    
    If Not KohdeTextBox Is Nothing Then
        If Form.Parent.CLoppuun = True Then
            ' Lisätään kentän loppuun
            Alku = Nz(KohdeTextBox.Value, "")
            Loppu = ""
        Else
            ' Lisätään kursorin kohtaan
            If IsNull(KohdeTextBox.Value) Then
                Alku = ""
                Loppu = ""
            Else
                Alku = Left$(KohdeTextBox.Value, Kursori)
                Loppu = Mid$(KohdeTextBox.Value, Kursori + 1)
            End If
        End If
        
        KohdeTextBox.Value = Alku & " {" & Me.TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        ' SelStart ja SelLength vaativat fokuksen — estetään Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = Kursori + Len(Me.TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe tekstin lisäyksessä: " & Err.Description, vbExclamation
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

