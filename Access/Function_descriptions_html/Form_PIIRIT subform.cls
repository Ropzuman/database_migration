Option Compare Database
Option Explicit

'================================================================================
' Lomake: PIIRIT subform (Piirit-alilomake)
' Tarkoitus: Piiriviittauksen lisäämisapuri päälomakkeen tekstikenttiin
' Päivitetty: 2025-11-13 — Virheenkäsittely ja null-suojaus lisätty
'             2026-03-03 — Me.-etuliite, SelStart-suojaus (Virhe 2185)
'
' Kuvaus:
'   Alilomake piiriviittausten lisäämiseksi päälomakkeen tekstikenttiin.
'   Käyttää {AlueKoodi-PiiriNumero}-muotoa.
'
' Riippuvuudet:
'   - Päälomake, jossa KohdeTextBox ja CLoppuun-ohjaimet
'   - Julkinen Kursori-muuttuja (kursorin sijainti)
'   - Ohjaimet: TEKSTI, AreaCode, LoopNo
'================================================================================

'================================================================================
' Aliohjelma: LisaaTeksti
' Tarkoitus: Lisää piiriviittaus kursorin kohdalle tai kentän loppuun
'
' Kuvaus:
'   Lisää tekstin {hakasulkeissa} nykyiseen kursorin sijaintiin tai kentän
'   loppuun. Käsittelee null-arvot ja yllapitää kursorin sijainnin.
' Huom: SelStart vaatii fokuksen — kietotaan virheensuojauksen sisään (Virhe 2185)
'================================================================================
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    Dim KohdeTeksti As String
    Dim TurvallinenKursori As Long  ' Rajattu kursori — ei voi ylitää tekstin pituutta

    On Error GoTo ErrorHandler

    If Not KohdeTextBox Is Nothing Then
        KohdeTeksti = Nz(KohdeTextBox.Value, "")
        ' Rajataan kursori tekstin pituuden sisään — estää Left$/Mid$-virheet fokuksenvaihdon jälkeen
        TurvallinenKursori = Kursori
        If TurvallinenKursori > Len(KohdeTeksti) Then TurvallinenKursori = Len(KohdeTeksti)
        If TurvallinenKursori < 0 Then TurvallinenKursori = 0

        If Form.Parent.CLoppuun = True Then
            ' Lisätään kentän loppuun
            Alku = KohdeTeksti
            Loppu = ""
        Else
            ' Lisätään kursorin kohtaan
            Alku = Left$(KohdeTeksti, TurvallinenKursori)
            Loppu = Mid$(KohdeTeksti, TurvallinenKursori + 1)
        End If

        KohdeTextBox.Value = Alku & " {" & Me.TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        ' SelStart ja SelLength vaativat fokuksen — estätään Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = TurvallinenKursori + Len(Me.TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Virhe tekstin lisäyksessä: " & Err.Description, vbExclamation
End Sub

Private Sub AreaCode_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub LoopNo_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub

Private Sub TEKSTI_DblClick(Cancel As Integer)
    LisaaTeksti
End Sub


