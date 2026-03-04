Option Compare Database
Option Explicit

'================================================================================
' Lomake: Linkkien vaihto
' Tarkoitus: Päivittää linkitettyjen taulujen polut nykyiseen tietokannan hakemistoon
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki lisätty
'
' Kuvaus:
'   Apulomake ulkoisten taulujen uudelleenlinkitykseen nykyiseen hakemistoon:
'   - Skannaa MSysObjects-järjestelmätaulun linkitetyistä tauluista
'   - Poimii tietokannan tiedostonimen nykyisestä linkityspolusta
'   - Linkittää taulun samaan tiedostoon nykyisessä hakemistossa
'   - Estää väärään hakemistoon osoittavat linkit tietokannan siirron jälkeen
'
' Riippuvuudet:
'   - MSysObjects-järjestelmätaulu (Accessin sisäinen)
'   - Linkitetyt Access-tietokantataulut
'
' Käyttötapaus:
'   Kun tietokanta siirretään uuteen kansioon, linkitetyt taulut osoittavat
'   edelleen vanhaan sijaintiin. Tämä lomake päivittää kaikki linkit nykyiseen
'   hakemistoon.
'================================================================================

'--------------------------------------------------------------------------------
' Tapahtuma: Command0_Click
' Tarkoitus: Linkittää kaikki ulkoiset taulut nykyiseen tietokannan hakemistoon
' Toiminta:
'   1. Haetaan nykyinen tietokannan hakemistopolku
'   2. Kysellään MSysObjects-taulusta kaikki linkitetyt taulut
'   3. Jokaiselle linkitetylle taululle:
'      - Poimitaan tiedostonimi nykyisestä linkityspolusta
'      - Vertaillaan hakemistopolkuja
'      - Jos polut eroavat, pudotetaan ja luodaan uudelleen linkki nykyiseen hakemistoon
'   4. Säilytetään taulun nimi ja rakenne
' Huom:
'   - Linkittää uudelleen vain, jos lähdehakemisto eroaa nykyisestä
'   - Käyttää kirjainkoosta riippumatonta polkuvertailua
'   - Estää saman taulun kaksinkertaisen uudelleenlinkityksen
'   - Käyttää DoCmd.TransferDatabase-komentoa linkkien uudelleenluomiseen
'--------------------------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer  ' Silmukkalasku
    Dim Polku As String             ' Nykyinen tietokannan hakemistopolku
    Dim Polku2 As String            ' Linkitetyn taulun hakemistopolku
    Dim Nimi As String              ' Tietokannan tiedostonimi
    Dim Taul As DAO.Recordset       ' MSysObjects-kyselyn tulokset
    Dim Taulu As String             ' Käsiteltävän taulun nimi
    Dim LowerPolku As String        ' Nykyinen polku pieniksi kirjaimiksi vertailua varten
    Dim LinksUpdated As Integer     ' Päivitettyjen linkkien laskuri

    ' Alustetaan laskuri
    LinksUpdated = 0
    
    ' Haetaan nykyinen tietokannan hakemistopolku
    Polku = Left$(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir$(CurrentDb.Name)))
    LowerPolku = LCase$(Polku)  ' Välimuistiin pieniksi kirjaimiksi vertailua varten
    
    ' Kysellään linkitetyt taulut
    Set Taul = CurrentDb.OpenRecordset("SELECT DATABASE, ForeignName, Name FROM MSysObjects WHERE Database Is Not Null")
    
    If Taul.EOF Then
        MsgBox "Ei linkitettyjä tauluja löytynyt", vbInformation, "Link Update"  ' "No linked tables found"
        Taul.Close
        Set Taul = Nothing
        Exit Sub
    End If
    
    ' Käsitellään jokainen linkitetty taulu
    Do While Not Taul.EOF
      ' Estetään saman taulun käsittely useampaan kertaan
      If Taul.Fields("Name") <> Taulu Then
          ' Poimitaan hakemisto ja tiedostonimi nykyisestä linkityspolusta
          For i = Len(Taul.Fields("Database")) To 1 Step -1
            If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
              Polku2 = LCase$(Left$(Taul.Fields("Database"), i))  ' Directory path
              Nimi = Mid$(Taul.Fields("Database"), i + 1)  ' Filename
              Exit For
            End If
          Next i
          
          ' Check if link path differs from current database path
          If Polku2 <> LowerPolku Then
            Taulu = Taul.Fields("Name")
            
            ' Pudotetaan nykyinen linkitetty taulu
            CurrentDb.Execute "DROP TABLE [" & Taulu & "]"
            
            ' Luodaan linkki uudelleen nykyiseen hakemistoon
            On Error Resume Next
            DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
            If Err = 0 Then
                LinksUpdated = LinksUpdated + 1
            End If
            Err.Clear
            On Error GoTo ErrorHandler
          End If
      End If
      Taul.MoveNext
    Loop
    
    ' Siivotaan objektit
    Taul.Close
    Set Taul = Nothing
    
    If LinksUpdated > 0 Then
        MsgBox "Päivitetty " & LinksUpdated & " linkkiä", vbInformation, "Link Update"  ' "Updated X links"
    Else
        MsgBox "Ei linkkejä päivitettävänä", vbInformation, "Link Update"  ' "No links to update"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Virhe linkkien päivittämisessä: " & Err.Description, vbCritical, "Linkkien vaihto"
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    On Error GoTo 0
End Sub
