Option Compare Database
Option Explicit

'================================================================================
Private Const POISSULJETTU_TAULU As String = "MotorData"
Private Const LINKITETTY_TAULU_TYYPPI As Long = 6

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
'      - Jos polut eroavat, linkki päivitetään nykyiseen hakemistoon
'   4. Säilytetään taulun nimi ja rakenne
' Huom:
'   - Linkittää uudelleen vain, jos lähdehakemisto eroaa nykyisestä
'   - Käyttää kirjainkoosta riippumatonta polkuvertailua
'   - Estää saman taulun kaksinkertaisen uudelleenlinkityksen
'   - Käyttää TableDef.RefreshLink-menetelmää linkkien päivittämiseen
'--------------------------------------------------------------------------------
Private Sub Command0_Click()
  Dim Polku As String, Polku2 As String, Nimi As String
  Dim Kohdetiedosto As String, Taulu As String
  Dim Taul As DAO.Recordset, db As DAO.Database, tdf As DAO.TableDef
  Dim kasitellyt As New Collection, epaonnistuneet As String
  Dim viimeinenKauttaviiva As Long
  Dim virhenumero As Long, virhekuvaus As String

  On Error GoTo ErrorHandler
  Set db = CurrentDb
  Polku = Left$(db.Name, Len(db.Name) - Len(Dir$(db.Name)))
  Set Taul = db.OpenRecordset( _
    "SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects " & _
    "WHERE ([Database] Is Not Null) AND ([Name] Not Like '" & POISSULJETTU_TAULU & "') " & _
    "AND ([Type] = " & LINKITETTY_TAULU_TYYPPI & ")", dbOpenSnapshot, dbForwardOnly)

  Do While Not Taul.EOF
    Taulu = Taul.Fields("Name")
    If Not JoKasitelty(kasitellyt, Taulu) Then
      kasitellyt.Add Taulu, Taulu
      viimeinenKauttaviiva = InStrRev(Taul.Fields("Database"), "\")
      If viimeinenKauttaviiva > 0 Then
        Polku2 = LCase$(Left$(Taul.Fields("Database"), viimeinenKauttaviiva))
        Nimi = Mid$(Taul.Fields("Database"), viimeinenKauttaviiva + 1)
        If Polku2 <> LCase$(Polku) Then
          Kohdetiedosto = Polku & Nimi
          If Len(Dir$(Kohdetiedosto)) = 0 Then
            epaonnistuneet = epaonnistuneet & "  - " & Taulu & " (tietokantaa ei löydy: " & Kohdetiedosto & ")" & vbCrLf
          Else
            Set tdf = db.TableDefs(Taulu)
            tdf.Connect = ";DATABASE=" & Kohdetiedosto
            On Error Resume Next
            tdf.RefreshLink
            If Err.Number <> 0 Then
              epaonnistuneet = epaonnistuneet & "  - " & Taulu & " (" & Err.Description & ")" & vbCrLf
              Err.Clear
            End If
            On Error GoTo ErrorHandler
            Set tdf = Nothing
          End If
        End If
      End If
    End If
    Taul.MoveNext
  Loop

  Taul.Close
  Set Taul = Nothing
  Set db = Nothing
  If Len(epaonnistuneet) > 0 Then
    MsgBox "Osa taulujen linkeistä ei päivittynyt:" & vbCrLf & epaonnistuneet, vbExclamation
  Else
    MsgBox "Valmis", vbInformation
  End If
  Exit Sub

ErrorHandler:
  virhenumero = Err.Number
  virhekuvaus = Err.Description
  On Error Resume Next
  If Not Taul Is Nothing Then Taul.Close
  Set Taul = Nothing
  Set tdf = Nothing
  Set db = Nothing
  On Error GoTo 0
  MsgBox "Virhe tauluyhteyksien päivityksessä (" & virhenumero & "): " & virhekuvaus, vbCritical
End Sub

Private Function JoKasitelty(ByRef lista As Collection, ByVal nimi As String) As Boolean
  On Error Resume Next
  Dim testi As Variant
  testi = lista(nimi)
  JoKasitelty = (Err.Number = 0)
  Err.Clear
  On Error GoTo 0
End Function
