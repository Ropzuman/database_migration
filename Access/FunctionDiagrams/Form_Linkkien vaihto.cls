Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_Linkkien vaihto
' SOVELLUS : FunctionDiagrams — Tietokantayhteyksien päivitys
' KUVAUS   : Päivittää kaikki linkitetyt taulut nykyiseen hakemistoon.
'            Lukee MSysObjects-järjestelmätaulusta linkitettyjen taulujen
'            polut, vertaa nykyiseen kannan sijaintiin ja uudelleenlinkittää
'            poikkeamat TableDef.RefreshLink-menetelmällä. Tarvitaan kun kanta
'            siirretään uuteen hakemistoon.
' HUOM     : Käsittelee vain Jet/ISAM-tyyppisiä linkitettyjä tauluja
'            (MSysObjects.Type = 6). ODBC-linkit eivät ole polkumuotoisia
'            eikä niitä pidä syöttää polun jäsennykseen.
' PÄIVITETTY: 2026-03-03 (tarkistettu ja korjattu)
'==========================================================================

Private Const POISSULJETTU_TAULU As String = "MotorData"
Private Const LINKITETTY_TAULU_TYYPPI As Long = 6 ' MSysObjects.Type: linkitetty Jet/ISAM-taulu

Private Sub Command0_Click()
On Error GoTo ErrorHandler

Debug.Print "Command0_Click: Tietokantataulujen linkkien päivitys alkaa"

Dim Polku As String
Dim Polku2 As String
Dim Nimi As String
Dim Taul As DAO.Recordset
Dim Taulu As String
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim kasitellyt As New Collection      ' jo käsitellyt taulujen nimet
Dim epaonnistuneet As String          ' RefreshLink-kutsussa epäonnistuneet taulut
Dim viimeinenKauttaviiva As Long
Dim Kohdetiedosto As String
Dim virhenumero As Long
Dim virhekuvaus As String
Dim vanhaYhteys As String

Set db = CurrentDb
Polku = Left$(db.Name, Len(db.Name) - Len(Dir(db.Name)))
Debug.Print "  Nykyisen tietokannan polku: " & Polku

Set Taul = db.OpenRecordset( _
    "SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects " & _
    "WHERE ([Database] Is Not Null) " & _
    "AND ([Name] Not Like '" & POISSULJETTU_TAULU & "') " & _
    "AND ([Type] = " & LINKITETTY_TAULU_TYYPPI & ")", _
    dbOpenSnapshot, dbForwardOnly)

Debug.Print "  Käsitellään linkitettyjä tauluja..."
Do While Not Taul.EOF
  Taulu = Taul.Fields("Name")

  ' Ohita taulu jos se on jo käsitelty (esim. useampi rivi samalle nimelle)
  If Not JoKasitelty(kasitellyt, Taulu) Then
    kasitellyt.Add Taulu, Taulu
    Debug.Print "  Tarkistetaan taulua: " & Taulu

    viimeinenKauttaviiva = InStrRev(Taul.Fields("Database"), "\")
    If viimeinenKauttaviiva = 0 Then
      ' Ei tunnistettua tiedostopolkua (esim. ODBC-linkki) - ohitetaan turvallisuussyistä
      Debug.Print "  OHITETTU (tiedostopolkua ei löytynyt, ei ehkä Jet-linkki): " & Taul.Fields("Database")
    Else
      Polku2 = LCase$(Left$(Taul.Fields("Database"), viimeinenKauttaviiva))
      Nimi = Mid$(Taul.Fields("Database"), viimeinenKauttaviiva + 1)

      If Polku2 <> LCase$(Polku) Then
        Debug.Print "  Päivitetään taulun linkki: " & Taulu & ", vanha polku: " & Taul.Fields("Database")

        Kohdetiedosto = Polku & Nimi
        If Len(Dir$(Kohdetiedosto)) = 0 Then
          Debug.Print "  LINKIN PÄIVITYS EPÄONNISTUI: kohdetietokantaa ei löytynyt: " & Kohdetiedosto
          epaonnistuneet = epaonnistuneet & "  - " & Taulu & " (tietokantaa ei löydy: " & Kohdetiedosto & ")" & vbCrLf
        Else
          ' Päivitä yhteys rikkomatta taulun rakennetta tai relaatioita.
          Set tdf = db.TableDefs(Taulu)
          vanhaYhteys = tdf.Connect
          ' Vaihdetaan vain tiedostopolku, jotta muut yhteysmääreet (esim. salasana) säilyvät
          If InStr(1, vanhaYhteys, Taul.Fields("Database"), vbTextCompare) > 0 Then
            tdf.Connect = Replace(vanhaYhteys, Taul.Fields("Database"), Kohdetiedosto, , , vbTextCompare)
          Else
            tdf.Connect = ";DATABASE=" & Kohdetiedosto
          End If

          On Error Resume Next
          tdf.RefreshLink
          If Err.Number <> 0 Then
            Debug.Print "  LINKIN PÄIVITYS EPÄONNISTUI: " & Err.Number & " - " & Err.Description
            epaonnistuneet = epaonnistuneet & "  - " & Taulu & " (" & Err.Description & ")" & vbCrLf
            ' Palautetaan alkuperäinen yhteys, ettei linkki jää rikki
            tdf.Connect = vanhaYhteys
            Err.Clear
          Else
            Debug.Print "  Linkki päivitetty kohteeseen: " & Kohdetiedosto
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
  Debug.Print "Command0_Click: VALMIS, MUTTA VIRHEITÄ TAPAHTUI"
  MsgBox "Osa taulujen linkeistä ei päivittynyt:" & vbCrLf & epaonnistuneet, vbExclamation
Else
  Debug.Print "Command0_Click: VALMIS - kaikkien taulujen linkit päivitetty"
  MsgBox "Valmis"
End If

Exit Sub

ErrorHandler:
  virhenumero = Err.Number
  virhekuvaus = Err.Description
  Debug.Print "*** VIRHE aliohjelmassa Command0_Click: " & virhenumero & " - " & virhekuvaus
    Debug.Print "    Nykyinen polku: " & Polku
    Debug.Print "    Taulu: " & Taulu
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    Set tdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    MsgBox "Virhe tauluyhteyksien päivityksessä (" & virhenumero & "): " & virhekuvaus, vbCritical
End Sub

' Palauttaa True jos taulun nimi on jo listassa (Collection-pohjainen haku).
Private Function JoKasitelty(ByRef lista As Collection, ByVal nimi As String) As Boolean
    On Error Resume Next
    Dim testi As Variant
    testi = lista(nimi)
    JoKasitelty = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function