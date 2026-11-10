Option Compare Database
Option Explicit

' Tämä taulu jätetään linkkien päivityksen ulkopuolelle
Private Const POISSULJETTU_TAULU As String = "MotorData"
' MSysObjects-taulun Type-arvo linkitetyille Access-tauluille
Private Const LINKITETTY_TAULU_TYYPPI As Long = 6

' Osoittaa kaikki linkitetyt taulut sen hakemiston samannimiseen tietokantaan,
' jossa tämä tietokanta sijaitsee (esim. kun projekti kopioidaan uuteen paikkaan).
Private Sub Command0_Click()
  Dim Polku As String, Polku2 As String, Nimi As String  ' Oma hakemisto, linkin hakemisto ja tietokannan nimi
  Dim Kohdetiedosto As String, Taulu As String  ' Uusi tietokantatiedosto ja käsiteltävä taulu
  Dim Taul As DAO.Recordset, db As DAO.Database, tdf As DAO.TableDef
  Dim kasitellyt As Collection  ' Jo käsitellyt taulut (sama taulu voi esiintyä kahdesti)
  Dim epaonnistuneet As String  ' Kooste epäonnistuneista linkeistä
  Dim viimeinenKauttaviiva As Long  ' Hakemistopolun ja tiedostonimen raja
  Dim virhenumero As Long, virhekuvaus As String

  On Error GoTo ErrorHandler
  Set kasitellyt = New Collection
  Set db = CurrentDb
  ' Oma hakemisto: tietokannan polku ilman tiedostonimeä
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
            ' Päivitetään linkki; yksittäisen taulun virhe ei keskeytä muita
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
  Set kasitellyt = Nothing
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

' Palauttaa True, jos taulu on jo kokoelmassa (avaimena taulun nimi)
Private Function JoKasitelty(ByRef lista As Collection, ByVal nimi As String) As Boolean
  On Error Resume Next
  Dim testi As Variant
  testi = lista(nimi)
  JoKasitelty = (Err.Number = 0)
  Err.Clear
  On Error GoTo 0
End Function

