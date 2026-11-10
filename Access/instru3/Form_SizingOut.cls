Option Compare Database
Option Explicit

'================================================================================
' Lomake: SizingOut
' Tarkoitus: Vie kyselytulokset CSV-tiedostoon suomalaisella lukumuotoilulla
' Päivitetty: 2025-11-11 - VBA7/64-bit-tuki lisätty
'             2026-03-03 - Me.-etuliitteet ja kommentit suomeksi
'
' Kuvaus:
'   Vie tietokantakyselyn tai -taulun puolipisteeroitettuun CSV-tiedostoon:
'   - Käyttäjä valitsee kyselyn/taulun, hakemiston ja tiedostonimen
'   - Muuntaa desimaalipilkut pisteiksi (suomalainen -> kansainvälinen muoto)
'   - Ympy röi kaikki arvot lainausmerkeillä
'   - Avaa kohdehakemiston Explorerissa viennin jälkeen
'   - Käsittelee null-arvot tyhjinä merkkijonoina
'
' Riippuvuudet:
'   - FileSystemObject (myöhäinen sidonta)
'   - Taulukko-kontrollissa määritetty kysely tai taulu
'
' Kontrollit:
'   - Taulukko: TextBox kyselyn/taulun nimelle (oletus: qryValveSizingOut)
'   - Hakem: TextBox kohdehakemiston polkua varten
'   - FileName: TextBox tiedostonimelle (ilman .csv-päätettä)
'   - Command8: Painike viennin käynnistämiseen
'================================================================================

'--------------------------------------------------------------------------------
' Tapahtuma: Command8_Click
' Tarkoitus: Vie valittu kysely/taulu CSV-tiedostoon
'
' Prosessi:
'   1. Tarkistetaan, että kohdehakemisto on olemassa
'   2. Avataan lähde-kysely/taulu recordset-muodossa
'   3. Luodaan CSV-tiedosto kohdepolkuun
'   4. Iteroidaan kaikki rivit:
'      - Kirjoitetaan ensimmäinen kenttä (lainausmerkeissä)
'      - Kirjoitetaan loput kentät puolipisteellä eroteltuina
'      - Muunnetaan desimaalipilkut pisteiksi
'      - Null-arvot kirjoitetaan tyhjinä merkkijonoina
'   5. Suljetaan tiedosto ja avataan kohdehakemisto
'
' CSV-muoto:
'   - Puolipisteerotteinen (;)
'   - Kaikki arvot lainausmerkeissä ("arvo")
'   - Desimaalierotin: piste (.) eikä pilkku (,)
'   - Rivinloppu: CRLF (Windows-standardi)
'
' Huomiot:
'   - Varmistaa, että hakemistopolku päättyy kenoviivaan
'   - Muuntaa suomalaisen desimaaliformaatin kansainväliseksi
'   - Avaa hakemiston käyttäjän helpottamiseksi
'   - Kattava virheenkäsittely siivouksineen
'--------------------------------------------------------------------------------
Private Sub Command8_Click()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Nykyinen tietokanta
    Dim Taulu As DAO.Recordset  ' Lähdedata-recordset
    Dim FSO As Object  ' Tiedostojärjestelmäobjekti tiedoston luomiseen (myöhäinen sidonta)
    Dim Tiedosto As Object  ' CSV-tiedoston kahva
    Dim Path As String  ' Kohdehakemiston polku
    Dim i As Integer  ' Kenttälaskuri
    Dim Osoitin As Long  ' Pilkun sijainti merkkijonossa
    Dim Numeerinen As Boolean  ' Onko kenttä numeerinen (ei formula-injection-suojausta)
    Dim Arvo As String  ' Nykyinen kenttäarvo
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set DB = CurrentDb
    Path = Me.Hakem.Value
    
    ' Varmistetaan, että polku päättyy kenoviivaan
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    ' Tarkistetaan, että kohdehakemisto on olemassa
    If Not FSO.FolderExists(Path) Then
        MsgBox "Hakemistoa: " & Path & " ei ollut", vbCritical, "Vientivirhe"
        Exit Sub
    End If
    
    ' Avataan lähde-kysely/taulu
    If Not LahdeOlemassa(DB, CStr(Nz(Me.Taulukko.Value, ""))) Then
        MsgBox "Taulua tai kyselyä '" & Me.Taulukko.Value & "' ei löydy", vbCritical, "Vientivirhe"
        Exit Sub
    End If
    Set Taulu = DB.OpenRecordset("SELECT * FROM [" & Me.Taulukko.Value & "]")
    
    ' Luodaan CSV-tiedosto kohdepolkuun
    Set Tiedosto = FSO.CreateTextFile(Path & Me.FileName.Value & ".csv")
    
    ' Viedään kaikki rivit
    Do While Not Taulu.EOF
        For i = 0 To Taulu.Fields.Count - 1
            Select Case Taulu.Fields(i).Type
                Case dbByte, dbInteger, dbLong, dbBigInt, dbCurrency, dbSingle, dbDouble, dbDecimal
                    Numeerinen = True
                Case Else
                    Numeerinen = False
            End Select

            ' Null-arvot kirjoitetaan tyhjinä
            If IsNull(Taulu.Fields(i).Value) Then
                Arvo = ""
            Else
                Arvo = CStr(Taulu.Fields(i).Value)

                ' Desimaalipilkku muutetaan pisteeksi vain luvuille (tekstikenttien pilkut säilyvät)
                If Numeerinen Or IsNumeric(Arvo) Then
                    Osoitin = InStr(Arvo, ",")
                    If Osoitin <> 0 Then
                        Arvo = Left$(Arvo, Osoitin - 1) & "." & Mid$(Arvo, Osoitin + 1)
                    End If
                End If
            End If

            ' Formula Injection -esto: heittomerkki vaarallisten =-@+-alkuisten tekstiarvojen eteen
            ' (numeerisia kenttiä, esim. negatiivisia lukuja, ei muokata)
            If Arvo <> "" And Not Numeerinen Then
                If Left$(Arvo, 1) = "=" Or Left$(Arvo, 1) = "@" Or _
                   Left$(Arvo, 1) = "+" Or Left$(Arvo, 1) = "-" Then
                    Arvo = "'" & Arvo
                End If
            End If
            ' RFC 4180: tuplataan lainausmerkit, jottei CSV-rakenne rikkoudu
            Arvo = Replace(Arvo, """", """""")

            If i > 0 Then Tiedosto.Write ";"
            Tiedosto.Write """" & Arvo & """"
        Next i

        Tiedosto.Write vbCrLf
        Taulu.MoveNext
    Loop
    
    ' Siivotaan resurssit
    Taulu.Close
    Tiedosto.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    
    ' Ilmoitetaan käyttäjälle ja avataan hakemisto
    MsgBox "Valmis", vbInformation, "Vienti valmis"
    Call Shell("EXPLORER """ & Path & """", vbNormalFocus)
    Exit Sub

ErrorHandler:
    MsgBox "Virhe tiedon viennissä: " & Err.Description, vbCritical, "Vientivirhe"
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    If Not Tiedosto Is Nothing Then Tiedosto.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Load
' Tarkoitus: Alustaa lomakkeen oletusarvoilla
'
' Oletukset:
'   - Taulukko (Kysely/taulu): qryValveSizingOut
'   - Hakem (Hakemisto): Käyttäjän määritettävä
'
' Huomiot:
'   - Oletuskysely kohdistuu venttiilien mitoitusdatan vientiin
'--------------------------------------------------------------------------------
' Tarkistaa, että annettu nimi on olemassa oleva taulu tai kysely (estää mielivaltaisen SQL:n)
Private Function LahdeOlemassa(ByVal DB As DAO.Database, ByVal Nimi As String) As Boolean
    Dim Td As DAO.TableDef  ' Tarkistettava taulu
    Dim Qd As DAO.QueryDef  ' Tarkistettava kysely

    For Each Td In DB.TableDefs
        If StrComp(Td.Name, Nimi, vbTextCompare) = 0 Then
            LahdeOlemassa = True
            Exit Function
        End If
    Next Td
    For Each Qd In DB.QueryDefs
        If StrComp(Qd.Name, Nimi, vbTextCompare) = 0 Then
            LahdeOlemassa = True
            Exit Function
        End If
    Next Qd
End Function

Private Sub Form_Load()
  Me.Taulukko.Value = "qryValveSizingOut"  ' Oletuskysely venttiilien mitoitusdatan vientiin
End Sub
