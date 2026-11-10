Option Compare Database
Option Explicit

Private Const POISSULJETTU_TAULU As String = "MotorData"
Private Const LINKITETTY_TAULU_TYYPPI As Long = 6

'================================================================================
' Lomake: Linkkien vaihto
' Tarkoitus: P‰ivitt‰‰ linkitettyjen taulujen polut nykyisen tietokannan hakemistoon
' P‰ivitetty: 2025-11-11 - VBA7/64-bit-tuki lis‰tty
'             2026-11-10 - K‰ytt‰m‰ttˆm‰t muuttujat poistettu, kommentit suomeksi
'
' Kuvaus:
'   Apulomake, joka linkitt‰‰ ulkoiset taulut uudelleen nykyisen tietokannan kansioon:
'   - Hakee linkitetyt taulut MSysObjects-taulusta
'   - Poimii tietokannan tiedostonimen nykyisest‰ linkkipolusta
'   - Linkitt‰‰ taulun saman nimiseen tiedostoon nykyisess‰ kansiossa
'   - Est‰‰ linkkien j‰‰misen vanhaan kansioon tietokantaa siirrett‰ess‰
'   - MotorData-taulu j‰tet‰‰n linkitt‰m‰tt‰
'
' Riippuvuudet:
'   - MSysObjects-j‰rjestelm‰taulu (Accessin sis‰inen)
'   - Linkitetyt Access-tietokantataulut
'
' K‰yttˆtapaus:
'   Kun tietokanta siirret‰‰n uuteen kansioon, linkitetyt taulut osoittavat yh‰
'   vanhaan sijaintiin. Lomake p‰ivitt‰‰ kaikki linkit nykyiseen kansioon.
'================================================================================

'--------------------------------------------------------------------------------
' Tapahtuma: Command0_Click
' Tarkoitus: Linkitt‰‰ kaikki ulkoiset taulut uudelleen nykyisen tietokannan kansioon
'
' Prosessi:
'   1. Haetaan nykyisen tietokannan hakemistopolku
'   2. Haetaan MSysObjects-taulusta kaikki linkitetyt taulut (paitsi MotorData)
'   3. Jokaiselle taululle:
'      - Poimitaan tiedostonimi nykyisest‰ linkkipolusta
'      - Verrataan hakemistopolkuja
'      - Jos polut eroavat, linkki p‰ivitet‰‰n nykyiseen kansioon
'   4. Taulun nimi ja rakenne s‰ilyv‰t
'
' Huomiot:
'   - Linkki p‰ivitet‰‰n vain, jos l‰hdehakemisto eroaa nykyisest‰
'   - Polkujen vertailu ei erottele isoja ja pieni‰ kirjaimia
'   - MotorData-taulu j‰tet‰‰n pois (vaatii erillisen k‰sittelyn)
'   - TableDef.RefreshLink p‰ivitt‰‰ linkin pudottamatta taulua
'--------------------------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo ErrorHandler
    Dim i As Integer  ' Merkkilaskuri polun j‰sent‰misess‰
    Dim Polku As String  ' Nykyisen tietokannan hakemistopolku
    Dim Polku2 As String  ' Linkitetyn taulun hakemistopolku
    Dim Nimi As String  ' Tietokantatiedoston nimi
    Dim Taul As DAO.Recordset  ' MSysObjects-kyselyn tulokset
    Dim Taulu As String  ' K‰sitelt‰v‰n taulun nimi
    Dim tdf As DAO.TableDef  ' Linkitetyn taulun m‰‰ritys p‰ivityst‰ varten
    Dim db As DAO.Database  ' Nykyinen tietokanta
    Dim Kohdetiedosto As String  ' Uusi linkitett‰v‰ tiedostopolku
    Dim virhenumero As Long  ' Virheen numero k‰sittelij‰ss‰
    Dim virhekuvaus As String  ' Virheen kuvaus k‰sittelij‰ss‰

    ' Haetaan nykyisen tietokannan hakemistopolku
    Set db = CurrentDb
    Polku = Left$(db.Name, Len(db.Name) - Len(Dir$(db.Name)))
    
    ' Haetaan kaikki linkitetyt taulut (paitsi MotorData)
    Set Taul = db.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like '" & POISSULJETTU_TAULU & "') AND ([Type] = " & LINKITETTY_TAULU_TYYPPI & "))", dbOpenSnapshot, dbForwardOnly)
    
    ' K‰sitell‰‰n jokainen linkitetty taulu
    Do While Not Taul.EOF
        ' Estet‰‰n saman taulun k‰sittely useaan kertaan
        If Taul.Fields("Name") <> Taulu Then
            ' Poimitaan hakemisto ja tiedostonimi nykyisest‰ linkkipolusta
            Nimi = "": Polku2 = ""  ' Nollataan, jottei edellisen taulun polku j‰‰ voimaan
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
               Polku2 = LCase$(Left$(Taul.Fields("Database"), i))  ' Hakemistopolku
                    Nimi = Mid$(Taul.Fields("Database"), i + 1)  ' Tiedostonimi
                    Exit For
                End If
            Next i
            
            ' Tarkistetaan, eroaako linkin polku nykyisen tietokannan polusta
            If Polku2 <> LCase$(Polku) Then
                Taulu = Taul.Fields("Name")
                
                Kohdetiedosto = Polku & Nimi
                If Len(Dir$(Kohdetiedosto)) = 0 Then
                    MsgBox "Tietokantatiedostoa ei lˆydy:" & vbCrLf & Kohdetiedosto, vbExclamation, "Linkitysvirhe"
                Else
                    ' P‰ivitet‰‰n linkki rikkomatta taulua tai sen relaatioita.
                    Set tdf = db.TableDefs(Taulu)
                    tdf.Connect = ";DATABASE=" & Kohdetiedosto
                    On Error Resume Next
                    tdf.RefreshLink
                    If Err.Number <> 0 Then
                        MsgBox "Virhe p‰ivitett‰ess‰ taulua " & tdf.Name & ": " & Err.Description, vbExclamation, "Linkin p‰ivitysvirhe"
                    End If
                    Err.Clear
                    On Error GoTo ErrorHandler
                    Set tdf = Nothing
                End If
            End If
        End If
        Taul.MoveNext
    Loop
    
    ' Siivotaan resurssit
    Taul.Close
    Set Taul = Nothing
    Set db = Nothing
    MsgBox "Valmis", vbInformation, "Linkkien p‰ivitys"
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
    MsgBox "Virhe tauluyhteyksien p‰ivityksess‰ (" & virhenumero & "): " & virhekuvaus, vbCritical
End Sub
