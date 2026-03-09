1. Form_Tee Kuvat.cls

    📊 Yhteenveto: Hieno refaktorointi! Aiemmin havaittu "busy wait" -ongelma (joka piti prosessorin käyttöasteen korkeana) on korjattu käyttämällä Windows API:n Sleep-funktiota (Sleep CLng(Aika) * 10). Lisäksi lomakekontrollien ekspliittinen viittaus (Me.-etuliite) ja DAO:n selkeyttäminen parantavat koodin laatua ja ylläpidettävyyttä. Myös AutoCADin GetObject-fallback säästää valtavasti aikaa massatulostuksissa.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus): Ei havaittu. Koodi on erittäin vakaalla pohjalla.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): Olet lisännyt DoEvents-kutsun fiksusti juuri ennen Sleep-funktiota. Tämä varmistaa, että Accessin käyttöliittymä (esim. peruuta-painike tai tilarivi) ei jäädy odotuksen aikana.

    💡 Parannusehdotukset (Ylläpidettävyys): Varmista aina massatulostusajojen (TeeKuvat_Click) virheenkäsittelijässä (ErrorHandler), että oAcad- ja oDoc-objektit asetetaan arvoon Nothing, vaikkei ohjelmaa haluttaisikaan sulkea (Quit), jotta taustalle ei jää "haamuobjekteja" muistiin.

2. General.bas (Käyttäjäseuranta SniffUser)

    📊 Yhteenveto: Kirjaa käyttäjän verkkonimen (GetUserNameA) ja tietokoneen nimen (GetComputerNameA) UsysUsers-lokiin. API-kutsut on toteutettu fiksusti null-turvallisina (Nz) ja viety 64-bit yhteensopivaan formaattiin.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus): Koodinpätkässä avataan transaktio komennolla DBEngine.BeginTrans ennen recordsetin täyttämistä. Tarkista, että koodin lopussa on ehdottomasti DBEngine.CommitTrans (tai workspace-tason vastaava). Jos CommitTrans jää uupumaan, taulu lukkiutuu ja tietoja katoaa. Samoin virheenkäsittelijässä tulee olla DBEngine.Rollback.

    💡 Parannusehdotukset (Ylläpidettävyys): Käytä transaktioissa mieluummin nimenomaista työtilaa (DBEngine.Workspaces(0)), sillä globaali DBEngine.BeginTrans koskettaa kaikkia avoimia tietokantayhteyksiä kyseisessä instanssissa.

    🛠️ Korjattu koodi (Transaktion varmistus):

VBA

  ' Kirjataan sisäänkirjautuminen tietokantaan tapahtumana
  Set db = CurrentDb
  Dim ws As DAO.Workspace
  Set ws = DBEngine.Workspaces(0)
  
  ws.BeginTrans
  On Error GoTo TransError ' Siirrytään transaktion peruutukseen virheessä
  
  Set Taulu = db.OpenRecordset("UsysUsers", dbOpenDynaset)
  With Taulu
      .AddNew
      .Fields(0) = Nz(NWUserName, "Unknown")
      .Fields(1) = Nz(CurrentUser(), "Unknown")
      .Fields(2) = Nz(CName, "Unknown")
      .Fields(3) = Now
      .Update
  End With
  Taulu.Close
  
  ws.CommitTrans ' HYVÄKSYTÄÄN TRANSAKTIO
  Set Taulu = Nothing
  Set db = Nothing
  Exit Function

TransError:
  ws.Rollback ' PERUTAAN VIRHETILANTEESSA
  MsgBox "Lokituksen tallennus epäonnistui: " & Err.Description

1. USysCheck.bas & For ACAD Utility.bas

    📊 Yhteenveto: Todella oppikirjamainen korjaus! Olet siirtänyt rakenteet (Type OPENFILENAME, Type POINTAPI) ehdollisen käännöksen ulkopuolelle, mikä korjasi aiemman Type Declaration -bugin, joka rikkoi Accessin lomakeviittaukset. API-kutsut on deklatoitu moitteettomasti LongPtr (osoittimet) ja Long (tavalliset numeroarvot) osalta.

    🚨 Kriittiset löydökset: Ei havaittu. Erinomainen 64-bit vakaus.

2. Form_DBUsers.cls

    📊 Yhteenveto: Käyttäjien lukkotiedoston (.ldb / .laccdb) luku suoraan binäärinä. Olet siistinyt null-merkkiin päättyvien (null-terminated) merkkijonojen parsintalogiikkaa huomattavasti.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky): Binääriluku (Do While Not EOF(iLDBFile)) on toimiva ja suhteellisen turvallinen. Varmista aiempien neuvojen mukaisesti, että moduulin ErrHandler-lohkossa on tarvittaessa komento Close iLDBFile (tai pelkkä Close), jotta lukkotiedosto ei jää haamulukkoon verkkoasemalle.

3. Module1.bas

    📊 Yhteenveto: InputBox-pohjainen syötteen validointiluuppi.

    ⚠️ Huomioitavaa (Toiminnallisuus): Nykyisellään Do While continueLoop pitää sisällään useita sisäkkäisiä If-Else -lauseita syötteen numeerisuudelle ja sallitulle välille (1-10). Logiikka on eettisesti oikein (Catch & Retry), mutta silmukan rakennetta voisi "litistää" (flatten) puhtaammaksi, mikä estää nk. "Hadouken"-efektin eli liian syvän sisennyksen.

    💡 Parannusehdotukset (Ylläpidettävyys): Return early / Fail fast -tyyli on usein luettavampi.

    🛠️ Korjattu koodi (Esimerkki siistimmästä silmukasta):

VBA

Sub CustomMessage()
    Dim strInput As String
    Dim n As Long

    Do
        strInput = InputBox("Enter a number between 1 and 10.")
        If strInput = "" Then Exit Sub ' Käyttäjä peruutti
        
        If Not IsNumeric(strInput) Then
            If MsgBox("Please enter a numeric value.", vbOKCancel, "Error!") = vbCancel Then Exit Sub
        Else
            n = CLng(strInput)
            If n >= 1 And n <= 10 Then
                ' Validointi onnistui!
                MsgBox "You entered valid number: " & n, vbInformation
                Exit Do
            Else
                If MsgBox("Number outside range." & vbCrLf & _
                          "You entered a number that is less than 1 or greater than 10.", _
                          vbOKCancel, "Error!") = vbCancel Then Exit Sub
            End If
        End If
    Loop
End Sub
