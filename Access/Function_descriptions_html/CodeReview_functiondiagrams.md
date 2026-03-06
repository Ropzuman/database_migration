🕵️‍♂️ Code Review -raportti: Piirikohtaiset toimintakuvaukset ja apumoduulit
Moduuli 1: KAANNOS.bas (Käännöslogiikka)
📊 Yhteenveto

Tämä moduuli vastaa laite- ja piiriviittausten (esim. {xx-xx-xx}) kääntämisestä selkokielisiksi nimiksi tietokannasta. Logiikka on ratkaisevassa osassa toimintakuvausten generoinnissa. Moduulissa on kuitenkin kriittinen tietoturvaan ja vakauteen liittyvä puute tietokantakyselyiden muodostamisessa.
🚨 Kriittiset löydökset (Tietoturva & Vakaus)

    SQL-injektion riski ja kyselyiden kaatuminen (DLookup): DLookup-funktiossa käytetään merkkijonojen yhdistämistä (string concatenation) ehtolausekkeen rakentamiseen: "...[AreaCode] = '" & Osat(0) & "'...". Jos yksikään Osat-taulukon alkio sisältää heittomerkin ('), Accessin SQL-moottori kaatuu syntaksivirheeseen. Vaikka Accessissa täysimittainen SQL-injektio on rajatumpi, tämä on silti tietoturva- ja vakausriski.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

    DLookup on Accessissa erittäin hidas funktio, varsinkin jos sitä kutsutaan toistuvasti isojen silmukoiden sisällä toimintokuvauksia generoitaessa (N+1 -ongelma).

💡 Parannusehdotukset (Ylläpidettävyys)

    Suorituskyvyn parantamiseksi toistuvat kääntämiset tulisi ideaalitilanteessa tehdä yhdellä kootulla Recordset-kyselyllä (DAO.Recordset), tai ainakin heittomerkit on puhdistettava parametreista ennen kyselyä.

🛠️ Korjattu koodi (Heittomerkkien esikäsittely)

Koska Accessin DLookup ei tue parametrisoituja kyselyitä suoraan, minimikorjaus on tuplata heittomerkit Replace-funktiolla ennen kyselyä.
VB.Net

' Alkuperäinen haavoittuva koodi:
' Poistettu = DLookup("[DELETED]", "Loops", "[AreaCode] = '" & Osat(0) & "' AND [LoopSymb] = '" & Osat(1) & "' AND [LoopNo] = '" & Osat(2) & "'")

' Korjattu turvallisempi versio:
Dim strAreaCode As String, strLoopSymb As String, strLoopNo As String

' Puhdistetaan syötteet (korvataan ' -> '')
strAreaCode = Replace(Nz(Osat(0), ""), "'", "''")
strLoopSymb = Replace(Nz(Osat(1), ""), "'", "''")
strLoopNo = Replace(Nz(Osat(2), ""), "'", "''")

Poistettu = DLookup("[DELETED]", "Loops", _
    "[AreaCode] = '" & strAreaCode & "' AND " &_
    "[LoopSymb] = '" & strLoopSymb & "' AND " & _
    "[LoopNo] = '" & strLoopNo & "'")

Moduuli 2: Form_PIIRIT subform.cls ja Form_MOTORS subform.cls
📊 Yhteenveto

Nämä alilomakkeet vastaavat tekstien (tagien) lisäämisestä päälomakkeen (FrmMUOKKAUS) tekstikenttään käyttäjän kursorin kohdalle. Koodi hoitaa tehtävänsä, mutta on vahvasti riippuvainen globaaleista tilamuuttujista, mikä tekee siitä herkän sivuvaikutuksille.
🚨 Kriittiset löydökset (Tietoturva & Vakaus)

    Ei välittömiä tietoturva-aukkoja. Koodi operoi käyttöliittymän tasolla.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

    Globaali Kursori-muuttuja (GeneralCodes.bas): Globaalin muuttujan käyttö kursorin paikan tallentamiseen on riski. Jos käyttäjä avaa useita ikkunoita tai siirtää fokuksen yllättäen toisaalle, kursorin arvo voi olla epäsynkassa aktiivisen kentän kanssa, johtaen tekstin lisäykseen täysin väärään paikkaan.

    Kätketyt virheet: On Error Resume Next on kääritty kursorin positioinnin ympärille (virhe 2185:n estämiseksi), mutta se voi peittää alleen oikeita loogisia virheitä, jos SetFocus epäonnistuu.

💡 Parannusehdotukset (Ylläpidettävyys)

    Refaktoroi kursorin paikan haku hyödyntämään päälomakkeen omaa tilaa tai palauttamaan arvo funktiosta, jotta eroon päästään hauraasta Public Kursori As Long -muuttujasta.

🛠️ Korjattu koodi (Turvallisempi tekstin lisäys)

Tarkistetaan lisäksi, ettei Kursori ole kentän pituutta suurempi, mikä estää Left$ ja Mid$ -funktioiden mahdolliset virheet.
VB.Net

Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    Dim KohdeTeksti As String
    Dim TurvallinenKursori As Long

    On Error GoTo ErrorHandler
    
    If Not KohdeTextBox Is Nothing Then
        KohdeTeksti = Nz(KohdeTextBox.Value, "")
        
        ' Varmistetaan, ettei globaali kursori osoita tekstin ulkopuolelle
        TurvallinenKursori = Kursori
        If TurvallinenKursori > Len(KohdeTeksti) Then TurvallinenKursori = Len(KohdeTeksti)
        If TurvallinenKursori < 0 Then TurvallinenKursori = 0
        
        If Form.Parent.CLoppuun = True Then
            Alku = KohdeTeksti
            Loppu = ""
        Else
            Alku = Left$(KohdeTeksti, TurvallinenKursori)
            Loppu = Mid$(KohdeTeksti, TurvallinenKursori + 1)
        End If
        
        KohdeTextBox.Value = Alku & " {" & Me.TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        
        On Error Resume Next ' Vain SelStart-operaatiota varten
        KohdeTextBox.SelStart = TurvallinenKursori + Len(Me.TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        On Error GoTo ErrorHandler
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe tekstin lisäyksessä: " & Err.Description, vbExclamation, "Virhe"
End Sub

Moduuli 3: USysCheck.bas (Käyttäjäseuranta)
📊 Yhteenveto

Moduuli käyttää Windows API -kutsuja selvittääkseen käyttäjän sekä tietokoneen nimen, ja kirjaa ne UsysUsers-tauluun. 64-bittiset API-deklaraatiot ovat kunnossa muutoslokin perusteella.
🚨 Kriittiset löydökset (Tietoturva & Vakaus)

    Muistivuoto ja lukkiutumisriski (Recordset): Vaikka ErrorHandler:ssa tehdään .Close, onnistuneessa suorituksessa Recordset (Taulu) jätetään sulkematta ennen aliohjelman päättymistä. Tämä jättää lukkoja tietokantaan ja voi ajan myötä johtaa tietokannan paisumiseen (database bloat) tai "Max Locks Exceeded" -kaatumisiin.

💡 Parannusehdotukset (Ylläpidettävyys)

    Kenttiin viittaus tulisi tehdä nimillä (.Fields("UserName")) indeksien (.Fields(0)) sijaan. Indeksien käyttö rikkoutuu välittömästi, jos tietokannan taulun kenttien järjestystä joskus muutetaan.

🛠️ Korjattu koodi (Resurssien hallinta ja nimetyt kentät)
VB.Net

    ' ... API-haut ...
    
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    
    With Taulu
        .AddNew
        .Fields("Verkkotunnus") = NWUserName  ' Käytä oikeaa kentän nimeä 0, 1, 2 sijaan
        .Fields("AccessTunnus") = CurrentUser()
        .Fields("KoneenNimi") = CName
        .Fields("Aikaleima") = Now
        .Update
    End With
    
    ' KRIITTINEN: Vapautetaan resurssit normaalin suorituksen lopuksi
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    
    Exit Function ' tai Sub

ErrorHandler:
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    MsgBox "Virhe kirjauksessa: " & Err.Description

Moduuli 4: Form_FrmASETUKSET.cls (Taulujen linkitys)
📊 Yhteenveto

Lomake päivittää ulkoisten tietokantojen linkitykset. Logiikka pudottaa vanhan taulun ja luo uuden linkitetyn TableDef-määrityksen.
⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

    Accessin instanssointi (CurrentDb): Koodissa kutsutaan CurrentDb toistuvasti (CurrentDb.TableDefs.Count, CurrentDb.TableDefs(i), CurrentDb.Execute, jne.). Accessissa CurrentDb luo joka kerta uuden ilmentymän tietokantaobjektista. Tämä on paitsi suorituskykysyöppö, se voi aiheuttaa tilanteen, jossa .Refresh-komento ei päivity samalle instanssille kunnolla.

🛠️ Korjattu koodi (Välimuistitetun Database-objektin käyttö)
VB.Net

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim i As Integer
    
    Set db = CurrentDb ' Alustetaan VAIN kerran
    
    For i = 0 To db.TableDefs.Count - 1
        If LCase$(db.TableDefs(i).Name) = LCase$(Taulukko) Then
            db.Execute "DROP TABLE [" & Taulukko & "]"
            Exit For
        End If
    Next i
    
    ' Luodaan linkitetty taulumääritys
    Set tdf = db.CreateTableDef(Taulukko)
    With tdf
        .Connect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
    
    Set tdf = Nothing
    Set db = Nothing

Moduuli 5: Form_DBUsers.cls (.LACCDB luku)
📊 Yhteenveto

Lukee binäärisesti Accessin oman .LACCDB-lukitustiedoston näyttääkseen aktiiviset käyttäjät.
🚨 Kriittiset löydökset (Tietoturva & Vakaus)

    Mahdollinen loputon silmukka (Infinite loop): Binäärisen tiedoston luku Do While Not EOF(iLDBFile) voi joissain korruptoituneen lukitustiedoston tapauksissa jäädä jumiin Access/VBA -ympäristöissä, jos tiedostoa muokataan samanaikaisesti.

    Tiedostokahvan sulku: Varmista, että Close #iLDBFile löytyy suorituspolun ja ErrorHandler:in lopusta, muuten tiedosto jää lukituksi (File in use -virhe) ohjelman sulkemiseen asti.

💡 Parannusehdotukset

Tarkista puskurien pituudet. Null-terminoitujen (Chr(0)) merkkijonojen purkaminen manuaalisesti While-silmukalla on hidasta, mutta toimivaa. Nykyaikaisempi ja luotettavampi tapa tutkia tietokannan käyttäjiä on käyttää ADO:n OpenSchema(adSchemaProviderSpecific, , "{947bb102-5d43-11d1-bdbf-00c04fb92675}") -kytkintä suoran tiedostoluvun sijaan, joka on Microsoftin virallisesti tukema ratkaisu tähän!
