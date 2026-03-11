# 🕵️‍♂️ Code Review Report & Implementation Instructions

**To the AI Coding Agent:** This is a senior-level code review report for the VBA modules (`Module1.bas`, `Module2.bas`, `Module3.bas`). Your task is to read these findings and implement the suggested fixes and refactorings into the codebase. Prioritize the critical findings (Security & Stability) first.

---

## 📊 Yhteenveto

Koodikanta on selkeästi rakennettu ja siinä on huomioitu monia tärkeitä suorituskykyä parantavia tekijöitä (kuten "Fast Mode" Excelin UI-päivitysten pysäyttämiseksi ja COM-kutsujen minimointi 2D-taulukoiden sekä `CopyFromRecordset`:n avulla). Siirtyminen 64-bittiseen ympäristöön ja ACE OLEDB -ajureiden käyttö on toteutettu loogisesti.

Koodissa on kuitenkin huomattavia riskejä tietoturvan (SQL-injektio) ja vakauden (virheenhallinta ja resurssivuodot) suhteen, ja koodin ylläpidettävyyttä voidaan parantaa merkittävästi vähentämällä globaalien muuttujien käyttöä ja toisteista koodia.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

1. **Tietoturva - SQL-injektiovaara (CWE-89):**
   - **Ongelma:** `Module1` lukee SQL-kyselyt suoraan Excel-soluista (`sSQL(1) = Sheets("Main").Cells(8 + Valinta, 3).Value`). Jos käyttäjä tai haittaohjelma muokkaa tätä solua, on mahdollista suorittaa vaarallista SQL-koodia (esim. `DROP TABLE`, `DELETE FROM`, tai luvaton datan eksfiltraatio).
   - **Ratkaisu:** Parasta olisi käyttää ADODB:n `Command`-olioita ja parametrisoituja kyselyitä. Jos ratkaisun luonne edellyttää dynaamisia kyselyitä taulukosta, kyselyt tulee ehdottomasti sanitoida ennen ajoa.

2. **Vakaus - `On Error Resume Next` -rakenteen liikakäyttö ja Excelin jäätymisriski:**
   - **Ongelma:** Makro laittaa Excelin tilaan, jossa ruudunpäivitys ja automaattilaskenta on pois päältä (`BeginFastMode`). Vaikka koodissa on `ErrorHandler`, laaja `On Error Resume Next` -käyttö tietokantayhteyksiä avattaessa voi johtaa siihen, että koodi jatkaa suoritusta osittain rikkoutuneessa tilassa. Pahimmassa tapauksessa globaali kaatuminen jättää käyttäjälle Excelin, joka vaikuttaa "jäätyneeltä" (koska `EndFastMode` ei koskaan ajaudu).
   - **Ratkaisu:** Tilaa muuttavissa makroissa on ehdottoman tärkeää käyttää *Try-Finally* -tyyppistä rakennetta VBA:ssa.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

1. **Suorituskyky - Sarakkeiden lukeminen silmukalla:**
   - **Huomio:** `Module2.bas` `HaeDocTiedot`-aliohjelmassa sarakkeita käydään läpi `Do...Loop` -rakenteella sarakkeittain. Vaikka koodissa on turvarajana `MAX_EXCEL_COLUMNS`, on solu kerrallaan iterointi hidasta. Parempi tapa on etsiä viimeinen käytetty sarake `End(xlToLeft)` -metodilla tai lukea koko otsikkorivi yhdellä kerralla 2D-taulukkoon muistiin ja iteroida sitä.

2. **Toiminnallisuus - Rajatarkistukset ja null-arvot:**
   - **Huomio:** Koodissa `Split(DIRev, Chr(10))` suoritetaan riippumatta siitä, onko solusta haettu arvo todellisuudessa tyhjä (Null). Tämä voi aiheuttaa tyyppivirheen (Type Mismatch). On parempi tarkistaa arvon pituus ennen splittausta.

3. **Toiminnallisuus - ADODB vs. DAO -hybridimalli:**
   - Koodi sekoittaa DAO:n (DB1) ja ADODB:n (DB2) saman prosessin sisällä. Vaikka tämä voi olla perusteltua, se kasvattaa muistijalanjälkeä ja tekee ylläpidosta monimutkaisempaa.

## 💡 Parannusehdotukset (Ylläpidettävyys)

1. **Poista julkiset globaalit muuttujat (Spaghetti State):**
   - Kymmenet `Public DI... As String` määrittelyt moduulin 1 alussa rikkovat kapselointiperiaatetta (Encapsulation). Ne altistavat koodin sivuongelmille.
   - **Ratkaisu:** Kokoa nämä yhteen `Type`-rakenteeseen (UDT) tai luokkaan (Class Module), jolloin niitä on helpompi käsitellä, nollata ja välittää argumentteina.

2. **DRY (Don't Repeat Yourself) - Provider Fallback:**
   - `Module1`:ssä ADODB-yhteyden avausyritykset versioilla 16.0, 15.0 ja 12.0 on tehty toisteisella koodilla. Eristä tämä omaksi funktioksi.

## 🛠️ Korjattu koodi (Referenssitoteutukset)

Agent, please use the following patterns when refactoring the code:

### 1. ADODB-yhteyden avaamisen eristäminen (DRY & Vakaus)

Lisää tämä `Module1`:een ja korvaa toisteinen koodi `HaeData`-aliohjelmassa:

```vba
' Hakee toimivan ADODB-yhteyden kokeilemalla saatavilla olevia moottoreita
Private Function LuoADODBYhteys(kantaPolku As String) As Object
    Dim conn As Object
    Dim providerVersions As Variant
    Dim i As Integer
    
    ' Kokeiltavat versiot prioriteettijärjestyksessä (64-bit Office / uudemmat ensin)
    providerVersions = Array("16.0", "15.0", "12.0")
    
    For i = LBound(providerVersions) To UBound(providerVersions)
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB." & providerVersions(i) & ";Data Source=" & kantaPolku
        
        On Error Resume Next
        conn.Open
        If Err.Number = 0 Then
            On Error GoTo 0
            Set LuoADODBYhteys = conn
            Exit Function ' Yhteys onnistui
        End If
        Err.Clear
        On Error GoTo 0
        Set conn = Nothing ' Siivotaan epäonnistunut yritys
    Next i
    
    ' Jos mikään ei onnistunut
    Set LuoADODBYhteys = Nothing
End Function

### 2. Oikeaoppinen virheenkäsittely & Vakaus (Try-Finally -malli VBA:ssa)

Näin varmistat, ettei Excel jää ikuisesti tilaan, jossa ruudunpäivitys on pois päältä. Refaktoroi pääfunktiot tätä mallia noudattaen:

Sub HaeData()
    On Error GoTo ErrorHandler
    BeginFastMode ' Ruudunpäivitys pois
    
    ' ... [Varsinainen koodi, tietokantahaut jne.] ...

SafeExit:
    ' Siivotaan resurssit
    On Error Resume Next
    If Not rsDAO Is Nothing Then rsDAO.Close: Set rsDAO = Nothing
    If Not dbDAO Is Nothing Then dbDAO.Close: Set dbDAO = Nothing
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not conn Is Nothing Then conn.Close: Set conn = Nothing
    On Error GoTo 0
    
    EndFastMode ' RUUDUNPÄIVITYS TAKAISIN (ajetaan aina)
    Exit Sub

ErrorHandler:
    MsgBox "Odottamaton virhe datan haussa: " & Err.Description, vbCritical, "Virhe " & Err.Number
    Resume SafeExit ' Hyppää aina SafeExit-lohkoon, jotta EndFastMode laukeaa
End Sub

### 3. Yksinkertainen SQL-sanitaattori (Tietoturva)

Koska koodi lukee kyselyt taulukosta, estä tuhoisat lausekkeet (DML/DDL). Agent, apply this check before executing dynamic queries.

Private Function OnTurvallinenSQL(ByVal sqlText As String) As Boolean
    Dim uSQL As String
    uSQL = UCase(sqlText)
    
    ' Estetään vaaralliset DML ja DDL komennot. Sallitaan vain SELECT tai tallennetut kyselyt.
    If InStr(uSQL, "DROP ") > 0 Or InStr(uSQL, "DELETE ") > 0 Or _
       InStr(uSQL, "UPDATE ") > 0 Or InStr(uSQL, "INSERT ") > 0 Or _
       InStr(uSQL, "ALTER ") > 0 Or InStr(uSQL, "EXEC ") > 0 Then
        OnTurvallinenSQL = False
    Else
        OnTurvallinenSQL = True
    End If
End Function
