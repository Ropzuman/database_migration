Option Compare Database
Option Explicit

'================================================================================
' Moduuli: KAANNOS (Käännös)
' Tarkoitus: Kääntää laite- ja piiriviittaukset kuvaaviksi nimiksi
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — Kommentit suomeksi
'
' Kuvaus:
'   Käsittelee tekstiä, joka sisältää laiteviittauksia muodossa
'   {xx-xx-xx kuvaus}, ja kääntää ne todellisiksi laite- tai piirinimiksi
'   MAINEQ- tai Loops-taulusta. Merkitsee poistetut tai puuttuvat kohteet
'   virhetagein.
'
' Riippuvuudet:
'   - Taulut: MAINEQ, Loops
'   - DLookup-funktio
'
' Esimerkkejä:
'   Syöttö:  "{60-20-01 Moottorin kuvaus}"
'   Tulos:   "60-20-01 Oikea moottorin nimi" (MAINEQ-taulusta)
'   Syöttö:  "{10-TIC-001 Piirin kuvaus}"
'   Tulos:   "10-TIC-001 Oikea piirin kuvaus" (Loops-taulusta)
'================================================================================

'================================================================================
' Funktio: Kaanna
' Tarkoitus: Kääntää laite-/piiriviittaukset todellisiksi nimiksi
' Parametrit:
'   Tieto — Teksti, joka sisältää {Alue-Tyyppi-Nro kuvaus} -muotoisia viittauksia
' Palauttaa: Käännetty teksti tai virheimärkität viittaukset
'
' Kuvaus:
'   Jäsentää tekstistä {POS kuvaus} -muodot, missä:
'   - POS-osa: Alue-Tyyppi-Nro (esim. 60-20-01 tai 10-TIC-001)
'   - Jos Alue = "60": Hakee MAINEQ-taulusta (moottoreiden laitedata)
'   - Muuten: Hakee Loops-taulusta (prosessipiirejä)
'
'   Virheenkäsittely:
'   - [ERR: Not found]: Laitetta ei löydy tietokannasta
'   - [DELETED!]: Laite on merkitty poistetuksi
'   - [ERR: No translation]: Laite löytyy, mutta nimike puuttuu
'================================================================================
Function Kaanna(Tieto As Variant) As Variant
    Dim OS As Long, OS2 As Long, OS3 As Long, OS4 As Long
    Dim tPOS As String
    Dim Osat As Variant
    Dim Nimitys As Variant
    Dim Poistettu As Variant
    Dim Virheet As Long
    Dim sPar0 As String, sPar1 As String, sPar2 As String  ' Sanitoidut DLookup-parametrit (heittomerkki duplikoitu)
    
    On Error GoTo ErrorHandler
    
    If IsNull(Tieto) Then
        Kaanna = Null
        Exit Function
    End If
    
    OS = InStr(Tieto, "{")
    If OS = 0 Then
        ' Ei käännettäviä viittauksia
        Kaanna = Tieto
        Exit Function
    End If
    
    ' Alustetaan tulos tekstillä ennen ensimmäistä viittausta
    Kaanna = Left$(Tieto, OS)
    
    Do While OS > 0
        ' Haetaan sijaintimerkit: { POS } -rakenne
        OS2 = InStr(OS + 1, Tieto, " ")    ' Välilyönti position jälkeen
        OS3 = InStr(OS + 1, Tieto, "}")    ' Sulkeva aaltosulku
        OS4 = InStr(OS3 + 1, Tieto, "{")   ' Seuraava avautuva aaltosulku
        
        ' Poimitaan positiokoodi (esim. "60-20-01" tai "10-TIC-001")
        tPOS = Mid$(Tieto, OS + 1, OS2 - OS - 1)
        Osat = Split(tPOS, "-")
        
        ' Sanitoidaan DLookup-parametrit heittomerkkien varalta — SQL-injektion esto
        sPar0 = Replace(Nz(Osat(0), ""), "'", "''")
        sPar1 = Replace(Nz(Osat(1), ""), "'", "''")
        sPar2 = Replace(Nz(Osat(2), ""), "'", "''")

        ' Valitaan taulu aluekoodin perusteella
        If Osat(0) = "60" Then
            ' Moottori/laite MAINEQ-taulusta
            Nimitys   = DLookup("[EqNameSW20]", "MAINEQ", "[Department] = '" & sPar1 & "' AND [EqSeq] = '" & sPar2 & "'")
            Poistettu = DLookup("[Deleted]",    "MAINEQ", "[Department] = '" & sPar1 & "' AND [EqSeq] = '" & sPar2 & "'")
        Else
            ' Prosessipiiri Loops-taulusta
            Nimitys   = DLookup("[Descr26_P]", "Loops", "[AreaCode] = '" & sPar0 & "' AND [LoopSymb] = '" & sPar1 & "' AND [LoopNo] = '" & sPar2 & "'")
            Poistettu = DLookup("[DELETED]",   "Loops", "[AreaCode] = '" & sPar0 & "' AND [LoopSymb] = '" & sPar1 & "' AND [LoopNo] = '" & sPar2 & "'")
        End If
        
        ' Tarkistetaan virheet ja merkitään tuntemattomiksi tarvittaessa
        If IsNull(Poistettu) Then
            Nimitys = "[ERR: Not found] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        ElseIf Poistettu Then
            Nimitys = "[DELETED!] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        ElseIf IsNull(Nimitys) Then
            Nimitys = "[ERR: No translation] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        End If
        
        ' Rakennetaan tulos: positio + käännös
        Kaanna = Kaanna & tPOS & " " & Nimitys
        
        ' Lisätään teksti tämän viittauksen jälkeen (tai loppuun)
        If OS4 <> 0 Then
            Kaanna = Kaanna & Mid$(Tieto, OS3, OS4 - OS3)
        Else
            Kaanna = Kaanna & Mid$(Tieto, OS3)
        End If
        
        ' Edetään seuraavaan viittaukseen
        OS = InStr(OS + 1, Tieto, "{")
    Loop
    
    ' Kirjataan virheet Immediate-ikkunaan debuggausta varten
    If Virheet > 0 Then
        Debug.Print "Kaanna: " & Virheet & " virhe(ttä) käännöksessä"
    End If
    
    Exit Function
    
ErrorHandler:
    Kaanna = "[ERR: " & Err.Description & "] " & Tieto
    Debug.Print "Kaanna-virhe: " & Err.Description
End Function
