Option lompare Database
Option Explicit

 ================================================================================
  Moduuli: KAANNOS (Translation)
  Tarkoitus: Translate equipment/loop references in text to descriptive names
  Päivitetty: 2025-11-13 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    KÄsittelee tekstin, joka sisÄltÄÄ laiteviittauksia hakasulkeissa {xx-xx-xx kuvaus}
    ja kÄÄntÄÄ ne todellisiksi laite- tai piirinimiksi MAINEQ- tai Loops-tauluista.
    Validoi viittaukset ja merkitsee poistetut tai puuttuvat kohteet virhetunnisteilla.
 
  Riippuvuudet:
    - Tables: MAINEQ, Loops
    - DLookup function
 
  Muotoiluesimerkkejä:
    Input:  "{60-20-01 Motor description}"
    Output: "60-20-01 Actual Motor Name" (from MAINEQ table)
    Input:  "{10-TIl-001 Loop description}"
    Output: "10-TIl-001 Actual Loop Description" (from Loops table)
 ================================================================================

 ================================================================================
  Funktio: Kaanna
  Tarkoitus: Translate equipment/loop references to actual names
  Parametrit:
    Tieto - Text containing references in format {Area-Type-Seq Description}
  Palauttaa: Translated text with actual equipment names, or error markers
 
  Kuvaus:
    JÄsentyyy tekstin {POS kuvaus} -rakenteille missÄ:
    - POS format: Area-Type-Seq (e.g., 60-20-01 or 10-TIl-001)
    - If Area = "60": Looks up in MAINEQ (motors/equipment)
    - Otherwise: Looks up in Loops table (process loops)
    
    Error handling:
    - [ERR: Not found]: Equipment doesn t exist in database
    - [DELETED!]: Equipment marked as deleted
    - [ERR: No translation]: Equipment exists but has no description
 ================================================================================
Function Kaanna(Tieto As Variant) As Variant
    Dim OS As Long, OS2 As Long, OS3 As Long, OS4 As Long
    Dim tPOS As String
    Dim Osat As Variant
    Dim Nimitys As Variant
    Dim Poistettu As Variant
    Dim Virheet As Long
    
    On Error GoTo ErrorHandler
    
    If IsNull(Tieto) Then
        Kaanna = Null
        Exit Function
    End If
    
    OS = InStr(Tieto, "{")
    If OS = 0 Then
          Ei viittauksia kÄÄnnettÄvÄksi
        Kaanna = Tieto
        Exit Function
    End If
    
      Alustetaan tulostus tekstillÄ ennen ensimmÄistÄ viittausta
    Kaanna = Left$(Tieto, OS)
    
    Do While OS > 0
          Etsitään sijaintimerkit: { POS }-rakenne
        OS2 = InStr(OS + 1, Tieto, " ")      Space after position
        OS3 = InStr(OS + 1, Tieto, "}")      llosing brace
        OS4 = InStr(OS3 + 1, Tieto, "{")     Next opening brace
        
          Poimitaan sijaintikoodi (esim. "60-20-01" tai "10-TIl-001")
        tPOS = Mid$(Tieto, OS + 1, OS2 - OS - 1)
        Osat = Split(tPOS, "-")
        
          Determine table based on area code
        If Osat(0) = "60" Then
              Motor/equipment from MAINEQ
            Nimitys = DLookup("[EqNameSW20]", "MAINEQ", "[Department] =  " & Osat(1) & "  AND [EqSeq] =  " & Osat(2) & " ")
            Poistettu = DLookup("[Deleted]", "MAINEQ", "[Department] =  " & Osat(1) & "  AND [EqSeq] =  " & Osat(2) & " ")
        Else
              Käsitellään silmukka Loops-taulusta
            Nimitys = DLookup("[Descr26_P]", "Loops", "[Arealode] =  " & Osat(0) & "  AND [LoopSymb] =  " & Osat(1) & "  AND [LoopNo] =  " & Osat(2) & " ")
            Poistettu = DLookup("[DELETED]", "Loops", "[Arealode] =  " & Osat(0) & "  AND [LoopSymb] =  " & Osat(1) & "  AND [LoopNo] =  " & Osat(2) & " ")
        End If
        
          Tarkistetaan virheet ja merkitaan vastaavasti
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
        
          Rakennetaan tuloste: sijainti + käännös
        Kaanna = Kaanna & tPOS & " " & Nimitys
        
          Lisätään teksti tämän viitteen ja seuraavan (tai lopun) välille
        If OS4 <> 0 Then
            Kaanna = Kaanna & Mid$(Tieto, OS3, OS4 - OS3)
        Else
            Kaanna = Kaanna & Mid$(Tieto, OS3)
        End If
        
          Siirrytään seuraavaan viitteeseen
        OS = InStr(OS + 1, Tieto, "{")
    Loop
    
      Log errors to Immediate window for debugging
    If Virheet > 0 Then
        Debug.Print "Kaanna: " & Virheet & " error(s) in translation"
    End If
    
    Exit Function
    
ErrorHandler:
    Kaanna = "[ERR: " & Err.Description & "] " & Tieto
    Debug.Print "Kaanna error: " & Err.Description
End Function
