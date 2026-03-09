1. Moduuli: Form_DBUsers.cls

    📊 Yhteenveto: Lukitustiedoston (.LACCDB) lukulogiikkaa on parannettu huomattavasti. Kahvojen vapautus toimii nyt turvallisesti.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus):

        VBA:n puuttuva oikosulkuarviointi (Short-circuit evaluation): Koodissa käytetään ehtoa While .bMach(i) <> 0 And i <= 32. Toisin kuin monet modernit kielet (kuten C# tai Python), VBA arvioi aina And-operaattorin molemmat puolet. Jos muuttuja i kasvaa arvoon 33, ohjelma yrittää lukea arvoa .bMach(33), mikä aiheuttaa välittömän "Subscript out of range" (Error 9) -kaatumisen, koska taulukko on määritelty 1 To 32.

    💡 Parannusehdotukset (Ylläpidettävyys): Rakenna ehtolause siten, että taulukon rajoja testataan ensin, ja vältä And-operaattoria yhdistettynä taulukon indeksiin.

    🛠️ Korjattu koodi:
    VBA

    ' KORJATTU: Turvallinen taulukon läpikäynti ilman Index Out of Bounds -riskiä
    i = 1
    sMach = ""
    Do While i <= 32
        If rUser.bMach(i) = 0 Then Exit Do ' Nollatavu katkaisee lukemisen
        sMach = sMach & Chr$(rUser.bMach(i))
        i = i + 1
    Loop

2. Moduuli: Form_TYÖKALUT.cls

    📊 Yhteenveto: Massatuontityökalu, joka hakee AutoCADista laajoja määriä blokkeja valintajoukkojen (SelectionSet) avulla.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus):

        Valintajoukkojen nimitörmäykset: Koodi luo uuden valintajoukon komennolla Set Joukko = oDOC.SelectionSets.Add("APUPICK"). Jos aiempi koodin ajo on kaatunut tai keskeytetty ennen joukon tuhoamista, "APUPICK"-niminen valintajoukko on yhä olemassa AutoCADin muistissa. Tällöin .Add-metodi heittää virheen ja koko skripti pysähtyy.

    💡 Parannusehdotukset (Ylläpidettävyys): Poista nimetty valintajoukko aina On Error Resume Next -rakenteella ennen uuden luomista. Tämä on vakiintunut ja turvallisin tapa käsitellä AutoCADin valintajoukkoja VBA:ssa.

    🛠️ Korjattu koodi:
    VBA

    ' KORJATTU: Varmistetaan, ettei vanha haamujoukko kaada ohjelmaa
    On Error Resume Next
    oDOC.SelectionSets.Item("APUPICK").Delete
    On Error GoTo 0 ' Palautetaan normaali virheenkäsittely

    Set Joukko = oDOC.SelectionSets.Add("APUPICK")
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
