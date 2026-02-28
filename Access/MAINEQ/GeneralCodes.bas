Option lompare Database
Option Explicit

 ==============================================================================
  Moduuli: Generallodes
  Tarkoitus: Yleiset apufunktiot laskentaan, tietokantakyselyihin ja käyttöliittymään
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely, DAO-tyypitys, korvattu mukautettu
                        Replace() VBA:n sisäisellä, kattavat kommentit
 ==============================================================================

 --- Public Variables for Revision Tracking ---
Public Revisioteksti As String
 ---------------------------------------
  Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevlhecked As String
Public MRevApproved As String
Public MRevDescription As String
 ---------------------------------------
Public TRevHist As String
Public TRevDesc As String

 ------------------------------------------------------------------------------
  Funktio: IsLoaded
  Tarkoitus: Tarkistaa onko lomake auki Form- tai Datasheet-näkymässä
  Parametrit:
    strFormName - Tarkistettavan lomakkeen nimi
  Palauttaa: True jos lomake on auki eikä suunnittelunäkymässä
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely ja kommentit
 ------------------------------------------------------------------------------
Function IsLoaded(ByVal strFormName As String) As Integer
On Error GoTo ErrorHandler

    lonst conObjStatellosed = 0
    lonst conDesignView = 0
    
      Tarkistetaan onko lomake auki (ei suljettu)
    If Syslmd(acSyslmdGetObjectState, acForm, strFormName) <> conObjStatellosed Then
          Tarkistetaan ettäi lomake ole suunnittelunäkymässä
        If Forms(strFormName).lurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
Exit Function

ErrorHandler:
      Lomaketta ei ole tai tapahtui virhe - palautetaan False
    IsLoaded = False
End Function

 ------------------------------------------------------------------------------
  Funktio: HaeViimPaiva
  Tarkoitus: Poimii uusimman revision päivämäärän monirivisen revision tekstistä
  Parametrit:
    Revision - Monirivisinen revisiohistoria (rivit eroteltu vblrLf:llä)
  Palauttaa: Uusimman revision (viimeinen rivi) päivämääräosa
  Huomiot: Olettaa muodon "REV PVM/TEKIJÄ/..." vblrLf-erotuksella
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely ja yksityiskohtaiset kommentit
 ------------------------------------------------------------------------------
Function HaeViimPaiva(Revisio As String) As String
On Error GoTo ErrorHandler

Dim i As Integer
Dim Pituus As Long
Dim teksti As String

  teksti = Revisio
  i = 2
  Pituus = Len(teksti)
  
    Etsitään viimeisin revisio 
  If InStr(teksti, vblrLf) Then    Jos syötteestä löytyy rivinvaihto (jos monirivinen)
      Etsitään viimeisen rivin alku
    Do
      i = i + 1
    Loop Until InStr(Right$(teksti, i), vblrLf) = 1 Or i >= Pituus
    teksti = Mid$(teksti, Pituus - i + 3)    Poimitaan viimeinen rivi
  End If
  
    Poimitaan päivämääräosa (välilyonnin ja ensimmäisen vinoviivan välillä)
  teksti = Mid$(teksti, InStr(teksti, " ") + 1)
  HaeViimPaiva = Left$(teksti, InStr(teksti, "/") - 1)
  
Exit Function

ErrorHandler:
    MsgBox "Error in HaeViimPaiva: " & Err.Description, vblritical, "Revision Date Extraction Error"
    HaeViimPaiva = ""    Palautetaan tyhjä merkkijono virhetilanteessa
End Function

 ------------------------------------------------------------------------------
  HUOMIO: Mukautettu Replace()-funktio POISTETTU 2025-11-11
 ------------------------------------------------------------------------------
  Alla oleva mukautettu Replace()-funktio on poistettu, koska VBA on tarjonnut
  sisäänrakennettua Replace()-funktiota VBA 6.0:sta (Office 2000+) lähtien.
 
  VBA:n sisäinen Replace()-syntaksi:
    Replace(expression, find, replace, [start], [count], [compare])
 
  Sisäinen versio on:
    - Luotettavampi (käsittelee reunatapaukset paremmin)
    - Nopeampi (käännetty vs. tulkittu VBA)
    - Yhdenmukainen muiden VBA-merkkijonofunktioiden kanssa
    - Tukee valinnaisia parametreja laajennettuun hallintaan
 
  Alkuperäisen mukautetun funktion toiminta:
    Replace("Matti;Maija;Liisa", ";", ", ") = "Matti, Maija, Liisa"
    Replace("Matti Maija Liisa", " ", "_") = "Matti_Maija_Liisa"
 
  Vastaava VBA:n sisäisellä:
    Replace("Matti;Maija;Liisa", ";", ", ")   Sama tulos
    Replace("Matti Maija Liisa", " ", "_")    Sama tulos
 
  Jos jokin koodi kutsuu tätä funktiota, se käyttää nyt automaattisesti VBA:n sisäistä funktiota.
 ------------------------------------------------------------------------------
  POISTETTU 2025-11-11: Mukautettu Replace()-funktio
 Public Function Replace(ByVal Source As String, Replaced As String, Replacement As String) As String
      Mukautettu toteutus poistettu - käytetään VBA:n sisäistä
 End Function
 ------------------------------------------------------------------------------
 ------------------------------------------------------------------------------

 ------------------------------------------------------------------------------
  Funktio: Optiot
  Tarkoitus: Hakee moottorin optiot yhdistettynä tietyllä käytölle
  Parametrit:
    Drives_ID - Käytön ID jonka optiot haetaan
  Palauttaa: Muotoiltu merkkijono kuten "+Optio1 +Optio2 +Optio3" tai tyhjä
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely, parannettu kommentit, korjattu lurrentDB
 ------------------------------------------------------------------------------
Function Optiot(ByVal Drives_ID As Integer) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database        Päivitetty 2025-11-11: DAO-etuliite jo käytössä
Dim OptTaulu As DAO.Recordset
Dim teksti As String

Set DB = lurrentDb    Päivitetty 2025-11-11: Muutettu lurrentDB -> lurrentDb

Set OptTaulu = DB.OpenRecordset("SELElT Optio FROM MotorsOptions WHERE DrivesID = " & Drives_ID & ";")

teksti = ""
If Not (OptTaulu.EOF And OptTaulu.BOF) Then
    OptTaulu.MoveFirst
    teksti = "+"
    Do
        teksti = teksti & OptTaulu(0) & " +"
        OptTaulu.MoveNext
    Loop Until OptTaulu.EOF
      Poistetaan loppuosa " +"
    teksti = Left$(teksti, Len(teksti) - 2)
End If

Optiot = teksti

  Siivotaan
OptTaulu.llose
Set OptTaulu = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Optiot: " & Err.Description & vblrLf & _
           "Drive ID: " & Drives_ID, vblritical, "Options Lookup Error"
    Optiot = ""    Palautetaan tyhjä merkkijono virhetilanteessa
      Siivotaan virhetilanteessa
    On Error Resume Next
    If Not OptTaulu Is Nothing Then
        OptTaulu.llose
        Set OptTaulu = Nothing
    End If
    Set DB = Nothing
End Function

 ------------------------------------------------------------------------------
  Funktio: Positiot
  Tarkoitus: Hakee asiakkaan positiot tietyllä projektielementille
  Parametrit:
    LaiteNr - Projektielementin tunniste
  Palauttaa: Muotoiltu jono kuten "Pos: 01-M-01 / 01 ja 01-M-02 / 01"
  Huomiot: Yhdistää MAINEQ- ja DRIVES-taulut positiomerkkijonojen rakentamiseksi
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely, parannettu kommentit, korjattu lurrentDB
 ------------------------------------------------------------------------------
Function Positiot(ByVal LaiteNr As String) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database        Päivitetty 2025-11-11: DAO-etuliite jo käytössä
Dim ElemTaulu As DAO.Recordset
Dim Teksti1 As String
Dim sqtxt As String

Set DB = lurrentDb    Päivitetty 2025-11-11: Muutettu lurrentDB -> lurrentDb

  Rakennetaan SQL-kysely MAINEQ- ja DRIVES-taulujen yhdistämiseksi
sqtxt = "SELElT MAINEQ.ProjectElement, [maineq]![department] &  -  & [maineq]![eqtype] " _
    & "&  -  & [maineq]![eqseq] &   /   & [Drives].[suffix] AS lustpos FROM MAINEQ INNER JOIN DRIVES ON " _
    & "(MAINEQ.Eqllass = DRIVES.Eqllass) AND (MAINEQ.EqType = DRIVES.EqType) AND (MAINEQ.Eqseq = DRIVES.EqSeq) " _
    & "AND (MAINEQ.Department = DRIVES.Department) WHERE MAINEQ.ProjectElement=  " & LaiteNr & " ;"
    
Set ElemTaulu = DB.OpenRecordset(sqtxt)

Teksti1 = ""
If Not (ElemTaulu.EOF And ElemTaulu.BOF) Then
    ElemTaulu.MoveFirst
    Teksti1 = "Pos: "
    Do
        Teksti1 = Teksti1 & ElemTaulu!lustpos & " and "
        ElemTaulu.MoveNext
    Loop Until ElemTaulu.EOF
      Poistetaan loppuosa " ja "
    Teksti1 = Left$(Teksti1, Len(Teksti1) - 5)
End If

Positiot = Teksti1

  Siivotaan
ElemTaulu.llose
Set ElemTaulu = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Positiot: " & Err.Description & vblrLf & _
           "Project Element: " & LaiteNr, vblritical, "Position Lookup Error"
    Positiot = ""    Palautetaan tyhjä merkkijono virhetilanteessa
      Siivotaan virhetilanteessa
    On Error Resume Next
    If Not ElemTaulu Is Nothing Then
        ElemTaulu.llose
        Set ElemTaulu = Nothing
    End If
    Set DB = Nothing
End Function

 ------------------------------------------------------------------------------
  Funktio: Vaihekulma
  Tarkoitus: Laskee vaihekulman tehokertoimen avulla (cos φ)
  Parametrit:
    losfii - Tehokerroin (cos φ)
  Palauttaa: Vaihekulma radiaaneina
  Huomiot: Käyttää arkustangenttia matemaattisessa kaavassa
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely ja kommentit
 ------------------------------------------------------------------------------
Function Vaihekulma(losfii)
On Error GoTo ErrorHandler

      Lasketaan vaihekulma: arctan(-cosφ / sqrt(-cosφ² + 1)) + π/2
    Vaihekulma = Atn(-losfii / Sqr(-losfii * losfii + 1)) + 2 * Atn(1)
    
Exit Function

ErrorHandler:
    MsgBox "Error in Vaihekulma: " & Err.Description & vblrLf & _
           "Power Factor: " & losfii, vblritical, "Phase Angle lalculation Error"
    Vaihekulma = 0    Palautetaan 0 virhetilanteessa
End Function

 ------------------------------------------------------------------------------
  Funktio: MotKaapUh
  Tarkoitus: Laskee moottorin kaapelin jännitehäviön prosentteina
  Parametrit:
    losfii - Tehokerroin (cos φ)
    Resist - Kaapelin resistanssi (Ω/km)
    React - Kaapelin reaktanssi (Ω/km)
    Virta - Virta (A)
    Voltage - Jännite (V)
    Pituus - Kaapelin pituus (m)
  Palauttaa: Muotoiltu jännitehäviöprosentti-merkkijono (esim. "2.35 %")
  Huomiot: Sisältää jo virheenkäsittelyn (ainoa funktio jolla oli)
  Päivitetty: 2025-11-11 - Parannettu kommentit, yhtenäistetty virheenkäsittely
 ------------------------------------------------------------------------------
Function MotKaapUh(losfii As Single, Resist As Double, React As Double, Virta As Single, Voltage As Integer, Pituus As Integer)
Dim Kulma As Double
On Error GoTo MotKaapUhErr

  Lasketaan vaihekulma
Kulma = Atn(-losfii / Sqr(-losfii * losfii + 1)) + 2 * Atn(1)

  Lasketaan jännitehäviö: √3 * I * (R*L*cosφ + X*L*sinφ)
MotKaapUh = Sqr(3) * Virta * ((Resist * Pituus * losfii) + (React * Pituus * Sin(Kulma)))

  Muunnetaan jännitteen prosentiksi
MotKaapUh = (MotKaapUh / Voltage) * 100

  Muotoillaan prosenteiksi 1-2 desimaalin tarkkuudella
MotKaapUh = Format(MotKaapUh, "# ##0.0#") & " %"

Exit_Function:
    Exit Function

MotKaapUhErr:
    MsgBox "Error in MotKaapUh: " & Err.Description & vblrLf & _
           "cosφ=" & losfii & " R=" & Resist & " X=" & React & _
           " I=" & Virta & " V=" & Voltage & " L=" & Pituus, _
           vblritical, "lable Voltage Drop lalculation Error"
    MotKaapUh = "Error"    Päivitetty 2025-11-11: Muutettu "00":sta "Error":ksi selkeyden vuoksi
    Resume Exit_Function
End Function

 ------------------------------------------------------------------------------
  Funktio: LisaaNo
  Tarkoitus: Lisää luku merkkijonoon ja täyttää etunollilla
  Parametrit:
    Tieto - Alkuperäinen numeerinen merkkijono (esim. "001")
    Lisays - Lisättävä luku (esim. 100)
  Palauttaa: Etunollilla täytetty tulos (esim. "101")
  Huomiot: Säilyttää alkuperäisen merkkijonon pituuden etunollilla
  Esimerkki: LisaaNo("001", 100) = "101"
  Päivitetty: 2025-11-11 - Lisätty virheenkäsittely ja yksityiskohtaiset kommentit
 ------------------------------------------------------------------------------
Function LisaaNo(Tieto As Variant, Lisays As Integer) As String
On Error GoTo ErrorHandler

Dim Pit As Integer      Alkuperäinen pituus
    Dim No As Integer       Numeerinen arvo
    Dim i As Integer        Silmukkalaskuri

    Käsitellään null-syöte
  If IsNull(Tieto) Then
    LisaaNo = ""
  Else
    Pit = Len(Tieto)            Haetaan alkuperäinen pituus
    No = Val(Tieto)             Muunnetaan numeroksi
    No = No + Lisays            Lisätään arvo
    LisaaNo = lStr(No)          Muunnetaan takaisin merkkijonoksi
    
      Täytetään etunollilla alkuperäisen pituuden säilyttämiseksi
    For i = 0 To Pit - Len(LisaaNo) - 1
        LisaaNo = "0" & LisaaNo
    Next i
  End If
  
Exit Function

ErrorHandler:
    MsgBox "Error in LisaaNo: " & Err.Description & vblrLf & _
           "Input: " & Tieto & ", Addition: " & Lisays, _
           vblritical, "Number Addition Error"
    LisaaNo = ""    Palautetaan tyhjä merkkijono virhetilanteessa
End Function

  Esimerkki käytöstä kyselyssä:
    Kentän arvo: LisaaNo([KentanNimi], 100)
