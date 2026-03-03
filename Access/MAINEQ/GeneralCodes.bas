Option Compare Database
Option Explicit

'==============================================================================
' Moduuli: GeneralCodes
' Tarkoitus: Yleiset apufunktiot laskentaa, tietokantahakuja ja käyttöliittymää varten
'==============================================================================

'--- Julkiset muuttujat revisionseurantaa varten ---
Public Revisioteksti As String
'---------------------------------------
' Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevChecked As String
Public MRevApproved As String
Public MRevDescription As String
'---------------------------------------
Public TRevHist As String
Public TRevDesc As String

'------------------------------------------------------------------------------
' Funktio: SetStartup
' Tarkoitus: Kirjaa käyttäjän kirjautumistiedot UsysUsers-tauluun sovelluksen käynnistyessä
' Parametrit: -
' Palautusarvo: -
' Huom: Kutsutaan AutoExec-makrosta. Virheet käsitellään hiljaisesti,
'       jottei kirjautumistulostus keskeytä sovelluksen avautumista.
'------------------------------------------------------------------------------
Function SetStartup()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database      ' Nykyinen tietokanta
    Dim Taulu As DAO.Recordset  ' UsysUsers-taulu kirjautumistietuetta varten
    Dim NWUserName As String    ' Verkkokäyttäjänimi Windows-API:sta
    Dim CName As String         ' Tietokoneen nimi ympäristömuuttujasta
    Dim BuffSize As Long        ' Puskurin koko API-kutsulle
    Dim NBuffer As String       ' Puskurimerkkijono API-kutsulle

    ' Haetaan verkkokäyttäjänimi Windows-API:lla (wu_GetUserName USysCheck.bas:ssa)
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If wu_GetUserName(NBuffer, BuffSize) Then
        NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        NWUserName = "Unknown"
    End If

    ' Haetaan tietokoneen nimi ympäristömuuttujasta (ei vaadi lisä-API:a)
    CName = Environ("COMPUTERNAME")
    If CName = "" Then CName = "Unknown"

    ' Kirjoitetaan kirjautumistietue UsysUsers-tauluun
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Verkkokäyttäjänimi
        .Fields(1) = CurrentUser()  ' Access-käyttäjänimi
        .Fields(2) = CName          ' Tietokoneen nimi
        .Fields(3) = Now            ' Kirjautumisaika
        .Update
    End With

    ' Siivotaan objektit
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
    ' Hiljainen virheenkäsittely — ei keskeytetä sovelluksen käynnistystä
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Funktio: IsLoaded
' Tarkoitus: Tarkistaa, onko lomake auki lomake- tai taulukkonäkymässä
' Parametrit:
'   strFormName - Tarkistettavan lomakkeen nimi
' Palautusarvo: True, jos lomake on auki eikä suunnittelunäkymässä
'------------------------------------------------------------------------------
Function IsLoaded(ByVal strFormName As String) As Integer
On Error GoTo ErrorHandler

    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    ' Tarkistetaan, onko lomake auki (ei suljettu)
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        ' Tarkistetaan, ettei lomake ole suunnittelunäkymässä
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
Exit Function

ErrorHandler:
    ' Lomaketta ei ole tai tapahtui virhe — palautetaan False
    IsLoaded = False
End Function

'------------------------------------------------------------------------------
' Funktio: HaeViimPaiva
' Tarkoitus: Poimii viimeisimmän revisionpäivämäärän monirivisenä revisiontekstistä
' Parametrit:
'   Revisio - Revisiohistoriamerkkijono (rivit erotettu vbCrLf:llä)
' Palautusarvo: Viimeisimmän revision päivämääräosa
'------------------------------------------------------------------------------
Function HaeViimPaiva(Revisio As String) As String
On Error GoTo ErrorHandler

Dim i As Integer
Dim Pituus As Long
Dim teksti As String

  teksti = Revisio
  i = 2
  Pituus = Len(teksti)
  
  ' Etsitään viimeisin revisio
  If InStr(teksti, vbCrLf) Then  ' Sisyötteessä on rivinvaihto — useampi rivit
    ' Etsitään viimeisen rivin alku
    Do
      i = i + 1
    Loop Until InStr(Right$(teksti, i), vbCrLf) = 1 Or i >= Pituus
    teksti = Mid$(teksti, Pituus - i + 3)  ' Poimitaan viimeinen rivi
  End If
  
  ' Poimitaan päivämääräosa (välilyönnin ja ensimmäisen kauttaviivan väliltä)
  teksti = Mid$(teksti, InStr(teksti, " ") + 1)
  HaeViimPaiva = Left$(teksti, InStr(teksti, "/") - 1)
  
Exit Function

ErrorHandler:
    MsgBox "Error in HaeViimPaiva: " & Err.Description, vbCritical, "Revision Date Extraction Error"
    HaeViimPaiva = ""  ' Virhetilanteessa palautetaan tyhjä merkkijono
End Function

'------------------------------------------------------------------------------
' HUOM: Mukautettu Replace()-funktio POISTETTU
'------------------------------------------------------------------------------
' Poistettu, koska VBA:ssa on ollut sisäänrakennettu Replace()-funktio
' versiosta 6.0 (Office 2000+) lähtien.
'------------------------------------------------------------------------------
' POISTETTU: mukautettu Replace()-funktio

'------------------------------------------------------------------------------
' Funktio: Optiot
' Tarkoitus: Hakee käytön moottorioptiot tietylle käytölle
' Parametrit:
'   Drives_ID - Haettavan käytön ID
' Palautusarvo: Muotoiltu teksti, esim. "+Optio1 +Optio2" tai tyhjä merkkijono
'------------------------------------------------------------------------------
Function Optiot(ByVal Drives_ID As Integer) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim OptTaulu As DAO.Recordset
Dim teksti As String

Set DB = CurrentDb

Set OptTaulu = DB.OpenRecordset("SELECT Optio FROM MotorsOptions WHERE DrivesID = " & Drives_ID & ";")

teksti = ""
If Not (OptTaulu.EOF And OptTaulu.BOF) Then
    OptTaulu.MoveFirst
    teksti = "+"
    Do
        teksti = teksti & OptTaulu(0) & " +"
        OptTaulu.MoveNext
    Loop Until OptTaulu.EOF
    ' Poistetaan loppuosa " +"
    teksti = Left$(teksti, Len(teksti) - 2)
  End If

  Optiot = teksti

' Siivotaan objektit
  OptTaulu.Close
  Set OptTaulu = Nothing
  Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Optiot: " & Err.Description & vbCrLf & _
           "Drive ID: " & Drives_ID, vbCritical, "Options Lookup Error"
    Optiot = ""  ' Virhetilanteessa palautetaan tyhjä merkkijono
    ' Siivotaan objektit virhetilanteessa
    On Error Resume Next
    If Not OptTaulu Is Nothing Then
        OptTaulu.Close
        Set OptTaulu = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Funktio: Positiot
' Tarkoitus: Hakee projektielementin asiakaspositiot
' Parametrit:
'   LaiteNr - Projektielementtitunnus
' Palautusarvo: Muotoiltu teksti, esim. "Pos: 01-M-01 / 01 and 01-M-02 / 01"
'------------------------------------------------------------------------------
Function Positiot(ByVal LaiteNr As String) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim ElemTaulu As DAO.Recordset
Dim Teksti1 As String
Dim sqtxt As String

Set DB = CurrentDb

' Rakennetaan SQL-kysely liittämällä MAINEQ- ja DRIVES-taulukot yhteen
sqtxt = "SELECT MAINEQ.ProjectElement, [maineq]![department] & '-' & [maineq]![eqtype] " _
    & "& '-' & [maineq]![eqseq] & ' / ' & [Drives].[suffix] AS Custpos FROM MAINEQ INNER JOIN DRIVES ON " _
    & "(MAINEQ.EqClass = DRIVES.EqClass) AND (MAINEQ.EqType = DRIVES.EqType) AND (MAINEQ.Eqseq = DRIVES.EqSeq) " _
    & "AND (MAINEQ.Department = DRIVES.Department) WHERE MAINEQ.ProjectElement= '" & LaiteNr & "';"
    
Set ElemTaulu = DB.OpenRecordset(sqtxt)

Teksti1 = ""
If Not (ElemTaulu.EOF And ElemTaulu.BOF) Then
    ElemTaulu.MoveFirst
    Teksti1 = "Pos: "
    Do
        Teksti1 = Teksti1 & ElemTaulu!Custpos & " and "
        ElemTaulu.MoveNext
    Loop Until ElemTaulu.EOF
    ' Poistetaan loppuosa " and "
    Teksti1 = Left$(Teksti1, Len(Teksti1) - 5)
  End If

  Positiot = Teksti1

' Siivotaan objektit
  ElemTaulu.Close
  Set ElemTaulu = Nothing
  Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Positiot: " & Err.Description & vbCrLf & _
           "Project Element: " & LaiteNr, vbCritical, "Position Lookup Error"
    Positiot = ""  ' Virhetilanteessa palautetaan tyhjä merkkijono
    ' Siivotaan objektit virhetilanteessa
    On Error Resume Next
    If Not ElemTaulu Is Nothing Then
        ElemTaulu.Close
        Set ElemTaulu = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Funktio: Vaihekulma
' Tarkoitus: Laskee vaihekulman tehokertoimen (cos φ) perusteella
' Parametrit:
'   Cosfii - Tehokerroin (cos φ)
' Palautusarvo: Vaihekulma radiaaneina
'------------------------------------------------------------------------------
Function Vaihekulma(Cosfii)
On Error GoTo ErrorHandler

    ' Lasketaan vaihekulma: arctan(-cosφ / sqrt(-cosφ² + 1)) + π/2
    Vaihekulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)
    
Exit Function

ErrorHandler:
    MsgBox "Error in Vaihekulma: " & Err.Description & vbCrLf & _
           "Power Factor: " & Cosfii, vbCritical, "Phase Angle Calculation Error"
    Vaihekulma = 0  ' Virhetilanteessa palautetaan 0
End Function

'------------------------------------------------------------------------------
' Funktio: MotKaapUh
' Tarkoitus: Laskee moottorikäaapelin jännitealenemaprosentin
' Parametrit:
'   Cosfii  - Tehokerroin (cos φ)
'   Resist  - Kaapelin resistanssi (Ω/km)
'   React   - Kaapelin reaktanssi (Ω/km)
'   Virta   - Virta (A)
'   Voltage - Jännite (V)
'   Pituus  - Kaapelin pituus (m)
' Palautusarvo: Muotoiltu jännitealenemaprosentti, esim. "2.35 %"
'------------------------------------------------------------------------------
Function MotKaapUh(Cosfii As Single, Resist As Double, React As Double, Virta As Single, Voltage As Integer, Pituus As Integer)
Dim Kulma As Double
On Error GoTo MotKaapUhErr

' Lasketaan vaihekulma
Kulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)

' Lasketaan jännitealenema: √3 * I * (R*L*cosφ + X*L*sinφ)
MotKaapUh = Sqr(3) * Virta * ((Resist * Pituus * Cosfii) + (React * Pituus * Sin(Kulma)))

' Muunnetaan prosentiksi jännitteestä
MotKaapUh = (MotKaapUh / Voltage) * 100

' Muotoillaan prosenttiluvuksi
MotKaapUh = Format(MotKaapUh, "# ##0.0#") & " %"

Exit_Function:
    Exit Function

MotKaapUhErr:
    MsgBox "Error in MotKaapUh: " & Err.Description & vbCrLf & _
           "cosφ=" & Cosfii & " R=" & Resist & " X=" & React & _
           " I=" & Virta & " V=" & Voltage & " L=" & Pituus, _
           vbCritical, "Cable Voltage Drop Calculation Error"
    MotKaapUh = "Error"
    Resume Exit_Function
End Function

'------------------------------------------------------------------------------
' Funktio: LisaaNo
' Tarkoitus: Lisää luvun merkkijonoon ja täyttää johtavilla nollilla
' Parametrit:
'   Tieto  - Alkuperäinen numeromerkkijono (esim. "001")
'   Lisays - Lisättävä luku (esim. 100)
' Palautusarvo: Padded-merkkijono (esim. "101")
' Esimerkki: LisaaNo("001", 100) = "101"
'------------------------------------------------------------------------------
Function LisaaNo(Tieto As Variant, Lisays As Integer) As String
On Error GoTo ErrorHandler

Dim Pit As Integer    ' Alkuperäinen merkkijonon pituus
Dim No As Integer     ' Numeerinen arvo
Dim i As Integer      ' Silmukkalasku

  ' Käsitellään nolla-syöte
  If IsNull(Tieto) Then
    LisaaNo = ""
  Else
    Pit = Len(Tieto)          ' Alkuperäinen pituus
    No = Val(Tieto)           ' Muunnetaan luvuksi
    No = No + Lisays          ' Lisätään
    LisaaNo = CStr(No)        ' Muunnetaan takaisin merkkijonoksi
    
    ' Täytetään johtavilla nollilla alkuperäisen pituuden säilyttämiseksi
    For i = 0 To Pit - Len(LisaaNo) - 1
        LisaaNo = "0" & LisaaNo
    Next i
  End If
  
Exit Function

ErrorHandler:
    MsgBox "Error in LisaaNo: " & Err.Description & vbCrLf & _
           "Input: " & Tieto & ", Addition: " & Lisays, _
           vbCritical, "Number Addition Error"
    LisaaNo = ""  ' Virhetilanteessa palautetaan tyhjä merkkijono
End Function

'Example usage in query:
'   Field: LisaaNo([FieldName], 100)
