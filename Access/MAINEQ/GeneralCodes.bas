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
' Funktio: IsLoaded
' Tarkoitus: Tarkistaa, onko lomake auki lomake- tai taulukkonäkymässä
' Parametrit:
'   strFormName - Tarkistettavan lomakkeen nimi
' Palautusarvo: True, jos lomake on auki eikä suunnittelunäkymässä
'------------------------------------------------------------------------------
Function IsLoaded(ByVal strFormName As String) As Boolean
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

Dim Rivi As String  ' Viimeinen revisiorivi
Dim Valil As Long  ' Ensimmäisen välilyönnin sijainti
Dim Vinoviiva As Long  ' Päivämäärän päättävän kauttaviivan sijainti

  ' Käytetään historian viimeistä riviä (aiempi toiminta säilytetty)
  Rivi = Revisio
  If InStr(Rivi, vbCrLf) > 0 Then Rivi = Mid$(Rivi, InStrRev(Rivi, vbCrLf) + 2)

  ' Poimitaan päivämääräosa (välilyönnin ja ensimmäisen kauttaviivan väliltä)
  Valil = InStr(Rivi, " ")
  If Valil = 0 Then Exit Function
  Rivi = Mid$(Rivi, Valil + 1)
  Vinoviiva = InStr(Rivi, "/")
  If Vinoviiva = 0 Then
    HaeViimPaiva = Rivi
  Else
    HaeViimPaiva = Left$(Rivi, Vinoviiva - 1)
  End If

Exit Function

ErrorHandler:
    MsgBox "Virhe funktiossa HaeViimPaiva: " & Err.Description, vbCritical, "Revisiopäivän poiminta"
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
Function Optiot(ByVal Drives_ID As Long) As String
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
    MsgBox "Virhe funktiossa Optiot: " & Err.Description & vbCrLf & _
           "Käytön ID: " & Drives_ID, vbCritical, "Optioiden haku"
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
    & "AND (MAINEQ.Department = DRIVES.Department) WHERE MAINEQ.ProjectElement= '" & Replace(LaiteNr, "'", "''") & "';"
    
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
    MsgBox "Virhe funktiossa Positiot: " & Err.Description & vbCrLf & _
           "Projektielementti: " & LaiteNr, vbCritical, "Positioiden haku"
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
Function Vaihekulma(Cosfii As Double) As Double
On Error GoTo ErrorHandler

    ' Lasketaan vaihekulma: arctan(-cosφ / sqrt(-cosφ² + 1)) + π/2.
    ' Rajatapauksessa |cosφ| = 1 neliöjuuri on nolla, joten kulma annetaan suoraan.
    If Cosfii >= 1 Then
        Vaihekulma = 0
    ElseIf Cosfii <= -1 Then
        Vaihekulma = 4 * Atn(1)
    Else
        Vaihekulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)
    End If
    
Exit Function

ErrorHandler:
    MsgBox "Virhe funktiossa Vaihekulma: " & Err.Description & vbCrLf & _
           "Tehokerroin: " & Cosfii, vbCritical, "Vaihekulman laskenta"
    Vaihekulma = 0  ' Virhetilanteessa palautetaan 0
End Function

'------------------------------------------------------------------------------
' Funktio: MotKaapUh
' Tarkoitus: Laskee moottorikaapelin jännitealenemaprosentin
' Parametrit:
'   Cosfii  - Tehokerroin (cos φ)
'   Resist  - Kaapelin resistanssi (Ω/km)
'   React   - Kaapelin reaktanssi (Ω/km)
'   Virta   - Virta (A)
'   Voltage - Jännite (V)
'   Pituus  - Kaapelin pituus (m)
' Palautusarvo: Muotoiltu jännitealenemaprosentti, esim. "2.35 %"
'------------------------------------------------------------------------------
Function MotKaapUh(Cosfii As Double, Resist As Double, React As Double, Virta As Double, Voltage As Long, Pituus As Double)
Dim Kulma As Double  ' Vaihekulma radiaaneina
Dim Alenema As Double  ' Jännitealenema voltteina
On Error GoTo MotKaapUhErr

' Lasketaan vaihekulma (rajatapaus cosφ = 1 käsitellään Vaihekulma-funktiossa)
Kulma = Vaihekulma(Cosfii)

' Lasketaan jännitealenema: √3 * I * (R*L*cosφ + X*L*sinφ)
Alenema = Sqr(3) * Virta * ((Resist * Pituus * Cosfii) + (React * Pituus * Sin(Kulma)))

' Muunnetaan prosentiksi jännitteestä
Alenema = (Alenema / Voltage) * 100

' Muotoillaan prosenttiluvuksi
MotKaapUh = Format(Alenema, "# ##0.0#") & " %"

Exit_Function:
    Exit Function

MotKaapUhErr:
    MsgBox "Virhe funktiossa MotKaapUh: " & Err.Description & vbCrLf & _
           "cosφ=" & Cosfii & " R=" & Resist & " X=" & React & _
           " I=" & Virta & " V=" & Voltage & " L=" & Pituus, _
           vbCritical, "Kaapelin jännitealenema"
    MotKaapUh = "Virhe"
    Resume Exit_Function
End Function

'------------------------------------------------------------------------------
' Funktio: LisaaNo
' Tarkoitus: Lisää luvun merkkijonoon ja täyttää johtavilla nollilla
' Parametrit:
'   Tieto  - Alkuperäinen numeromerkkijono (esim. "001")
'   Lisays - Lisättävä luku (esim. 100)
' Palautusarvo: Johtavilla nollilla täytetty merkkijono (esim. "101")
' Esimerkki: LisaaNo("001", 100) = "101"
'------------------------------------------------------------------------------
Function LisaaNo(Tieto As Variant, Lisays As Long) As String
On Error GoTo ErrorHandler

Dim Pit As Long       ' Alkuperäinen merkkijonon pituus
Dim No As Long        ' Numeerinen arvo
Dim i As Integer      ' Silmukkalasku

  ' Null-syötteestä palautetaan tyhjä merkkijono
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
    MsgBox "Virhe funktiossa LisaaNo: " & Err.Description & vbCrLf & _
           "Syöte: " & Tieto & ", lisäys: " & Lisays, _
           vbCritical, "Luvun lisäys"
    LisaaNo = ""  ' Virhetilanteessa palautetaan tyhjä merkkijono
End Function

' Esimerkki kyselyssä:
'   Kenttä: LisaaNo([Kentännimi], 100)
