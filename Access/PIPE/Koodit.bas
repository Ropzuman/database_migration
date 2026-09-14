Option Compare Database
Option Explicit

'================================================================================
' Moduuli: Koodit
' Tarkoitus: Ydinfunktiot ja AutoCAD-integraatio PIPE-tietokannalle
' P‰ivitetty: 2025-11-12 - VBA7/64-bit-tuki lis‰tty
'
' Kuvaus:
'   Tarjoaa PIPE-tietokannan keskeiset toiminnot:
'   - AutoCAD-dokumentti-integraatio (blokkien korostus ja zoomi)
'   - K‰ytt‰j‰n kirjautumisseuranta (verkkonimet, tietokoneen nimi)
'   - Blokkien avaaminen tietokantatietueista (venttiilit, putkilinjat)
'   - Merkkijonon j‰sennysfunktiot
'
' Riippuvuudet:
'   - AutoCAD Application (COM-automaatio)
'   - UsysUsers-taulu (kirjautumisseuranta)
'   - MANUALVALVES, PIPELINES, PIPELINEDATA, MANVALVEDATA-taulut
'   - advapi32.dll (GetUserName-API)
'   - kernel32.dll (GetComputerName-API)
'================================================================================


'--------------------------------------------------------------------------------
' Windows API -m‰‰rittelyt
'--------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

Private Const AC_MODEL_SPACE As Long = 0  ' AutoCADin acModelSpace-arvo late bindingia varten
Private Const AC_MAX_WINDOW As Long = 3   ' AutoCADin acMax-arvo late bindingia varten
Private Const AC_ZOOM_SCALED_RELATIVE As Long = 1  ' AutoCADin acZoomScaledRelative-arvo

'--------------------------------------------------------------------------------
' Aliohjelma: AvaaBlock
' Tarkoitus: Avaa ja korostaa AutoCAD-blokki aktiivisesta tietokantatalukosta
' Toiminta:
'   1. M‰‰rit‰‰n, kutsutaanko MANUALVALVES- vai PIPELINES-taulusta
'   2. Lukee blokin tiedot tietokannasta (polku, piirustus, k‰sittelyviite)
'   3. Kutsuu AvaaKuvasta avatakseen piirustuksen ja zoomaakseen blokkiin
' Huom:
'   - Toimii MANUALVALVES- ja PIPELINES-tietotaulun n‰kymiss‰
'   - Siirt‰‰ PIPELINES-linjoille frmOpenPIPELINE-lomakkeelle
'   - Vaatii k‰ynniss‰ olevan AutoCAD-instanssin
'--------------------------------------------------------------------------------
Sub AvaaBlock()
On Error GoTo ErrorHandler
    Dim DWG As String       ' Piirustuksen tiedostonimi
    Dim Polku As String     ' Polku piirustuskansioon
    Dim Handle As String    ' AutoCAD-blokin k‰sittelyviite
    Dim Info As String      ' Blokin tunnistusmerkkijono
    Dim i As Integer        ' Silmukkalasku
    Dim Tyyppi As Integer   ' Tyyppi: 1=k‰siventtiili, 0=putkilinja
    Dim Taulu As DAO.Recordset  ' Tietokantakyselyn tulokset
    Dim Alue As String      ' Aluekoodi
    Dim Linja As String     ' Linjanumero

    ' Haetaan suhteellinen polku virtauskaaviokansioon
    Polku = Left$(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name))) & "..\..\R\FlowSheets\"

    ' M‰‰ritet‰‰n l‰hdetaulu
    If UCase$(Application.CurrentObjectName) = "MANUALVALVES" Or UCase$(Application.CurrentObjectName) = "PIPELINES" Then
      If UCase$(Application.CurrentObjectName) = "MANUALVALVES" Then Tyyppi = 1
      If Tyyppi = 1 Then
        ' Haetaan k‰siventtiilin blokkitiedot
        Set Taulu = CurrentDb.OpenRecordset("SELECT * FROM MANVALVEDATA WHERE AREACODE = '" & Screen.ActiveDatasheet("Area").Value & "' AND VAL_NO = '" & Screen.ActiveDatasheet("ValveNo").Value & "'")
        DWG = LCase$(Taulu.Fields("ImpFileID"))
        Handle = Taulu.Fields("Handles")
        Info = Taulu.Fields("AREACODE") & "-" & Taulu.Fields("VAL_NO")
      Else
        ' Putkilinjoille avataan valintaikkuna
        Alue = Nz(Screen.ActiveDatasheet("Area").Value)
        Linja = Nz(Screen.ActiveDatasheet("LineNo").Value)
        DoCmd.OpenForm "frmOpenPIPELINE", acNormal, , , , , Alue & "," & Linja
        Exit Sub
      End If
      If Polku = "" Then
        MsgBox "Ei ole tietoa miss‰ kuvassa kohde " & Info & " on !", vbCritical, "Etsi kohde"  ' "No information where object is!"
        Exit Sub
      End If
      AvaaKuvasta Polku, DWG, Handle, Info
    ElseIf UCase$(Application.CurrentObjectName) = "PIPELINEDATA" Or UCase$(Application.CurrentObjectName) = "MANVALVEDATA" Then
        ' Avataan taulukkon‰kym‰st‰
        Polku = Screen.ActiveDatasheet("PATH").Value
        If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
        
        DWG = LCase$(Screen.ActiveDatasheet("ImpFileID").Value)
        Handle = Screen.ActiveDatasheet("Handles").Value
        If UCase$(Application.CurrentObjectName) = "PIPELINEDATA" Then
          Info = Screen.ActiveDatasheet("DEP").Value & "-" & Screen.ActiveDatasheet("LINENO").Value
        Else
          Info = Screen.ActiveDatasheet("AREACODE").Value & "-" & Screen.ActiveDatasheet("VALVE_NO").Value
        End If
        AvaaKuvasta Polku, DWG, Handle, Info
    Else
      MsgBox "MANUALVALVES tai PIPELINES taulukon tulee avaoinna n‰ytˆll‰!", vbCritical, "Etsi kohde"  ' "Table must be open!"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Virhe blokin avaamisessa: " & Err.Description, vbCritical, "AvaaBlock"
End Sub
'--------------------------------------------------------------------------------
' Aliohjelma: AvaaKuvasta
' Tarkoitus: Avaa AutoCAD-piirustus ja zoomaa/korostaa m‰‰ritetty blokki
' Parametrit:
'   Polku - Hakemistopolku piirustustiedostoon
'   Nimi  - Piirustuksen tiedostonimi (.dwg)
'   Handle - AutoCAD-blokin k‰sittelyviite (hex-merkkijono)
'   Info  - Blokin tunnistus virheilmoituksia varten
' Toiminta:
'   1. Yhdistet‰‰n k‰ynniss‰ olevaan AutoCAD-instanssiin
'   2. Tarkistetaan, onko piirustus jo auki; tarvittaessa avataan
'   3. Etsit‰‰n blokki k‰sittelyviitteen perusteella
'   4. Zoomataan blokkiin ja korostetaan se
'   5. Aktivoidaan AutoCAD-ikkuna
'--------------------------------------------------------------------------------
Public Sub AvaaKuvasta(Polku As String, Nimi As String, Handle As String, Info As String)
On Error GoTo ErrorHandler
    Dim oACAD As Object      ' AcadApplication - late binding (64-bit)
    Dim Entity As Object    ' AcadEntity - late binding (64-bit)
    Dim MinPoint As Variant ' Rajaruudun minimipiste
    Dim MaxPoint As Variant ' Rajaruudun maksimipiste
    Dim i As Integer        ' Silmukkalasku
    Dim OK As Boolean       ' Piirustus avattu -lippu
    Dim LowerNimi As String ' Piirustuksen nimi pieniksi muutettuna

    ' Normalisoidaan piirustuksen nimi yhtenev‰‰ vertailua varten
    LowerNimi = LCase$(Nimi)

    ' Yhdistet‰‰n k‰ynniss‰ olevaan AutoCADiin
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    If Err <> 0 Then
      MsgBox "K‰ynniss‰ olevaa AutoCADi‰ ei lˆytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Etsi Kohde"  ' "Running AutoCAD not found!"
      Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Tarkistetaan, onko piirustus jo auki (v‰limuistissa oleva nimi)
    OK = False
    For i = 0 To oACAD.Documents.Count - 1
      If LCase$(oACAD.Documents(i).Name) = LowerNimi Then
        oACAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    
    ' Avataan piirustus, jos ei jo auki
    If Not OK Then
      On Error Resume Next
      oACAD.Documents.Open Polku & Nimi
      If Err = 0 Then
        OK = True
      Else
        MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Nimi, vbCritical, "Etsi kohde"  ' "Error opening document"
        Err.Clear
      End If
      On Error GoTo ErrorHandler
    End If
    
    ' Etsit‰‰n ja korostetaan blokki, jos piirustus avautui onnistuneesti
    If OK Then
      If Handle = "" Then
        MsgBox "Kohteen  " & Info & " sijainti ei ole tiedossa, vain kuva avattiin.", vbCritical, "Etsi kohde"  ' "Block location unknown"
      Else
        oACAD.ActiveDocument.ActiveSpace = AC_MODEL_SPACE
        On Error Resume Next
        Set Entity = oACAD.ActiveDocument.HandleToObject(Handle)
        If Err <> 0 Then
          MsgBox "Kuvasta ei lˆytynyt kohdetta tietokannan tiedoilla (Handle oli v‰‰r‰)!", vbCritical, "Etsi kohde"  ' "Block not found (wrong handle)"
          Err.Clear
          On Error GoTo ErrorHandler
        Else
          On Error GoTo ErrorHandler
          ' Zoomataan blokkiin ja korostetaan se
          Entity.GetBoundingBox MinPoint, MaxPoint
          oACAD.ActiveDocument.WindowState = AC_MAX_WINDOW
          oACAD.ZoomWindow MinPoint, MaxPoint
          oACAD.ZoomScaled 0.3, AC_ZOOM_SCALED_RELATIVE
          Entity.Highlight True
        End If
      End If
      AppActivate oACAD.Caption, True
    End If
    
    ' Siivotaan objektit
    Set Entity = Nothing
    Set oACAD = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Set Entity = Nothing
    Set oACAD = Nothing
    On Error GoTo 0
    MsgBox "Virhe AvaaKuvasta-rutiinissa: " & Err.Description, vbCritical, "Virhe"
End Sub
'--------------------------------------------------------------------------------
' Funktio: NetworkUserName
' Tarkoitus: Hakee Windowsin verkkok‰ytt‰j‰nimen
' Palautusarvo: Merkkijono - verkkok‰ytt‰j‰nimi tai "Tuntematon"
' Huom: K‰ytet‰‰n DOCUMENTS-kannan lomakkeissa (Form_USysReserve, Form_USysAddDocument)
'--------------------------------------------------------------------------------
Public Function NetworkUserName() As String
    ' Haetaan Windows-verkkok‰ytt‰j‰nimi (k‰ytet‰‰n DOCUMENTS-kannan lomakkeissa)
    Dim BuffSize As Long
    Dim NBuffer As String
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetUserName(NBuffer, BuffSize) Then
        NetworkUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        NetworkUserName = "Tuntematon"
    End If
End Function

'--------------------------------------------------------------------------------
' Funktio: SetStartup
' Tarkoitus: Kirjaa k‰ytt‰j‰n kirjautumistiedot UsysUsers-tauluun
' Toiminta:
'   1. Hakee verkkok‰ytt‰j‰nimen Windows-API:lla
'   2. Hakee tietokoneen nimen Windows-API:lla
'   3. Hakee Access-tietokannan k‰ytt‰j‰nimen
'   4. Kirjoittaa tiedot UsysUsers-tauluun aikaleimaneen
' Huom:
'   - Kutsutaan AutoExec-makrosta tai k‰ynnistyslomakkeesta
'   - Virheet k‰sitell‰‰n hiljaisesti ó ei keskeytet‰ sovelluksen k‰ynnistyst‰
'--------------------------------------------------------------------------------
Function SetStartup()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database      ' Tietokantaviittaus
    Dim Taulu As DAO.Recordset  ' UsysUsers-taulun tietue
    Dim NWUserName As String    ' Verkkok‰ytt‰j‰nimi Windowsista
    Dim CName As String         ' Tietokoneen nimi Windowsista
    Dim BuffSize As Long        ' Puskurin koko API-kutsuille
    Dim NBuffer As String       ' Merkkijonopuskuri API-kutsuille

    ' Haetaan verkkok‰ytt‰j‰nimi Windows API:n avulla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Tuntematon"
    End If

    ' Haetaan tietokoneen nimi Windows API:n avulla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Tuntematon"
    End If

    ' Kirjoitetaan kirjautumistietue UsysUsers-tauluun
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Verkkok‰ytt‰j‰nimi
        .Fields(1) = CurrentUser()  ' Access-k‰ytt‰j‰nimi
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
    ' Hiljainen virheenk‰sittely ó ei keskeytet‰ sovelluksen k‰ynnistyst‰
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' Funktio: POIMI
' Tarkoitus: Poimii osan viivalla erotetusta merkkijonosta
' Parametrit:
'   Tieto - J‰sennett‰v‰ merkkijono (muoto: "osa1-osa2-osa3")
'   osa   - Poimittavan osan numero (1-pohjainen indeksi)
' Palautusarvo: Variant - Poimittu osa tai Null, jos syˆte on tyhj‰/Null
' Esimerkki: POIMI("ALUE-123-VENTTIILI", 2) palauttaa "123"
'--------------------------------------------------------------------------------
Function POIMI(Tieto As Variant, osa As Integer) As Variant
On Error GoTo ErrorHandler
    Dim Osat As Variant  ' Merkkijonon osat taulukkona

    If IsNull(Tieto) Or Tieto = "" Then
       POIMI = Null
    Else
      Osat = Split(Tieto, "-")
      POIMI = Osat(osa - 1)  ' Muunnetaan 1-pohjaisesta 0-pohjaiseen indeksiin
    End If
    Exit Function

ErrorHandler:
    POIMI = Null
End Function
