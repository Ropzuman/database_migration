Option Compare Database
Option Explicit

' OPENFILENAME-tyyppi täytyy esitellä ehtokääntämislohkon ulkopuolella Access-lomakkeiden yhteensopivuuden vuoksi
Public Type OPENFILENAME
    lStructSize As Long
#If VBA7 Then
    hwndOwner As LongPtr
    hInstance As LongPtr
#Else
    hwndOwner As Long
    hInstance As Long
#End If
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
#If VBA7 Then
    lCustData As LongPtr
    lpfnHook As LongPtr
#Else
    lCustData As Long
    lpfnHook As Long
#End If
    lpTemplateName As String
End Type

' KORJATTU: Muutettu "Private Declare" -> "Public Declare"
' (nSize: LongPtr → ByRef Long — Win32 DWORD on 32-bittinen, ei osoitinkokoinen)
#If VBA7 Then
    Public Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
#Else
    Public Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
#End If

' Moduulitason tilamuuttujat viimeisimmälle hakuehdolle
Private m_last_criteria As Variant
Private m_last_used As Variant

' Tallentaa viimeksi käytetyn arvon ja hakuehdon
Function Set_last(Values As Variant, criterias As Variant) As Variant
    m_last_criteria = criterias
    m_last_used = Values
    Set_last = m_last_used
End Function

' Palauttaa viimeksi käytetyn arvon
Function Show_last(criterias As Variant) As Variant
    Show_last = m_last_used
End Function

' Palauttaa viimeksi käytetyn hakuehdon
Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = m_last_criteria
End Function

'------------------------------------------------------------------------------
' Funktio: SniffUser
' Tarkoitus: Kirjaa käyttäjän kirjautumistiedot UsysUsers-tauluun käynnistyksen yhteydessä
' Parametrit: -
' Palautusarvo: -
' Huom: Kutsutaan AutoExec-makrosta. Virheet käsitellään hiljaisesti,
'       jottei kirjautumiskirjaus keskeytä sovelluksen avautumista.
'------------------------------------------------------------------------------
Function SniffUser()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database      ' Tietokantaviittaus
    Dim Taulu As DAO.Recordset  ' UsysUsers-taulun tietue
    Dim NWUserName As String    ' Verkkokäyttäjänimi Windowsista
    Dim CName As String         ' Tietokoneen nimi
    Dim BuffSize As Long        ' Puskurin koko API-kutsulle
    Dim NBuffer As String       ' Merkkijonopuskuri API-kutsulle

    ' Haetaan verkkokäyttäjänimi wu_GetUserName-API:lla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If wu_GetUserName(NBuffer, BuffSize) Then
        NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        NWUserName = "Tuntematon"
    End If

    ' Haetaan tietokoneen nimi ympäristömuuttujasta
    CName = Environ("COMPUTERNAME")
    If CName = "" Then CName = "Tuntematon"

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

