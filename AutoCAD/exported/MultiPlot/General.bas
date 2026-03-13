Attribute VB_Name = "General"
' --- 64-bitti korjattu: PtrSafe lisatty kaikkiin Declare-lauseisiin ---
Public FSO As New FileSystemObject
Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' --------- [ CHOOSE FILE ] -----------------
Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As LongPtr
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
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
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As LongPtr)
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Type BrowseInfo
    hOwner      As LongPtr
    pIDLRoot       As LongPtr
    pszDisplayName As LongPtr
    lpszTitle      As LongPtr
    ulFlags        As Long
    lpfn           As LongPtr
    lParam         As LongPtr
    iImage         As Long
End Type
Public SPaikka As String
Public SSuunta As String
Public SYlosalais As Boolean
Public SXOffset As Double
Public SYOffset As Double
Public SKork As Double
Public SFName As Boolean
Public SDate As Boolean
Public Tiedostot() As String

Public Function ValitseHakem(Handle As LongPtr, Optional StartPath As String) As String
  Dim lpIDList As LongPtr
  Dim lpSelPath As Long
  Dim ThePath As String
  Dim Otsikko As String
  Dim tBrowseInfo As BrowseInfo
  Const LMEM_FIXED = &H0
  Const LMEM_ZEROINIT = &H40
  Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
  Const BIF_RETURNONLYFSDIRS = 1
  Const BIF_DONTGOBELOWDOMAIN = 2
  Const MAX_PATH = 260

  If IsMissing(StartPath) Or StartPath = "" Then
    CDialogPath = "K:\PROJECTS"
  Else
    If Right(StartPath, 2) = ":\" Then
      CDialogPath = StartPath
    ElseIf Right(StartPath, 1) = "\" Then
      CDialogPath = Left(StartPath, Len(StartPath) - 1)
    Else
      CDialogPath = StartPath
    End If
  End If
     
     Otsikko = "Choose path:"
     With tBrowseInfo
        .hOwner = Handle
        .pIDLRoot = 0
        .lpszTitle = lstrcat(Otsikko, "")
        .lpfn = DummyFunc(AddressOf BrowseCallbackProc)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        ThePath = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, ThePath
        ThePath = Left(ThePath, InStr(ThePath, vbNullChar) - 1)
    Else
      ThePath = ""
    End If
    CoTaskMemFree lpIDList
    If ThePath <> "" Then
      If Right(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
    End If
    ValitseHakem = ThePath
End Function
Public Function BrowseCallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
  Const BFFM_INITIALIZED = 1
  Const BFFM_SETSELECTION = &H466
  Dim retval As Long         'Return value
  On Error Resume Next
  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, True, ByVal CDialogPath)
  Case Else
  End Select
  BrowseCallbackProc = 0
  Err.Clear
End Function
Public Sub MPlot()
Dim i As Integer
Dim Nimi As String
  'Näytetään tulostusformi
  Formi.Show
End Sub
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function
Public Function KeraaTiedostot(Kaikki As Boolean) As Boolean
  ReDim Tiedostot(0)
  Ajastin = Timer
  If Formi.FPath.Value = "" Or Formi.FName.Value = "" Then
    MsgBox "Please choose path and file(s) first.", vbCritical, "Plot"
    Keraatiedosto = False
    Exit Function
  End If
  If Right(Formi.FPath.Value, 1) <> "\" Then Formi.FPath.Value = Formi.FPath.Value & "\"
  LueTiedosto Formi.FPath.Value & Formi.FName.Value, Kaikki
  KeraaTiedostot = True
End Function

Private Sub LueTiedosto(Tiedosto As String, Kaikki As Boolean)
Dim oTiedosto As Scripting.TextStream
Dim Rivi As String
Dim i As Integer
Dim Alku As String
Dim Polku As String
Dim Tied As String
  Polku = Left(Tiedosto, InStrRev(Tiedosto, "\"))
  If UCase(Right(Tiedosto, 3)) = "LST" Then 'Jos kysymyksessä on lista
    Tied = Dir(Tiedosto)
    Do While Tied <> ""
      Set oTiedosto = FSO.OpenTextFile(Polku & Tied, 1) 'Avataan lista
      Do While Not oTiedosto.AtEndOfStream 'Käydään läpi listan kaikki rivit
        DoEvents
        Rivi = oTiedosto.ReadLine
        Alku = Mid(Rivi, 1, 1)
        If Alku = " " Or Alku = ";" Or Rivi = "" Then 'Rivi ei sisällä tietoa tai se on kommentti
          'Ei mitään
        ElseIf Alku = "@" Then 'Rivi on linkki toiseen listaan
          LueTiedosto Mid(Rivi, 2), Kaikki 'Luetaan toisen lista sisältö
        Else
          If InStr(Rivi, "\") Then 'Rivi on suora osoitus tiedostoon polkuinen
            Tiedosto = Rivi
          ElseIf Len(Rivi) > 38 Then 'Rivi on DOS listaus ja tiedostonimi alkaa merkistä 38
            Tiedosto = Formi.FPath.Value & Mid(Rivi, 38)
          Else 'Listassa on vain tiedoston nimi, joten polku tulee formissa olevasta polusta
            Tiedosto = Formi.FPath.Value & Rivi
          End If
          If Right(UCase(Tiedosto), 4) <> ".DWG" And Right(UCase(Tiedosto), 4) <> ".DXF" Then Tiedosto = Tiedosto & ".DWG"
          LisaaListaan Tiedosto
          If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
            Exit Do
          End If
        End If
        If Kaikki = False Then Exit Sub
      Loop
      Set oTiedosto = Nothing
      Tied = Dir
    Loop
  Else
    Tied = Dir(Tiedosto)
    Do While Tied <> ""
      LisaaListaan Polku & Tied
      Tied = Dir
      If Kaikki = False Then Exit Sub
    Loop
  End If
End Sub
Private Sub LisaaListaan(Tiedosto As String)
  ReDim Preserve Tiedostot(UBound(Tiedostot) + 1)
  Tiedostot(UBound(Tiedostot)) = Tiedosto
End Sub
