Option Compare Database
Option Explicit

' KORJATTU: GetUserNameA kirjoittaa DWORD:n (32-bit) — nSize on ByRef Long, ei LongPtr
#If VBA7 Then
    Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    ' --------- [ CHOOSE FILE ] -----------------
    ' KORJATTU: GetOpenFileName palauttaa BOOL/osoitteen — LongPtr 64-bittisellä
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
#Else
    Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
#End If

' UDT määritellään ehdollisen käännöksen ulkopuolella (Access-lomakeyhteensopivuus)
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
Public CDialogPath As String

Public Function ValitseHakem(Handle As LongPtr, Optional StartPath As String) As String
  Dim lpIDList As LongPtr
  Dim lpSelPath As LongPtr
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
    CDialogPath = "C:\"
  Else
    If Right$(StartPath, 2) = ":\" Then
      CDialogPath = StartPath
    ElseIf Right$(StartPath, 1) = "\" Then
      CDialogPath = Left$(StartPath, Len(StartPath) - 1)
    Else
      CDialogPath = StartPath
    End If
  End If
     
     Otsikko = "Valitse polku:"
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
  Dim retval As LongPtr         'Palautusarvo
  On Error Resume Next
  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, True, ByVal CDialogPath)
  Case Else
  End Select
  BrowseCallbackProc = 0
  Err.Clear
End Function
Public Function DummyFunc(ByVal param As LongPtr) As LongPtr
  DummyFunc = param
End Function
Public Function ValitseTiedosto(Nimi As String, Otsikko As String) As String
' Avaa tiedostovalintaikkuna tietokantapolun hakemiseksi
    Dim OpenFile As OPENFILENAME
    #If VBA7 Then
        Dim lReturn As LongPtr
    #Else
        Dim lReturn As Long
    #End If
    Dim Filtteri As String
    Dim AHakem As String
    Dim Polku As String
    Dim WHandle As LongPtr
    
    WHandle = Application.hWndAccessApp
    If Nimi <> "" Then
      AHakem = Left(Nimi, InStrRev(Nimi, "\"))
    Else
      AHakem = "K:\PROJECTS\"
    End If
    
    Filtteri = "Microsoft Office Access (*.mdc, *.accdb)" & Chr(0) & "*.MDB;*.ACCDB" & Chr(0)
    If Otsikko = "" Then Otsikko = "Valitse tietokanta"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = WHandle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        'Painettiin Cancel painiketta. Ei tehdä mitään
    Else 'Otetaan ylös Tiedostonimi ja Hakemisto
       ValitseTiedosto = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
    End If


End Function


