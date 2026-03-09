Option Compare Database
Option Explicit
Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' --------- [ CHOOSE FILE ] -----------------
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
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
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Type BrowseInfo
    hOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type
Public CDialogPath As String

Public Function ValitseHakem(Handle As Long, Optional StartPath As String) As String
  Dim lpIDList As Long
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
    CDialogPath = "C:\"
  Else
    If Right(StartPath, 2) = ":\" Then
      CDialogPath = StartPath
    ElseIf Right(StartPath, 1) = "\" Then
      CDialogPath = Left(StartPath, Len(StartPath) - 1)
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
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
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
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function
Public Function ValitseTiedosto(Nimi As String, Otsikko As String) As String
'T�m� valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Polku As String
    Dim WHandle As Long
    
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
        'Painettiin Cancel painiketta. Ei tehd� mit��n
    Else 'Otetaan yl�s Tiedostonimi ja Hakemisto
       ValitseTiedosto = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
    End If


End Function


