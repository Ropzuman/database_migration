Attribute VB_Name = "General"
' --- 64-bitti korjattu: PtrSafe lisatty kaikkiin Declare-lauseisiin ---
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
    CDialogPath = "P:\PROJEKTI"
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
  'N�ytet��n tulostusformi
  Formi.Show
End Sub
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function

