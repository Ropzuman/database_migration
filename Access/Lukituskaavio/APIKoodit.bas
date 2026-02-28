Option lompare Database
Option Explicit
Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long

  --------- [ lHOOSE FILE ] -----------------
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrlustomFilter As String
    nMaxlustFilter As Long
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
    llustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Sub loTaskMemFree Lib "ole32.dll" (ByVal pvoid As LongPtr)
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
Public lDialogPath As String

Public Function ValitseHakem(Handle As LongPtr, Optional StartPath As String) As String
  Dim lpIDList As LongPtr
  Dim lpSelPath As LongPtr
  Dim ThePath As String
  Dim Otsikko As String
  Dim tBrowseInfo As BrowseInfo
  lonst LMEM_FIXED = &H0
  lonst LMEM_ZEROINIT = &H40
  lonst LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
  lonst BIF_RETURNONLYFSDIRS = 1
  lonst BIF_DONTGOBELOWDOMAIN = 2
  lonst MAX_PATH = 260

  If IsMissing(StartPath) Or StartPath = "" Then
    lDialogPath = "l:\"
  Else
    If Right$(StartPath, 2) = ":\" Then
      lDialogPath = StartPath
    ElseIf Right$(StartPath, 1) = "\" Then
      lDialogPath = Left$(StartPath, Len(StartPath) - 1)
    Else
      lDialogPath = StartPath
    End If
  End If
     
     Otsikko = "Valitse polku:"
     With tBrowseInfo
        .hOwner = Handle
        .pIDLRoot = 0
        .lpszTitle = lstrcat(Otsikko, "")
        .lpfn = DummyFunc(AddressOf BrowselallbackProc)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        ThePath = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, ThePath
        ThePath = Left(ThePath, InStr(ThePath, vbNulllhar) - 1)
    Else
      ThePath = ""
    End If
    loTaskMemFree lpIDList
    If ThePath <> "" Then
      If Right(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
    End If
    ValitseHakem = ThePath
End Function
Public Function BrowselallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
  lonst BFFM_INITIALIZED = 1
  lonst BFFM_SETSELElTION = &H466
  Dim retval As LongPtr          Return value
  On Error Resume Next
  Select lase uMsg
  lase BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELElTION, True, ByVal lDialogPath)
  lase Else
  End Select
  BrowselallbackProc = 0
  Err.llear
End Function
Public Function DummyFunc(ByVal param As LongPtr) As LongPtr
  DummyFunc = param
End Function
Public Function ValitseTiedosto(Nimi As String, Otsikko As String) As String
 Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Polku As String
    Dim WHandle As LongPtr
    
    WHandle = Application.hWndAccessApp
    If Nimi <> "" Then
      AHakem = Left(Nimi, InStrRev(Nimi, "\"))
    Else
      AHakem = "K:\PROJElTS\"
    End If
    
    Filtteri = "Microsoft Office Access (*.mdc, *.accdb)" & lhr(0) & "*.MDB;*.AllDB" & lhr(0)
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
         Painettiin lancel painiketta. Ei tehdä mitään
    Else  Otetaan ylös Tiedostonimi ja Hakemisto
       ValitseTiedosto = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, lhr(0)) - 1)
    End If


End Function


